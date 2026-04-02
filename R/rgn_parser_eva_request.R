  # ============================================================================
  # EVA REQUEST — TAXONOMIC HARMONIZATION (Euro+Med / POWO)
  # Workflow after Grenié et al. (2023)
  # =============================================================================

  # ── SETUP ─────────────────────────────────────────────────────────────────────
  # kewr is not on CRAN: devtools::install_github("barnabywalker/kewr")

library(rgnparser)
library(tidyverse)
library(readxl)
library(writexl)
library(stringdist)
library(kewr)

# ── 1. LOAD & CLEAN ──────────────────────────────────────────────────────────

df_red_list <- read_excel("data/red_list_plants_germany.xlsx")
df_red_list_cleaned <- df_red_list |>
  filter(red_list_category %in% c(1, 2, 3, "R")) |>
  rename(name_full = name)

df_forest_species <- read_excel("data/forest_plants_germany.xls")
df_forest_species_cleaned <- df_forest_species |>
  pivot_longer(cols = NT:D, names_to = "distribution", values_to = "value") |>
  tidyr::extract(value, into = c("habitus", "code"), regex = "([BKO])(.*)") |>
  filter(!str_detect(code, "2\\.1|2\\.2")) |>
  mutate(code = case_when(
    code == "1.1" ~ "closed_forest",
    code == "1.2" ~ "edge_disturbance",
    TRUE ~ code
  )) |>
  filter(str_detect(habitus, "K")) |>
  rename(name_full = `Wissenschaftlicher Name`)

# ── 2. PARSE NAMES WITH RGNPARSER ────────────────────────────────────────────

#Per Grenié et al. (2023): decompose name string, store author separately,
# use only "Genus species" (canonicalsimple) for downstream matching.

parse_names <- function(names) {
  gn_parse_tidy(names) |>
    transmute(
      name_full      = verbatim,
      name_binomial  = canonicalsimple,   # "Genus species" — used for matching
      name_author    = authorship,         # stored separately per Grenié et al.
      parse_quality  = quality,
      cardinality    = cardinality
    )
}

red_list_parsed  <- parse_names(df_red_list_cleaned$name_full)
forest_parsed    <- parse_names(df_forest_species_cleaned$name_full)


# ── 3. FLAG VALID BINOMIALS, INTRASPECIFIC TAXA & EXCLUDED NAMES ──────────────
# Cardinality == 2  → clean "Genus species"
# Cardinality == 3  → intraspecific (subsp./var.) — handled separately below
# Quality <= 2      → acceptable parse (1 = perfect, 2 = minor issues)
# Exclusion pattern extended with agg., cf., × (hybrids) per Grenié et al.

EXCLUDE_PATTERN <- regex(
  "\\bsp\\.\\b|\\bspp\\.\\b|indet|sect\\.|gruppe|\\bcf\\.\\b|\\bagg\\.\\b|×",
  ignore_case = TRUE
)

classify_names <- function(parsed_df) {
  parsed_df |>
    mutate(
      name_status = case_when(
        str_detect(name_full, EXCLUDE_PATTERN)          ~ "excluded_keyword",
        parse_quality > 2                               ~ "excluded_quality",
        cardinality == 1                                ~ "excluded_genus_only",
        cardinality == 2 & parse_quality <= 2           ~ "valid_binomial",
        cardinality == 3 & parse_quality <= 2           ~ "intraspecific",
        TRUE                                            ~ "excluded_other"
      ),
      valid = name_status == "valid_binomial"
    )
}

red_list_parsed <- classify_names(red_list_parsed)
forest_parsed   <- classify_names(forest_parsed)

# ── 3a. INTRASPECIFIC HANDLING ────────────────────────────────────────────────
# Per workflow iii.3: try exact POWO match first; if not possible collapse to
# species level and document. Collapse is done here so these names enter the
# POWO query at step 5 alongside valid binomials.
# The collapse_flag column allows you to track these in sensitivity analysis.

handle_intraspecific <- function(parsed_df) {
  parsed_df |>
    mutate(
      # For trinomials extract only "Genus species" as the fallback
      name_binomial_collapsed = if_else(
        name_status == "intraspecific",
        word(name_binomial, 1, 2),
        name_binomial
      ),
      collapse_flag = name_status == "intraspecific"
    )
}

red_list_parsed <- handle_intraspecific(red_list_parsed)
forest_parsed   <- handle_intraspecific(forest_parsed)

# Inspect what is excluded or collapsed — review before proceeding
message("── Red list: excluded / collapsed names ──")
red_list_parsed |>
  filter(name_status != "valid_binomial") |>
  select(name_full, name_status, parse_quality, cardinality)

message("── Forest list: excluded / collapsed names ──")
forest_parsed |>
  filter(name_status != "valid_binomial") |>
  select(name_full, name_status, parse_quality, cardinality)

# ── 4. ATTACH PARSED NAMES BACK TO CLEANED DATA FRAMES ───────────────────────
# Use name_binomial_collapsed as the working matching key (= name_binomial for
# valid binomials, = collapsed species name for intraspecific taxa).

parsed_cols <- c(
  "name_binomial", "name_binomial_collapsed", "name_author",
  "parse_quality", "cardinality", "name_status", "valid",
  "collapse_flag", "taxon_std"
)

df_red_list_cleaned <- df_red_list_cleaned |>
  select(-any_of(parsed_cols)) |>
  left_join(red_list_parsed, by = "name_full")

df_red_list_cleaned <- df_red_list_cleaned |>
  mutate(
    taxon_std = case_when(
      valid         ~ name_binomial,
      collapse_flag ~ name_binomial_collapsed,  # intraspecific collapsed
      TRUE          ~ NA_character_
    )
  )

df_forest_species_cleaned <- df_forest_species_cleaned |>
  select(-any_of(parsed_cols)) |>
  left_join(forest_parsed, by = "name_full")

df_forest_species_cleaned <- df_forest_species_cleaned |>
  mutate(
    taxon_std = case_when(
      valid         ~ name_binomial,
      collapse_flag ~ name_binomial_collapsed,
      TRUE          ~ NA_character_
    )
  )

# Verify joins — if either is empty the name_full keys are mismatched
stopifnot(
  "Red list join produced no valid rows"    = any(!is.na(df_red_list_cleaned$taxon_std)),
  "Forest list join produced no valid rows" = any(!is.na(df_forest_species_cleaned$taxon_std))
)

# ── 5. POWO QUERY (BEFORE JOIN & FUZZY MATCHING) ──────────────────────────────
# Per workflow ii: match against Euro+Med via POWO *before* the final join so
# synonym resolution propagates into taxon_accepted, which is then the join key.
# Per workflow ii.3: record API access date and POWO version note.

query_powo_safe <- function(name, wait = 0.3) {
  Sys.sleep(wait)
  tryCatch({
    result <- search_powo(name, filters = c("species", "accepted"))
    tidy(result) |>
      mutate(query_name = name)
  }, error = function(e) {
    tibble(
      query_name = name,
      name       = NA_character_,
      accepted   = NA,
      fqId       = NA_character_,
      match_note = as.character(e)
    )
  })
}

# Collect all unique names that need querying (valid binomials + collapsed
# intraspecific names from both lists)
names_to_query <- unique(c(
  df_red_list_cleaned    |> filter(!is.na(taxon_std)) |> pull(taxon_std),
  df_forest_species_cleaned |> filter(!is.na(taxon_std)) |> pull(taxon_std)
))

message(sprintf("Querying POWO for %d unique names …", length(names_to_query)))
powo_results_raw <- purrr::map_dfr(names_to_query, query_powo_safe)

# ── 5a. HOMONYM DETECTION ─────────────────────────────────────────────────────
# Per workflow iii.2: flag cases where POWO returns > 1 accepted hit.
# These require manual disambiguation using the EVA distribution header.

powo_results_summary <- powo_results_raw |>
  filter(!is.na(name)) |>
  group_by(query_name) |>
  mutate(n_powo_hits = n()) |>
  ungroup()

homonym_flags <- powo_results_summary |>
  filter(n_powo_hits > 1) |>
  select(query_name, n_powo_hits, name) |>
  distinct()

if (nrow(homonym_flags) > 0) {
  message(sprintf(
    "⚠ %d name(s) returned multiple POWO hits — homonym check required (see homonym_flags):",
    n_distinct(homonym_flags$query_name)
  ))
  print(homonym_flags)
} else {
  message("✓ No homonyms detected.")
}

# ── 5b. BUILD POWO LOOKUP TABLE ───────────────────────────────────────────────
# Take the first accepted hit per name. Homonyms are flagged above for manual
# resolution; single-hit names are used directly.

powo_lookup <- powo_results_raw |>
  filter(!is.na(name)) |>
  group_by(query_name) |>
  slice(1) |>
  ungroup() |>
  transmute(
    taxon_std        = query_name,
    taxon_powo       = name,
    powo_fqid        = if ("fqId" %in% names(pick(everything()))) fqId else NA_character_,
    homonym_flag     = query_name %in% homonym_flags$query_name,
    powo_access_date = Sys.Date()
  )

# POWO version — no versioning endpoint available; record access date as proxy
POWO_VERSION_NOTE <- paste0(
  "POWO accessed via kewr on ", Sys.Date(),
  ". No formal version endpoint available; date serves as citable timestamp."
)
message(POWO_VERSION_NOTE)

# Names with no POWO match — need manual review or will fall to fuzzy matching
unmatched_powo <- names_to_query[!names_to_query %in% powo_lookup$taxon_std]
if (length(unmatched_powo) > 0) {
  message(sprintf("%d name(s) returned no POWO match:", length(unmatched_powo)))
  print(unmatched_powo)
}

# ── 5c. JOIN POWO RESULTS & RESOLVE SYNONYMS ─────────────────────────────────
# Per workflow iii.1: if name is a synonym, replace taxon_std with the POWO
# accepted name (taxon_accepted) and log the mapping.
# taxon_accepted is the key column used in all downstream operations.

attach_powo <- function(df) {
  df |>
    left_join(powo_lookup, by = "taxon_std") |>
    mutate(
      powo_match = case_when(
        !is.na(taxon_powo) & taxon_powo == taxon_std ~ "exact",
        !is.na(taxon_powo) & taxon_powo != taxon_std ~ "synonym_resolved",
        TRUE                                          ~ "no_powo_match"
      ),
      # taxon_accepted propagates the accepted name where available
      taxon_accepted = case_when(
        powo_match %in% c("exact", "synonym_resolved") ~ taxon_powo,
        TRUE                                           ~ taxon_std  # keep parsed name as fallback
      )
    )
}

df_red_list_cleaned       <- attach_powo(df_red_list_cleaned)
df_forest_species_cleaned <- attach_powo(df_forest_species_cleaned)

# Synonym mapping log for methods reporting
synonym_log <- bind_rows(
  df_red_list_cleaned,
  df_forest_species_cleaned
) |>
  filter(powo_match == "synonym_resolved") |>
  distinct(taxon_std, taxon_accepted, powo_access_date)

message(sprintf("Synonym log: %d synonym → accepted name mappings.", nrow(synonym_log)))


# ── 6. FUZZY MATCHING FOR STILL-UNMATCHED NAMES ───────────────────────────────
# Per workflow iv.1: Levenshtein similarity >= 0.9 (not Jaro-Winkler).
# Reference pool is POWO-accepted names, not the internal parsed list.
# Per workflow iv.2: export a random verification sample (~20 names or 10%).
# Per workflow iv.3: flag fuzzy-matched records for sensitivity analysis.

# Reference pool: POWO-accepted names from successful queries
powo_reference_pool <- powo_lookup |>
  pull(taxon_powo) |>
  unique() |>
  na.omit()

lv_similarity <- function(a, b) {
  # Normalised Levenshtein similarity: 1 - (lv_dist / max_string_length)
  dists   <- stringdist(tolower(a), tolower(b), method = "lv")
  max_len <- pmax(nchar(tolower(a)), nchar(tolower(b)))
  1 - (dists / max_len)
}

fuzzy_match_lv <- function(unmatched_names, reference_names, min_sim = 0.9) {
  map_dfr(unmatched_names, function(nm) {
    sims    <- lv_similarity(nm, reference_names)
    best_idx <- which.max(sims)
    best_sim <- sims[best_idx]
    tibble(
      name_full       = nm,
      fuzzy_match     = reference_names[best_idx],
      lv_similarity   = best_sim,
      fuzzy_accepted  = best_sim >= min_sim
    )
  })
}

unmatched_red    <- df_red_list_cleaned |>
  filter(powo_match == "no_powo_match") |>
  pull(name_full) |>
  unique()

unmatched_forest <- df_forest_species_cleaned |>
  filter(powo_match == "no_powo_match") |>
  pull(name_full) |>
  unique()

fuzzy_red    <- fuzzy_match_lv(unmatched_red,    powo_reference_pool)
fuzzy_forest <- fuzzy_match_lv(unmatched_forest, powo_reference_pool)

# Inspect fuzzy results before accepting — sorted by weakest match first
message("── Fuzzy matches (red list) — review before accepting ──")
fuzzy_red    |> arrange(lv_similarity)
message("── Fuzzy matches (forest list) — review before accepting ──")
fuzzy_forest |> arrange(lv_similarity)

# ── 6a. VERIFICATION SAMPLE (workflow iv.2) ───────────────────────────────────
# Export ~20 names or 10% (whichever is larger) for manual verification
set.seed(42)  # reproducibility
fuzzy_all_accepted <- bind_rows(
  fuzzy_red    |> filter(fuzzy_accepted) |> mutate(source = "red_list"),
  fuzzy_forest |> filter(fuzzy_accepted) |> mutate(source = "forest")
)

n_verify <- max(20L, ceiling(0.1 * nrow(fuzzy_all_accepted)))
fuzzy_verify_sample <- fuzzy_all_accepted |>
  slice_sample(n = min(n_verify, nrow(fuzzy_all_accepted)))

write_xlsx(
  fuzzy_verify_sample,
  path = "output/fuzzy_manual_verification.xlsx"
)
message(sprintf(
  "Fuzzy verification sample: %d rows written to output/fuzzy_manual_verification.xlsx",
  nrow(fuzzy_verify_sample)
))

# ── 6b. PATCH FUZZY MATCHES (with match_method flag, workflow iv.3) ───────────

patch_fuzzy <- function(df, fuzzy_result) {
  fuzzy_lookup <- fuzzy_result |>
    filter(fuzzy_accepted) |>
    select(name_full, taxon_accepted_fuzzy = fuzzy_match, lv_similarity)
  
  df |>
    left_join(fuzzy_lookup, by = "name_full") |>
    mutate(
      taxon_accepted = coalesce(taxon_accepted, taxon_accepted_fuzzy),
      match_method   = case_when(
        powo_match %in% c("exact", "synonym_resolved") ~ powo_match,
        !is.na(taxon_accepted_fuzzy)                   ~ "fuzzy",
        collapse_flag & !is.na(taxon_accepted)         ~ "intraspecific_collapsed",
        !is.na(taxon_accepted)                         ~ "exact",   # fallback for parsed names
        TRUE                                           ~ "unmatched"
      ),
      lv_similarity = if_else(match_method == "fuzzy", lv_similarity, NA_real_)
    ) |>
    select(-taxon_accepted_fuzzy)
}

df_red_list_cleaned       <- patch_fuzzy(df_red_list_cleaned,       fuzzy_red)
df_forest_species_cleaned <- patch_fuzzy(df_forest_species_cleaned, fuzzy_forest)


# ── 7. MANUAL FIXES ───────────────────────────────────────────────────────────
# Section-level Taraxacum names, typo corrections etc.

manual_fixes <- tribble(
  ~name_full,                                               ~taxon_std_manual,
  "KNAUTIA SERPENTINICOLA KOLÁŘ ET AL.",                    "Knautia serpentinicola",
  "CORALLORRHIZA TRIFIDA CHATEL.",                          "Corallorhiza trifida",
  "TARAXACUM SECT. CUCULLATA SOEST",                        NA_character_,
  "TARAXACUM SECT. FONTANA SOEST",                          NA_character_,
  "TARAXACUM LITORALE-GRUPPE",                              NA_character_,
  "TARAXACUM SECT. NAEVOSA M.P. CHRIST.",                   NA_character_,
  "TARAXACUM SECT. OBLIQUA DAHLST.",                        NA_character_,
  "TARAXACUM SECT. PALUSTRIA (H. LINDB.) DAHLST.",          NA_character_,
  "TARAXACUM SECT. RHODOCARPA SOEST",                       NA_character_
)

patch_manual <- function(df) {
  df |>
    left_join(manual_fixes, by = "name_full") |>
    mutate(
      taxon_accepted = coalesce(taxon_accepted, taxon_std_manual),
      match_method   = if_else(
        is.na(match_method) & !is.na(taxon_std_manual),
        "manual",
        match_method
      )
    ) |>
    select(-taxon_std_manual)
}

df_red_list_cleaned       <- patch_manual(df_red_list_cleaned)
df_forest_species_cleaned <- patch_manual(df_forest_species_cleaned)


# ── 8. FINAL JOIN ON taxon_accepted ───────────────────────────────────────────
# Join key is taxon_accepted (POWO-validated accepted name), not the raw
# parsed name. This ensures synonyms from both lists are correctly reconciled.
# code is retained; remove distinct() artefact from previous version.

df_red_list_forest_species <- df_red_list_cleaned |>
  filter(!is.na(taxon_accepted)) |>
  select(taxon_accepted, taxon_std, name_full, name_author,
         red_list_category, powo_match, match_method,
         collapse_flag, homonym_flag, lv_similarity, powo_access_date) |>
  inner_join(
    df_forest_species_cleaned |>
      filter(!is.na(taxon_accepted)) |>
      select(taxon_accepted, code),
    by      = "taxon_accepted",
    # many-to-many arises when one accepted name covers multiple forest codes
    # (e.g. a species present in both closed_forest and edge_disturbance).
    # This is expected; distinct() below collapses to one row per species.
    relationship = "many-to-many"
  ) |>
  # Retain code: keep the first code per species (document this choice in methods)
  distinct(taxon_accepted, .keep_all = TRUE)

# Attach POWO version note as metadata
attr(df_red_list_forest_species, "powo_version_note") <- POWO_VERSION_NOTE
attr(df_red_list_forest_species, "powo_access_date")  <- as.character(Sys.Date())


# ── 9. MATCH SUMMARY ──────────────────────────────────────────────────────────
# Report counts per match_method for methods section

match_summary <- df_red_list_forest_species |>
  count(match_method, sort = TRUE) |>
  mutate(pct = round(100 * n / sum(n), 1))

message("── Match method summary ──")
print(match_summary)

# Species flagged for sensitivity analysis (fuzzy + intraspecific_collapsed)
sensitivity_flags <- df_red_list_forest_species |>
  filter(match_method %in% c("fuzzy", "intraspecific_collapsed")) |>
  select(taxon_accepted, taxon_std, match_method, lv_similarity,
         collapse_flag, powo_match)

message(sprintf(
  "%d species flagged for sensitivity analysis (fuzzy or intraspecific collapse).",
  nrow(sensitivity_flags)
))

# Homonym review table
if (nrow(homonym_flags) > 0) {
  message("── Names requiring homonym resolution via EVA header ──")
  print(
    df_red_list_forest_species |>
      filter(homonym_flag) |>
      select(taxon_accepted, taxon_std, name_full, powo_match)
  )
}
# ── 10. EXPORT ────────────────────────────────────────────────────────────────

write_xlsx(
  list(
    species_list       = df_red_list_forest_species,
    synonym_log        = synonym_log,
    match_summary      = match_summary,
    sensitivity_flags  = sensitivity_flags,
    homonym_review     = homonym_flags,
    fuzzy_verify       = fuzzy_verify_sample
  ),
  path = "cleaned_data/eva_request_species_harmonized.xlsx"
)

message("Done. Output written to cleaned_data/eva_request_species_harmonized.xlsx")