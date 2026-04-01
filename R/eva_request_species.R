# used packages
library(devtools)
library(LCVP)
library(lcvplants)
library(rangeModelMetadata)
library(readxl)
library(tidyverse)
library(writexl) 

# --- Load & clean red list ---------------------------------------------------
df_red_list <- read_excel("data/red_list_plants_germany.xlsx")
df_red_list_cleaned<- df_red_list |>  
  filter(
    red_list_category %in% c(1, 2, 3, "R") # filter red list species  based on categories
  ) |> 
  rename(
    name_full= name
  )  
# --- Load & clean forest species -------------------------------------------- 
df_forest_species <- read_excel("data/forest_plants_germany.xls")
df_forest_species_cleaned <- df_forest_species |> 
  pivot_longer(
    cols = NT:D,
    names_to = "distribution", 
    values_to = "Value"
  ) |> 
  extract(
    Value, 
    into = c(
      "habitus",
      "code"
    ),
    regex = "([BKO])(.*)"
  ) |> 
  filter(!str_detect(
    code, "2\\.1|2\\.2")
  )|> 
  mutate(code = case_when(
    code == "1.1" ~ "closed_forest",  
    code == "1.2" ~ "edge_disturbance",
    TRUE ~ code)
  ) |> 
  filter(
    str_detect(
      habitus,"K")
  ) |> 
  rename(
    name_full= `Wissenschaftlicher Name`
  ) 
# --- Exact Search: Harmonize names via LCVP -----------------------------------------------
red_list_harmonized <- lcvp_search(
  df_red_list_cleaned$name_full
  )
forest_harmonized <- lcvp_search(
  df_forest_species_cleaned$name_full
  )


df_red_list_cleaned <- df_red_list_cleaned |>
  mutate(
    taxon_std = trimws(
      red_list_harmonized$Output.Taxon
      )
    )

df_forest_species_cleaned <- df_forest_species_cleaned |>
  mutate(
    taxon_std = trimws(
      forest_harmonized$Output.Taxon
      )
    )
#----------fuzzy search for NAs--------------------------------------
unmatched_red_names <- df_red_list_cleaned |> 
  filter(
    is.na(taxon_std)
    ) |> 
  pull(
    name_full
    ) |> 
  unique()

unmatched_forest_names <- df_forest_species_cleaned |> 
  filter(
    is.na(taxon_std)
    ) |> 
  pull(
    name_full
    )|> 
  unique()

fuzzy_red    <- lcvp_fuzzy_search(
  unmatched_red_names,
  max_distance = 0.1,
  keep_closest = TRUE,
  progress_bar = TRUE
  )
fuzzy_forest <- lcvp_fuzzy_search(
  unmatched_forest_names,
  max_distance = 0.1,
  keep_closest = TRUE,
  progress_bar = TRUE
  )
# Inspection of results
fuzzy_red    |>
  mutate(Input.Name = paste(Input.Genus, Input.Epitheton)) |>
  select(Input.Name, Output.Taxon, Name.Distance) |> 
  arrange(desc(Name.Distance))

fuzzy_forest |> 
  mutate(Input.Name = paste(Input.Genus, Input.Epitheton)) |>
  select(Input.Name, Output.Taxon, Name.Distance) |>
  arrange(desc(Name.Distance))

#  lookup from fuzzy results and patch in
fuzzy_red_lookup <- fuzzy_red |>
  filter(
    !is.na(Output.Taxon)
    ) |>
  mutate(
    name_full = paste(
      Input.Genus, Input.Epitheton
      )
    ) |> 
  select(
    name_full,
    taxon_std_fuzzy = Output.Taxon
    ) |>
  mutate(
    taxon_std_fuzzy = trimws(
      taxon_std_fuzzy
      )
    )


df_red_list_cleaned <- df_red_list_cleaned |>
  left_join(fuzzy_red_lookup, by = "name_full") |>
  mutate(taxon_std = coalesce(taxon_std, taxon_std_fuzzy)) |>
  select(-taxon_std_fuzzy)

if (!is.null(fuzzy_forest)) {
  fuzzy_forest_lookup <- fuzzy_forest |>
    filter(!is.na(Output.Taxon)) |>
    mutate(name_full = paste(Input.Genus, Input.Epitheton)) |> 
    select(name_full, taxon_std_fuzzy = Output.Taxon) |>
    mutate(taxon_std_fuzzy = trimws(taxon_std_fuzzy))

  df_forest_species_cleaned <- df_forest_species_cleaned |>
    left_join(fuzzy_forest_lookup, by = "name_full") |>
    mutate(taxon_std = coalesce(taxon_std, taxon_std_fuzzy)) |>
    select(-taxon_std_fuzzy)
}
                                                      

#------------- Manual fixes for still unmatched names ---------------
manual_fixes <- c(
  "KNAUTIA SERPENTINICOLA KOLÁŘ ET AL."            = "Knautia serpentinicola",  # strip "et al."
  "TARAXACUM SECT. CUCULLATA SOEST"                = NA,  # section-level, no species equivalent
  "TARAXACUM SECT. FONTANA SOEST"                  = NA,
  "TARAXACUM LITORALE-GRUPPE"                      = NA,  # not a valid binomial
  "TARAXACUM SECT. NAEVOSA M.P. CHRIST."           = NA,
  "TARAXACUM SECT. OBLIQUA DAHLST."                = NA,
  "TARAXACUM SECT. PALUSTRIA (H. LINDB.) DAHLST."  = NA,
  "TARAXACUM SECT. RHODOCARPA SOEST"               = NA,
  "CORALLORRHIZA TRIFIDA CHATEL."                  = "Corallorhiza trifida"     # fix spelling
)

# Run lcvp_search on the correctable names only
correctable <- enframe(
  manual_fixes,
  name = "name_full",
  value = "name_corrected"
  ) |> 
  filter(
    !is.na(name_corrected
           )
    )
correction_results <- lcvp_search(correctable$name_corrected)

# Attach standardized Output.Taxon back to the manual fixes table
manual_lookup <- correctable |>
  mutate(taxon_std_manual = trimws(correction_results$Output.Taxon)) |>
  select(name_full, taxon_std_manual) |>
  bind_rows(                                     # add the NA entries (invalid names)
    manual_fixes |>
      enframe(name = "name_full", value = "name_corrected") |> 
      filter(is.na(name_corrected)) |>
      select(name_full) |>
      mutate(taxon_std_manual = NA_character_)
  )

# Patch manual fixes into both data frames
df_red_list_cleaned <- df_red_list_cleaned |>
  left_join(manual_lookup, by = "name_full") |>
  mutate(taxon_std = coalesce(taxon_std, taxon_std_manual)) |>
  select(-taxon_std_manual)

df_forest_species_cleaned <- df_forest_species_cleaned |>
  left_join(manual_lookup, by = "name_full") |>
  mutate(taxon_std = coalesce(taxon_std, taxon_std_manual)) |>
  select(-taxon_std_manual)

# --- STEP 4: Final join -------------------------------------------------
df_red_list_forest_species <- df_red_list_cleaned |>
  filter(!is.na(taxon_std)) |> 
  select(taxon_std, name_full, red_list_category) |>
  inner_join(
    df_forest_species_cleaned |>
      filter(!is.na(taxon_std)) |> 
      select(
        taxon_std,
        distribution,
        code
        ),
    by = "taxon_std",
    relationship = "many-to-many") |> 
  distinct(
    taxon_std, 
    .keep_all = TRUE
    ) |>  
  select(-distribution)

write.csv(df_red_list_forest_species,"cleaned_data/red_list_forest_species_germany.csv",row.names = FALSE)

