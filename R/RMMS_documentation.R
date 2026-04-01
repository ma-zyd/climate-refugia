library(rangeModelMetadata)

#------------------Authorship---------------------------------------
rmm <- rmmTemplate(family = NULL)
rmm$authorship$rmmName <- "Zydorek_2026_red_list_forest_herbs_germany"
rmm$authorship$names <- "Zydorek, Matthias"
rmm$authorship$contact <- "matthias.zydorek@stud.uni-goettingen.de"
rmm$authorship$organization <- "University of Göttingen"

#---------------Data------------------------------------------
rmm$dataoccurence$taxon <- df_red_list_forest_species$taxon_std
rmm$data$occurence$dataType <- "presence only"
rmm$dataoccurence$source <- "European Vegetation Archive (EVA)"
