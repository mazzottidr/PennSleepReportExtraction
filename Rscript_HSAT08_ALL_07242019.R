################################
############ Extract HSAT_S08
################################
library(dplyr)
library(lubridate)
library(stringr)
source('~/relevant_scripts/ExtractPSGReports_HSAT_8_function.R')
ALL_annotated <- readRDS("PennSleepDatabase_Paths_ALL_07242019.Rdata")
extracted_h08 <- extract_hsat_8(filter(ALL_annotated, predictedTableFormat=="8", predictedStudyType=="HSAT")$linux_path)