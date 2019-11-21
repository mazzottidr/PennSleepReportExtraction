################################
############ Extract Baseline_03
################################
library(dplyr)
library(lubridate)
library(stringr)
source('~/relevant_scripts/ExtractPSGReports_Baseline_3_function.R')
ALL_annotated <- readRDS("PennSleepDatabase_Paths_ALL_07242019.Rdata")
extracted_b03 <- extract_baseline_3(filter(ALL_annotated, predictedTableFormat=="3", predictedStudyType=="Baseline")$linux_path)