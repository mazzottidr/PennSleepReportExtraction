################################
############ Extract Baseline_02
################################
library(dplyr)
library(lubridate)
library(stringr)
source('~/relevant_scripts/ExtractPSGReports_Baseline_2_function.R')
ALL_annotated <- readRDS("PennSleepDatabase_Paths_ALL_07242019.Rdata")
extracted_b02 <- extract_baseline_2(filter(ALL_annotated, predictedTableFormat=="2", predictedStudyType=="Baseline")$linux_path)