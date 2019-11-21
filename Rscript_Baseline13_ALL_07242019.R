################################
############ Extract Baseline_13
################################
library(dplyr)
library(lubridate)
library(stringr)
source('~/relevant_scripts/ExtractPSGReports_Baseline_13_function.R')
ALL_annotated <- readRDS("PennSleepDatabase_Paths_ALL_07242019.Rdata")
extracted_b13 <- extract_baseline_13(filter(ALL_annotated, predictedTableFormat=="13", predictedStudyType=="Baseline")$linux_path)