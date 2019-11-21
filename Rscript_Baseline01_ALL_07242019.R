################################
############ Extract Baseline_01
################################
library(dplyr)
library(lubridate)
library(stringr)
source('~/relevant_scripts/ExtractPSGReports_Baseline_1_function.R')
ALL_annotated <- readRDS("PennSleepDatabase_Paths_ALL_07242019.Rdata")
extracted_b01 <- extract_baseline_1(filter(ALL_annotated, predictedTableFormat=="1", predictedStudyType=="Baseline")$linux_path)