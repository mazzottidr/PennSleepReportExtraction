################################
############ Extract Split_S01
################################
library(dplyr)
library(lubridate)
library(stringr)
source('~/relevant_scripts/ExtractPSGReports_Split_1_function.R')
ALL_annotated <- readRDS("PennSleepDatabase_Paths_ALL_07242019.Rdata")
extracted_s01 <- extract_split_1(filter(ALL_annotated, predictedTableFormat=="1", predictedStudyType=="Split")$linux_path)