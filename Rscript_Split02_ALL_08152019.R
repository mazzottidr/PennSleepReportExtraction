################################
############ Extract Split_S01
################################
library(dplyr)
library(lubridate)
library(stringr)
source('~/relevant_scripts/ExtractPSGReports_Split_2_function.R')
ALL_annotated <- readRDS("PennSleepDatabase_Paths_ALL_07242019.Rdata")
extracted_s02 <- extract_split_2(filter(ALL_annotated, predictedTableFormat=="2", predictedStudyType=="Split")$linux_path)