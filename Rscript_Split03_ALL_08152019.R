################################
############ Extract Split_S03
################################
library(dplyr)
library(lubridate)
library(stringr)
source('~/relevant_scripts/ExtractPSGReports_Split_3_function.R')
ALL_annotated <- readRDS("PennSleepDatabase_Paths_ALL_07242019.Rdata")
extracted_s03 <- extract_split_3(filter(ALL_annotated, predictedTableFormat=="3", predictedStudyType=="Split")$linux_path)