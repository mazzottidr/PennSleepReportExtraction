library(dplyr)
library(lubridate)
library(stringr)

setwd("/data1/home/diegomaz/ALL_07232019")

# Extracting data from all reports

# Load previously annotated data
ALL_annotated <- readRDS("/data1/home/diegomaz/All_processing/all_with_paths_wData_ALL_042219.Rdata")
ALL_annotated$IsExtracted <- FALSE

#Fix paths to new computer
ALL_annotated$linux_path <- paste0("/data1/home/diegomaz/sleepstudyreports_Dec2018/", sapply(strsplit(ALL_annotated$path,"/", useBytes=T), "[[", 2))

# Save
saveRDS(ALL_annotated, "PennSleepDatabase_Paths_ALL_07242019.Rdata")

# Get numbers to extract
table(ALL_annotated$predictedTableFormat, ALL_annotated$predictedStudyType)

# Potential extractions:
# Type  Baseline Baseline_MSLT  HSAT Letter  MSLT Split Treatment
# 1     14655             0     0      0     0  6146      6576
# 2      5351            22     0      0     0  1867      1004
# 3      8312             0     0      0     0  1878       693
# 4         8             0     0      0     0    11         4
# 5      5051             0     6      8   247   545      2398
# 6      1867             0     0      0     0  1691        27
# 7         0             0     0      0     0     0         0
# 8         1             0  1148      0     0     0        53
# 9         4            14     0      0     0     0         0
# 10        4             0     0      0   648     0         0
# 11        5             0     0      0     0     0         0
# 12        0             0     0      0     0     0         0
# 13      794             0     0      0     0     3         0
# 14      220             0     0    113     0     0        26


#### Important notes about study formats:
# 1, 2 and 3 - working well
# 4 - too low frequency
# 5 - information present in text, not tables, usually only first summary page of the report
# 6 - it lacks information about BMI, gender, and detailed sleep study info. Usually older studies
# 13 - Base line is same format as 2
# 14 - letter only

# 07/24/2019 - using available scripts, expected to extract 34464 reports (Baseline 1,2,3 and Split 1)
# 08/15/2019 - using available scripts, expected to extract 40151 reports (Baseline 1,2,3,13 and Split 1,2,3 and HSAT 8)



##### Create blocks below as independent Rscripts

# Current extraction pipelines

################################
############ Extract Baseline_01
################################
library(dplyr)
library(lubridate)
library(stringr)
source('~/relevant_scripts/ExtractPSGReports_Baseline_1_function.R')
ALL_annotated <- readRDS("PennSleepDatabase_Paths_ALL_07242019.Rdata")
extracted_b01 <- extract_baseline_1(filter(ALL_annotated, predictedTableFormat=="1", predictedStudyType=="Baseline")$linux_path)

################################
############ Extract Baseline_02
################################
library(dplyr)
library(lubridate)
library(stringr)
source('~/relevant_scripts/ExtractPSGReports_Baseline_2_function.R')
ALL_annotated <- readRDS("PennSleepDatabase_Paths_ALL_07242019.Rdata")
extracted_b02 <- extract_baseline_2(filter(ALL_annotated, predictedTableFormat=="2", predictedStudyType=="Baseline")$linux_path)

################################
############ Extract Baseline_03
################################
library(dplyr)
library(lubridate)
library(stringr)
source('~/relevant_scripts/ExtractPSGReports_Baseline_3_function.R')
ALL_annotated <- readRDS("PennSleepDatabase_Paths_ALL_07242019.Rdata")
extracted_b03 <- extract_baseline_3(filter(ALL_annotated, predictedTableFormat=="3", predictedStudyType=="Baseline")$linux_path)

################################
############ Extract Baseline_13
################################
library(dplyr)
library(lubridate)
library(stringr)
source('~/relevant_scripts/ExtractPSGReports_Baseline_13_function.R')
ALL_annotated <- readRDS("PennSleepDatabase_Paths_ALL_07242019.Rdata")
extracted_b13 <- extract_baseline_13(filter(ALL_annotated, predictedTableFormat=="13", predictedStudyType=="Baseline")$linux_path)

################################
############ Extract Split_S01
################################
library(dplyr)
library(lubridate)
library(stringr)
source('~/relevant_scripts/ExtractPSGReports_Split_1_function.R')
ALL_annotated <- readRDS("PennSleepDatabase_Paths_ALL_07242019.Rdata")
extracted_s01 <- extract_split_1(filter(ALL_annotated, predictedTableFormat=="1", predictedStudyType=="Split")$linux_path)

################################
############ Extract Split_S02
################################
library(dplyr)
library(lubridate)
library(stringr)
source('~/relevant_scripts/ExtractPSGReports_Split_2_function.R')
ALL_annotated <- readRDS("PennSleepDatabase_Paths_ALL_07242019.Rdata")
extracted_s02 <- extract_split_2(filter(ALL_annotated, predictedTableFormat=="2", predictedStudyType=="Split")$linux_path)

################################
############ Extract Split_S03
################################
library(dplyr)
library(lubridate)
library(stringr)
source('~/relevant_scripts/ExtractPSGReports_Split_3_function.R')
ALL_annotated <- readRDS("PennSleepDatabase_Paths_ALL_07242019.Rdata")
extracted_s03 <- extract_split_3(filter(ALL_annotated, predictedTableFormat=="3", predictedStudyType=="Split")$linux_path)

################################
############ Extract HSAT_S08
################################
library(dplyr)
library(lubridate)
library(stringr)
source('~/relevant_scripts/ExtractPSGReports_HSAT_8_function.R')
ALL_annotated <- readRDS("PennSleepDatabase_Paths_ALL_07242019.Rdata")
extracted_h08 <- extract_hsat_8(filter(ALL_annotated, predictedTableFormat=="8", predictedStudyType=="HSAT")$linux_path)


###################### DONE OR RUNNING UNTIL HERE ##############################




