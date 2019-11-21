# Create REDcap dataset

library(dplyr)
library(lubridate)
library(stringr)
library(purrr)

setwd("/data1/home/diegomaz/ALL_07232019/combined_data")

# Load data
SleepStudy_GAP_PMBB_Summary <- readRDS("SleepStudy_GAP_PMBB_Summary.Rdata")
PennSleepMetadata <- readRDS("PennSleepMetadata_AfterExtraction_08162019.Rdata")
extracted_sleep_data <- readRDS("ExtractedSleepData_lists.Rdata")

REDCap_template <- read.csv("UniversityOfPennsylvaniaGeneti_ImportTemplate_2019-02-28.csv", stringsAsFactors = F)
labels_REDCap <- read.csv("Labels_REDCap_08212019.csv", stringsAsFactors = F)

label_file_df <- data.frame(list_name=names(extracted_sleep_data), label_id_names=c("B01_label",
                                                                                    "B02_label",
                                                                                    "B03_label",
                                                                                    "B13_label",
                                                                                    "H08_label",
                                                                                    "S01_label",
                                                                                    "S02_label",
                                                                                    "S03_label"), stringsAsFactors = F)



fix_height <- function(h) {
        return(sapply(strsplit(as.character(h),"'|\""),
               function(x){12*as.numeric(x[1]) + as.numeric(x[2])}))
}

##### Clean and ensure data is right in each dataset
# OBS: not excluding outliers, only values that are not valid or not make sense

##############
########## B01
##############

B01_RC <- extracted_sleep_data$PennSleepDatabase_Baseline01_20190724200859.Rdata %>%
        select(labels_REDCap$B01_label[!is.na(labels_REDCap$B01_label)])

B01_RC$B01_PSG_PHI_Filename <- map_chr(strsplit(B01_RC$B01_PSG_PHI_Filename, "/"), pluck(2))
B01_RC$B01_PSG_Study_date <- mdy(B01_RC$B01_PSG_Study_date)
#B01_RC$B01_PSG_Lights_off <- hms(B01_RC$B01_PSG_Lights_off)
#B01_RC$B01_PSG_Lights_on <- hms(B01_RC$B01_PSG_Lights_on)
B01_RC$B01_PSG_Min_SaO2_TST_perc[B01_RC$B01_PSG_Min_SaO2_TST_perc<0] <- NA
B01_RC$B01_PSG_Sex[grepl("Female", B01_RC$B01_PSG_Sex)] <- "Female"
B01_RC$B01_PSG_Sex[grepl("Male", B01_RC$B01_PSG_Sex)] <- "Male"
B01_RC$B01_PSG_Sex[grepl("female", B01_RC$B01_PSG_Sex)] <- "Female"
B01_RC$B01_PSG_Sex[grepl("^male", B01_RC$B01_PSG_Sex)] <- "Male"
B01_RC$B01_PSG_Sex[B01_RC$B01_PSG_Sex=="BP :"] <- NA
B01_RC$B01_PSG_Age_at_Study[B01_RC$B01_PSG_Age_at_Study<=0] <- NA
B01_RC$B01_PSG_BMI[B01_RC$B01_PSG_BMI>500] <- NA
B01_RC$B01_PSG_BMI[B01_RC$B01_PSG_BMI==0] <- NA
B01_RC$B01_PSG_Height <- fix_height(B01_RC$B01_PSG_Height)
B01_RC$B01_PSG_Height[B01_RC$B01_PSG_Height>110] <- NA
B01_RC$B01_PSG_Weight <- as.numeric(gsub("([0-9]+).*$", "\\1",B01_RC$B01_PSG_Weight))
B01_RC$B01_PSG_Weight[B01_RC$B01_PSG_Weight<=10] <- NA
B01_RC$B01_PSG_Blood_Pressure_sbp <- as.numeric(gsub("\\.+", "", gsub("-.+", "", gsub("/.+", "", B01_RC$B01_PSG_Blood_Pressure))))
B01_RC$B01_PSG_Blood_Pressure_dbp <- as.numeric(gsub(".+\\\\", "", gsub(".+-", "", gsub(".+/", "", B01_RC$B01_PSG_Blood_Pressure))))
B01_RC$B01_PSG_Blood_Pressure_sbp[B01_RC$B01_PSG_Blood_Pressure_sbp>500] <- NA
B01_RC$B01_PSG_Blood_Pressure_dbp[B01_RC$B01_PSG_Blood_Pressure_dbp>500] <- NA
B01_RC$B01_PSG_Blood_Pressure <- NA

b01_label_df <- data.frame(b01=labels_REDCap$B01[!is.na(labels_REDCap$B01)], rc=labels_REDCap$REDCap_label[!is.na(labels_REDCap$B01)], stringsAsFactors = F)

REDCap_B01_df <- data.frame(id=paste0("B01_", sprintf("%08d", 1:nrow(B01_RC))), stringsAsFactors = F)
for (v in colnames(REDCap_template)) {
        
        REDCap_B01_df[,v] <- B01_RC[,b01_label_df$b01[b01_label_df$rc==v]]
        
        
}
# Fix BP
REDCap_B01_df$sbp_at_study <- B01_RC$B01_PSG_Blood_Pressure_sbp
REDCap_B01_df$dbp_at_study <- B01_RC$B01_PSG_Blood_Pressure_dbp


##############
########## B02
##############

B02_RC <- extracted_sleep_data$PennSleepDatabase_Baseline02_20190816221324.Rdata %>%
        select(labels_REDCap$B02_label[!is.na(labels_REDCap$B02_label)])

B02_RC$B02_PSG_PHI_Filename <- map_chr(strsplit(B02_RC$B02_PSG_PHI_Filename, "/"), pluck(6))
B02_RC$B02_PSG_NPBMStudyInfoStudyDate <- mdy(B02_RC$B02_PSG_NPBMStudyInfoStudyDate)
#B02_RC$B02_PSG_VAR_StartTime <- hms(B02_RC$B02_PSG_VAR_StartTime)
#B02_RC$B02_PSG_VAR_EndTime <- hms(B02_RC$B02_PSG_VAR_EndTime)
B02_RC$B02_PSG_NPBMPatientInfoGender[grepl("female", B02_RC$B02_PSG_NPBMPatientInfoGender)] <- "Female"
B02_RC$B02_PSG_NPBMPatientInfoGender[grepl("^male", B02_RC$B02_PSG_NPBMPatientInfoGender)] <- "Male"
B02_RC$B02_PSG_Age_at_Study[B02_RC$B02_PSG_Age_at_Study<=0] <- NA
B02_RC$B02_PSG_NPBMPatientInfoBMI[B02_RC$B02_PSG_NPBMPatientInfoBMI==0] <- NA
B02_RC$B02_PSG_NPBMPatientInfoHeight[B02_RC$B02_PSG_NPBMPatientInfoHeight>110] <- NA
B02_RC$B02_PSG_NPBMPatientInfoHeight[B02_RC$B02_PSG_NPBMPatientInfoHeight==0] <- NA
B02_RC$B02_PSG_NPBMPatientInfoWeight[B02_RC$B02_PSG_NPBMPatientInfoWeight==0] <- NA
B02_RC$B02_PSG_NPBMCustomPatientInfoBloodPressure_sbp <- as.numeric(gsub("\\.+", "", gsub("-.+", "", gsub("/.+", "", B02_RC$B02_PSG_NPBMCustomPatientInfoBloodPressure))))
B02_RC$B02_PSG_NPBMCustomPatientInfoBloodPressure_dbp <- as.numeric(gsub(".+\\\\", "", gsub(".+-", "", gsub(".+/", "", B02_RC$B02_PSG_NPBMCustomPatientInfoBloodPressure))))
B02_RC$B02_PSG_NPBMCustomPatientInfoBloodPressure_sbp[B02_RC$B02_PSG_NPBMCustomPatientInfoBloodPressure_sbp>500] <- NA
B02_RC$B02_PSG_NPBMCustomPatientInfoBloodPressure_dbp[B02_RC$B02_PSG_NPBMCustomPatientInfoBloodPressure_dbp>500] <- NA

B02_label_df <- data.frame(B02=labels_REDCap$B02[!is.na(labels_REDCap$B02)], rc=labels_REDCap$REDCap_label[!is.na(labels_REDCap$B02)], stringsAsFactors = F)

REDCap_B02_df <- data.frame(id=paste0("B02_", sprintf("%08d", 1:nrow(B02_RC))), stringsAsFactors = F)
for (v in colnames(REDCap_template)) {
        
        REDCap_B02_df[,v] <- B02_RC[,B02_label_df$B02[B02_label_df$rc==v]]
        
        
}
# Fix BP
REDCap_B02_df$sbp_at_study <- B02_RC$B02_PSG_NPBMCustomPatientInfoBloodPressure_sbp
REDCap_B02_df$dbp_at_study <- B02_RC$B02_PSG_NPBMCustomPatientInfoBloodPressure_dbp

##############
########## B03
##############

B03_RC <- extracted_sleep_data$PennSleepDatabase_Baseline03_20190817050915.Rdata %>%
        select(labels_REDCap$B03_label[!is.na(labels_REDCap$B03_label)])

B03_RC$B03_PSG_PHI_Filename <- map_chr(strsplit(B03_RC$B03_PSG_PHI_Filename, "/"), pluck(6))
B03_RC$B03_PSG_NPBMStudyInfoStudyDate <- mdy(B03_RC$B03_PSG_NPBMStudyInfoStudyDate)
#B03_RC$B03_PSG_VAR_LightsOff <- hms(B03_RC$B03_PSG_VAR_LightsOff)
#B03_RC$B03_PSG_VAR_LightsOn <- hms(B03_RC$B03_PSG_VAR_LightsOn)
B03_RC$B03_PSG_NPBMPatientInfoGender[grepl("^male", B03_RC$B03_PSG_NPBMPatientInfoGender)] <- "Male"
B03_RC$B03_PSG_NPBMPatientInfoBMI[B03_RC$B03_PSG_NPBMPatientInfoBMI>500] <- NA
B03_RC$B03_PSG_NPBMPatientInfoBMI[B03_RC$B03_PSG_NPBMPatientInfoBMI==0] <- NA
B03_RC$B03_PSG_NPBMPatientInfoHeight[B03_RC$B03_PSG_NPBMPatientInfoHeight>110] <- NA
B03_RC$B03_PSG_NPBMPatientInfoHeight[B03_RC$B03_PSG_NPBMPatientInfoHeight==0] <- NA
B03_RC$B03_PSG_NPBMPatientInfoWeight[B03_RC$B03_PSG_NPBMPatientInfoWeight==0] <- NA
B03_RC$B03_PSG_NPBMPatientInfoWeight[B03_RC$B03_PSG_NPBMPatientInfoWeight>1400] <- NA

B03_label_df <- data.frame(B03=labels_REDCap$B03[!is.na(labels_REDCap$B03)], rc=labels_REDCap$REDCap_label[!is.na(labels_REDCap$B03)], stringsAsFactors = F)

REDCap_B03_df <- data.frame(id=paste0("B03_", sprintf("%08d", 1:nrow(B03_RC))), stringsAsFactors = F)
for (v in colnames(REDCap_template)) {
        
        REDCap_B03_df[,v] <- B03_RC[,B03_label_df$B03[B03_label_df$rc==v]]
        
        
}


##############
########## S01
##############

S01_RC <- extracted_sleep_data$PennSleepDatabase_Split01_20190815205518.Rdata %>%
        select(labels_REDCap$S01_label[!is.na(labels_REDCap$S01_label)])


S01_RC$S01_PSG_PHI_Filename <- map_chr(strsplit(S01_RC$S01_PSG_PHI_Filename, "/"), pluck(2))
S01_RC$S01_PSG_Study_date <- mdy(S01_RC$S01_PSG_Study_date)
#S01_RC$S01_PSG_BaselineStartTime_baseline <- hms(S01_RC$S01_PSG_BaselineStartTime_baseline)
#S01_RC$S01_PSG_BaselineEndTime_baseline <- hms(S01_RC$S01_PSG_BaselineEndTime_baseline)
S01_RC$S01_PSG_REM_Latency_baseline[S01_RC$S01_PSG_REM_Latency_baseline<0] <- NA
S01_RC$S01_PSG_Mean_SaO2_awake_perc_baseline[S01_RC$S01_PSG_Mean_SaO2_awake_perc_baseline>100] <- NA
S01_RC$S01_PSG_Min_SaO2_TST_perc_baseline[S01_RC$S01_PSG_Min_SaO2_TST_perc_baseline<0] <- NA
S01_RC$S01_PSG_Sex[grepl("female", S01_RC$S01_PSG_Sex)] <- "Female"
S01_RC$S01_PSG_Sex[grepl("^male", S01_RC$S01_PSG_Sex)] <- "Male"
S01_RC$S01_PSG_Sex[grepl("MALE", S01_RC$S01_PSG_Sex)] <- "Male"
S01_RC$S01_PSG_Age_at_Study[S01_RC$S01_PSG_Age_at_Study<=0] <- NA
S01_RC$S01_PSG_BMI[S01_RC$S01_PSG_BMI==0] <- NA
S01_RC$S01_PSG_Height <- fix_height(S01_RC$S01_PSG_Height)
S01_RC$S01_PSG_Height[S01_RC$S01_PSG_Height>110] <- NA
S01_RC$S01_PSG_Weight <- as.numeric(gsub("([0-9]+).*$", "\\1",S01_RC$S01_PSG_Weight))
S01_RC$S01_PSG_Weight[S01_RC$S01_PSG_Weight<=10] <- NA
S01_RC$S01_PSG_Blood_Pressure_sbp <- as.numeric(gsub("\\.+", "", gsub("-.+", "", gsub("/.+", "", S01_RC$S01_PSG_Blood_Pressure))))
S01_RC$S01_PSG_Blood_Pressure_dbp <- as.numeric(gsub(".+\\\\", "", gsub(".+-", "", gsub(".+/", "", S01_RC$S01_PSG_Blood_Pressure))))
S01_RC$S01_PSG_Blood_Pressure_sbp[S01_RC$S01_PSG_Blood_Pressure_sbp>500] <- NA
S01_RC$S01_PSG_Blood_Pressure_dbp[S01_RC$S01_PSG_Blood_Pressure_dbp>500] <- NA

S01_label_df <- data.frame(S01=labels_REDCap$S01[!is.na(labels_REDCap$S01)], rc=labels_REDCap$REDCap_label[!is.na(labels_REDCap$S01)], stringsAsFactors = F)

REDCap_S01_df <- data.frame(id=paste0("S01_", sprintf("%08d", 1:nrow(S01_RC))), stringsAsFactors = F)
for (v in colnames(REDCap_template)) {
        
        REDCap_S01_df[,v] <- S01_RC[,S01_label_df$S01[S01_label_df$rc==v]]
        
        
}

# Fix BP
REDCap_S01_df$sbp_at_study <- S01_RC$S01_PSG_Blood_Pressure_sbp
REDCap_S01_df$dbp_at_study <- S01_RC$S01_PSG_Blood_Pressure_dbp



##############
########## S02
##############

S02_RC <- extracted_sleep_data$PennSleepDatabase_Split02_20190816210127.Rdata %>%
        select(labels_REDCap$S02_label[!is.na(labels_REDCap$S02_label)])


S02_RC$S02_PSG_PHI_Filename <- map_chr(strsplit(S02_RC$S02_PSG_PHI_Filename, "/"), pluck(6))
S02_RC$S02_PSG_NPBMStudyInfoStudyDate <- mdy(S02_RC$S02_PSG_NPBMStudyInfoStudyDate)
#S02_RC$S02_PSG_StartTime_baseline <- hms(S02_RC$S02_PSG_StartTime_baseline)
#S02_RC$S02_PSG_EndTime_baseline <- hms(S02_RC$S02_PSG_EndTime_baseline)
S02_RC$S02_PSG_NPBMPatientInfoGender[grepl("^male", S02_RC$S02_PSG_NPBMPatientInfoGender)] <- "Male"
S02_RC$S02_PSG_Age_at_Study[S02_RC$S02_PSG_Age_at_Study==0] <- NA
S02_RC$S02_PSG_NPBMPatientInfoBMI[S02_RC$S02_PSG_NPBMPatientInfoBMI==0] <- NA
S02_RC$S02_PSG_NPBMPatientInfoHeight[S02_RC$S02_PSG_NPBMPatientInfoHeight==0] <- NA
S02_RC$S02_PSG_NPBMPatientInfoWeight[S02_RC$S02_PSG_NPBMPatientInfoWeight==0] <- NA
S02_RC$S02_PSG_Blood_Pressure_sbp <- as.numeric(gsub("\\.+", "", gsub("-.+", "", gsub("/.+", "", S02_RC$S02_PSG_NPBMCustomPatientInfoBloodPressure))))
S02_RC$S02_PSG_Blood_Pressure_dbp <- as.numeric(gsub(".+\\\\", "", gsub(".+-", "", gsub(".+/", "", S02_RC$S02_PSG_NPBMCustomPatientInfoBloodPressure))))
S02_RC$S02_PSG_Blood_Pressure_sbp[S02_RC$S02_PSG_Blood_Pressure_sbp>500] <- NA
S02_RC$S02_PSG_Blood_Pressure_dbp[S02_RC$S02_PSG_Blood_Pressure_dbp>500] <- NA
S02_RC$S02_PSG_Blood_Pressure_sbp[S02_RC$S02_PSG_Blood_Pressure_sbp==0] <- NA
S02_RC$S02_PSG_Blood_Pressure_dbp[S02_RC$S02_PSG_Blood_Pressure_dbp==0] <- NA

S02_label_df <- data.frame(S02=labels_REDCap$S02[!is.na(labels_REDCap$S02)], rc=labels_REDCap$REDCap_label[!is.na(labels_REDCap$S02)], stringsAsFactors = F)

REDCap_S02_df <- data.frame(id=paste0("S02_", sprintf("%08d", 1:nrow(S02_RC))), stringsAsFactors = F)
for (v in colnames(REDCap_template)) {
        
        REDCap_S02_df[,v] <- S02_RC[,S02_label_df$S02[S02_label_df$rc==v]]
        
}

# Fix BP
REDCap_S02_df$sbp_at_study <- S02_RC$S02_PSG_Blood_Pressure_sbp
REDCap_S02_df$dbp_at_study <- S02_RC$S02_PSG_Blood_Pressure_dbp


##############
########## S03
##############

S03_RC <- extracted_sleep_data$PennSleepDatabase_Split03_20190820162902.Rdata %>%
        select(labels_REDCap$S03_label[!is.na(labels_REDCap$S03_label)])


S03_RC$S03_PSG_PHI_Filename <- map_chr(strsplit(S03_RC$S03_PSG_PHI_Filename, "/"), pluck(6))
S03_RC$S03_PSG_NPBMStudyInfoStudyDate <- mdy(S03_RC$S03_PSG_NPBMStudyInfoStudyDate)
#S03_RC$S03_PSG_StartTime_baseline <- hms(S03_RC$S03_PSG_StartTime_baseline)
#S03_RC$S03_PSG_EndTime_baseline <- hms(S03_RC$S03_PSG_EndTime_baseline)
S03_RC$S03_PSG_OxStats_Diagnostic_MeanSaO2_PCT_Wake[S03_RC$S03_PSG_OxStats_Diagnostic_MeanSaO2_PCT_Wake==1] <- NA
S03_RC$S03_PSG_OxStats_Diagnostic_SpO2LessThan89_PCT_Minutes_TST[S03_RC$S03_PSG_OxStats_Diagnostic_SpO2LessThan89_PCT_Minutes_TST==1] <- NA
S03_RC$S03_PSG_OxStats_Diagnostic_SpO2LessThan88_PCT_Minutes_TST[S03_RC$S03_PSG_OxStats_Diagnostic_SpO2LessThan88_PCT_Minutes_TST==1] <- NA
S03_RC$S03_PSG_OxStats_Diagnostic_MinSaO2_PCT_TST[S03_RC$S03_PSG_OxStats_Diagnostic_MinSaO2_PCT_TST==1] <- NA
S03_RC$S03_PSG_OxStats_Diagnostic_MeanSaO2_PCT_TST[S03_RC$S03_PSG_OxStats_Diagnostic_MeanSaO2_PCT_TST==1] <- NA
S03_RC$S03_PSG_NPBMPatientInfoBMI[S03_RC$S03_PSG_NPBMPatientInfoBMI<1] <- NA
S03_RC$S03_PSG_NPBMPatientInfoHeight[S03_RC$S03_PSG_NPBMPatientInfoHeight==0] <- NA
S03_RC$S03_PSG_NPBMPatientInfoHeight[S03_RC$S03_PSG_NPBMPatientInfoHeight>110] <- NA
S03_RC$S03_PSG_NPBMPatientInfoWeight[S03_RC$S03_PSG_NPBMPatientInfoWeight==0] <- NA


S03_label_df <- data.frame(S03=labels_REDCap$S03[!is.na(labels_REDCap$S03)], rc=labels_REDCap$REDCap_label[!is.na(labels_REDCap$S03)], stringsAsFactors = F)

REDCap_S03_df <- data.frame(id=paste0("S03_", sprintf("%08d", 1:nrow(S03_RC))), stringsAsFactors = F)
for (v in colnames(REDCap_template)) {
        
        REDCap_S03_df[,v] <- S03_RC[,S03_label_df$S03[S03_label_df$rc==v]]
        
}


##############
########## H08
##############

H08_RC <- extracted_sleep_data$PennSleepDatabase_HSAT08_20190816111435.Rdata %>%
        select(labels_REDCap$H08_label[!is.na(labels_REDCap$H08_label)])

H08_RC$H08_PSG_PHI_Filename <- map_chr(strsplit(H08_RC$H08_PSG_PHI_Filename, "/"), pluck(6))
H08_RC$H08_PSG_PHI_StudyDate <- mdy(H08_RC$H08_PSG_PHI_StudyDate)
H08_RC$H08_PSG_Supine_AHI_index[H08_RC$H08_PSG_Supine_AHI_index<0] <- NA
H08_RC$H08_PSG_NonSupine_AHI_index[H08_RC$H08_PSG_NonSupine_AHI_index<0] <- NA
H08_RC$H08_PSG_LowestSpO2[H08_RC$H08_PSG_LowestSpO2<0] <- NA
H08_RC$H08_PSG_MeanSpO2[H08_RC$H08_PSG_MeanSpO2<0] <- NA
H08_RC$H08_PSG_PatientGender[grepl("female", H08_RC$H08_PSG_PatientGender)] <- "Female"
H08_RC$H08_PSG_PatientBMI[H08_RC$H08_PSG_PatientBMI<1] <- NA
H08_RC$H08_PSG_PatientHeight_value <- NA # unreliable strings
H08_RC$H08_PSG_PatientWeight_value[H08_RC$H08_PSG_PatientWeight_value==0] <- NA

H08_label_df <- data.frame(H08=labels_REDCap$H08[!is.na(labels_REDCap$H08)], rc=labels_REDCap$REDCap_label[!is.na(labels_REDCap$H08)], stringsAsFactors = F)

REDCap_H08_df <- data.frame(id=paste0("H08_", sprintf("%08d", 1:nrow(H08_RC))), stringsAsFactors = F)
for (v in colnames(REDCap_template)) {
        
        REDCap_H08_df[,v] <- H08_RC[,H08_label_df$H08[H08_label_df$rc==v]]
        
}

##### Combine all into one REDCap dataset
REDCap_loaded <- bind_rows(REDCap_template,REDCap_B01_df, REDCap_B02_df, REDCap_B03_df, REDCap_S01_df, REDCap_S02_df, REDCap_S03_df, REDCap_H08_df)

saveRDS(REDCap_loaded, "REDCap_PennSleepDatabase_preFormat_08212019.Rdata")

# Create a copy
REDCap_PennSleep <- REDCap_loaded

### Create master key
PennSleepDatabase_IDs <- REDCap_PennSleep[,c("study_id", "id")]
colnames(PennSleepDatabase_IDs) <- c("filename", "PennStudyID") # PennStudyID is per study, not per individual
# Bring MRN
PennSleepMetadata_wGAP_wPMBB <- readRDS("PennSleepMetadata_wGAP_wPMBB_081919.Rdata")
PennSleepDatabase_IDs <- merge(PennSleepDatabase_IDs, PennSleepMetadata_wGAP_wPMBB[,c("file_name", "SleepStudy_MRN")], by.x = "filename", by.y = "file_name", all.x=T, sort=F)
PennSleepDatabase_IDs <- distinct(PennSleepDatabase_IDs)



# Create Unique list of PennSubjectID form MRNs
#set.seed(12345)
#MRN_PennSubjectID_df <- data.frame(SleepStudy_MRN=unique(PennSleepDatabase_IDs$SleepStudy_MRN), PennSubjectID=sample(paste0("PENNSLEEP_", sprintf("%08d", 1:length(unique(PennSleepDatabase_IDs$SleepStudy_MRN))))), stringsAsFactors = F)
#saveRDS(MRN_PennSubjectID_df, "MRN_PennSubjectID_df_CreatedOn08222019.Rdata")
MRN_PennSubjectID_df <- readRDS("MRN_PennSubjectID_df_CreatedOn08222019.Rdata")



REDCap_PennSleep <- merge(REDCap_PennSleep, PennSleepDatabase_IDs, by.x="study_id", by.y = "filename", all.x=T, sort=F)
REDCap_PennSleep <- merge(REDCap_PennSleep, MRN_PennSubjectID_df, by.x="SleepStudy_MRN", by.y = "SleepStudy_MRN", all.x=T, sort=F)
REDCap_PennSleep$id <- NULL

REDCap_PennSleep$study_id <- PennSleepDatabase_IDs$REDCAP_study_id
REDCap_PennSleep$ptid <- PennSleepDatabase_IDs$PennInternalID


#### Create large database first with longitudinal feature on REDCap and load the template here
REDCap_Longitudinal_Template <- read.csv("UniversityOfPennsylvaniaPennSl_ImportTemplate_2019-08-22.csv", stringsAsFactors = F)
REDCap_PennSleep <- bind_rows(REDCap_Longitudinal_Template, REDCap_PennSleep)

### Fix IDs
REDCap_PennSleep$study_id <- REDCap_PennSleep$PennSubjectID
REDCap_PennSleep$ptid <- REDCap_PennSleep$PennSubjectID
REDCap_PennSleep$sleep_study_id <- REDCap_PennSleep$PennStudyID
REDCap_PennSleep$SleepStudy_MRN <- NULL
REDCap_PennSleep$PennStudyID <- NULL
REDCap_PennSleep$PennSubjectID <- NULL

### Ensure categorical representations are matching appropritely to REDCap data dictionary
REDCap_PennSleep$site <- 5 # University of Pennsylvania
REDCap_PennSleep$sleep_study_avail <- 1 # Yes
REDCap_PennSleep$sleep_study_date <- format(ymd(REDCap_PennSleep$sleep_study_date), "%Y-%m-%d") # converted to specified format in REDcap
REDCap_PennSleep$sleep_study_type[REDCap_loaded$sleep_study_type=="Baseline Study" | REDCap_loaded$sleep_study_type=="Split Night Study"] <- 1 # 1=In-laboratory Polysomnography (PSG)
REDCap_PennSleep$sleep_study_type[REDCap_loaded$sleep_study_type=="HSAT" ] <- 2 # 2=Home Sleep Test (HST)
REDCap_PennSleep$sleep_study_type <- as.numeric(REDCap_PennSleep$sleep_study_type)
REDCap_PennSleep$edf_avail <- 2 # 2=Unknown
REDCap_PennSleep$psg_study_type[REDCap_loaded$psg_study_type=="Baseline Study" | REDCap_loaded$psg_study_type=="HSAT"] <- 1 #Whole (Full Diagnostic Study)
REDCap_PennSleep$psg_study_type[REDCap_loaded$psg_study_type=="Split Night Study"] <- 2 # 2=Split (Diagnostic and Titration)
REDCap_PennSleep$psg_study_type <- as.numeric(REDCap_PennSleep$psg_study_type)
REDCap_PennSleep$scoring_method <- 4 # Other
REDCap_PennSleep$scoring_method_other <- "Unknown" # Other
REDCap_PennSleep$supp_o2_used <- 2 # Unknown
REDCap_PennSleep$analysis_start_time <- format(parse_date_time(x = REDCap_PennSleep$analysis_start_time, orders = c("H:M:S", "I:M:S %p")), "%H:%M")
REDCap_PennSleep$analysis_stop_time <- format(parse_date_time(x = REDCap_PennSleep$analysis_stop_time, orders = c("H:M:S", "I:M:S %p")), "%H:%M")
REDCap_PennSleep$gender[REDCap_loaded$gender=="Male"] <- 1
REDCap_PennSleep$gender[REDCap_loaded$gender=="Female"] <- 2
REDCap_PennSleep$gender[REDCap_loaded$gender=="Unknown"] <- 3 # Other
REDCap_PennSleep$gender <- as.numeric(REDCap_PennSleep$gender)
REDCap_PennSleep$gender_other_desc[REDCap_loaded$gender=="Unknown"] <- "Unknown"
REDCap_PennSleep$sbp_at_study <-round(REDCap_PennSleep$sbp_at_study)
REDCap_PennSleep$dbp_at_study <-round(REDCap_PennSleep$dbp_at_study)
REDCap_PennSleep$sleep_study_data_collection_form_complete <- 1 # Unverified

# Add redcap_event_name
# Sort by date of study
REDCap_PennSleep <- REDCap_PennSleep %>%
        mutate(date = mdy(sleep_study_date)) %>%
        group_by(study_id) %>%
        arrange(study_id, date)
REDCap_PennSleep$date <- NULL


# Remove studies with repeated dates
REDCap_PennSleep <- REDCap_PennSleep %>%
        group_by(study_id) %>%
        distinct(sleep_study_date, .keep_all = T)

# Create order (redcap_event_name)
REDCap_PennSleep <- REDCap_PennSleep %>%
        group_by(study_id) %>%
        mutate(event=paste0("visit_",1:n(),"_arm_1"))
REDCap_PennSleep$redcap_event_name <- REDCap_PennSleep$event
REDCap_PennSleep$event <- NULL

#############################################################################
#############################################################################

##### Save REDCap Ready ALL (from 2002 to 2018)
timestring <- gsub("-|\ |:", "\\1", Sys.time())
saveRDS(REDCap_PennSleep, paste0("PennSleepDatabase_REDCapReady_20021019to20181015_All_",timestring,".Rdata"))

# Split in 10000 file chuncks chunks (file too big)
split_ids <- split(1:nrow(REDCap_PennSleep), ceiling(seq_along(1:nrow(REDCap_PennSleep))/10000))

for (id_id in 1:length(split_ids)) {
        
        write.csv(REDCap_PennSleep[split_ids[[id_id]],], paste0("PennSleepDatabase_REDCapReady_20021019to20181015_All_",timestring,"_part",names(split_ids)[id_id],".csv"), row.names = F, na = "")
        
}
write.csv(REDCap_PennSleep, "~/PennSleepDatabase_deliverables/PennSleepDatabase_REDCapReady_20021019to20181015_All_08222019.csv", row.names = F, na = "")


#############################################################################
#############################################################################

###### Filter SAGS data only
# Including only the oldest sleep study, regardless whether th

REDCap_PennSleep_SAGS <- REDCap_PennSleep

#### Filter to accomodate SAGS
SleepStudy_GAP_PMBB_Summary_wID <- merge(SleepStudy_GAP_PMBB_Summary, MRN_PennSubjectID_df, by="SleepStudy_MRN", all.x=T)
SleepStudy_GAP_PMBB_Summary_wID$IncludeSAGS <- !is.na(SleepStudy_GAP_PMBB_Summary_wID$PennSubjectID) & (SleepStudy_GAP_PMBB_Summary_wID$PMBB_consented | SleepStudy_GAP_PMBB_Summary_wID$GAP_withDNA)

# Remove out of range, only get first record visit, exclude the variables redcap_event_name and sleep_study_id and filter only those with PMBB consented and GAP with DNA
REDCap_PennSleep_SAGS <- REDCap_PennSleep_SAGS %>%
        filter(age_at_study>=18, age_at_study<=88) %>%
        filter(redcap_event_name=="visit_1_arm_1") %>% # this represents the earliest study
        select(-redcap_event_name, -sleep_study_id) %>%
        filter(study_id %in% SleepStudy_GAP_PMBB_Summary_wID$PennSubjectID[SleepStudy_GAP_PMBB_Summary_wID$IncludeSAGS])

write.csv(REDCap_PennSleep_SAGS, "PennSAGSDatabase_REDCapReady_PMBBconsented_GAPwithDNA_08222019.csv", row.names = F, na = "")
write.csv(REDCap_PennSleep_SAGS, "~/PennSleepDatabase_deliverables/PennSAGSDatabase_REDCapReady_PMBBconsented_GAPwithDNA_08222019.csv", row.names = F, na = "")

##### Create Master ID Key
PennSleepDatabaseKEY <- merge(merge(PennSleepDatabase_IDs, MRN_PennSubjectID_df, by = "SleepStudy_MRN", all.x=T), SleepStudy_GAP_PMBB_Summary_wID, by = "SleepStudy_MRN", all.x=T)
PennSleepDatabaseKEY$PennSubjectID.y <- NULL
colnames(PennSleepDatabaseKEY)[colnames(PennSleepDatabaseKEY)=="PennSubjectID.x"] <- "PennSubjectID"
write.csv(PennSleepDatabaseKEY, "~/PennSleepDatabase_deliverables/PennSleepDatabase_IDKEY_08222019.csv", row.names = F)

### NEXT:

### Harmonize ALL variables across studies and create an universal data dictionary (see Spout: https://github.com/nsrr/spout)



