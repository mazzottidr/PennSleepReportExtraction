library(xml2)
library(xlsx)
library(dplyr)
#library(cyphr)
library(lubridate)
library(textreadr)
library(docxtractr)



############### Function definitions ##################

# Define function to get tables from word documents - or use function from package docxtractr
get_tbls <- function(word_doc) {
        
        tmpd <- tempdir()
        tmpf <- tempfile(tmpdir=tmpd, fileext=".zip")
        
        file.copy(word_doc, tmpf)
        unzip(tmpf, exdir=sprintf("%s/docdata", tmpd))
        
        doc <- read_xml(sprintf("%s/docdata/word/document.xml", tmpd))
        
        unlink(tmpf)
        unlink(sprintf("%s/docdata", tmpd), recursive=TRUE)
        
        ns <- xml_ns(doc)
        
        tbls <- xml_find_all(doc, ".//w:tbl", ns=ns)
        
        lapply(tbls, function(tbl) {
                
                cells <- xml_find_all(tbl, "./w:tr/w:tc", ns=ns)
                rows <- xml_find_all(tbl, "./w:tr", ns=ns)
                dat <- data.frame(matrix(xml_text(cells), 
                                         ncol=(length(cells)/length(rows)), 
                                         byrow=TRUE), 
                                  stringsAsFactors=FALSE)
                colnames(dat) <- dat[1,]
                dat <- dat[-1,]
                rownames(dat) <- NULL
                dat
                
        })
        
}


#Create trim function
trim <- function (x) gsub("^\\s+|\\s+$", "", x)

#Function to find fields
find_field <- function(field, table) {
        
        #For testing
        #field = "NPBMPatientInfoLastName"
        #table = TABLES
        
        
        tryCatch({
                tb_idx <- grep(field, table)[length(grep(field, table))] #Assuming the last table will contain what is needed
                
                tb <- rbind(colnames(table[[tb_idx]]), table[[tb_idx]])
                colnames(tb) <- NULL
                
                tb_col_idx <- grep(field, tb)
                tb_col <-  unlist(tb[,tb_col_idx])
                tb_field <- grep(field, tb_col, value = T)
                
                
                num_fields <- length(strsplit(as.character(tb_field), "MERGEFORMAT ")[[1]])-1
                
                if (num_fields==2) {
                        result <- trim(paste(strsplit(strsplit(as.character(tb_field), "MERGEFORMAT ")[[1]],
                                                      "  DOC")[[2]][1], strsplit(strsplit(as.character(tb_field), "MERGEFORMAT ")[[1]],"  DOC")[[3]][1]))
                } else if (num_fields==1) {
                        
                        result <- trim(strsplit(as.character(tb_field), "MERGEFORMAT ")[[1]][2])
                        
                } else if (num_fields==0) {
                        
                        if (!grepl("^DOCPROPERTY", tb_field)) { # if does not start with DOCPROPERTY, get very first field
                                result <- trim(strsplit(strsplit(as.character(tb_field), "MERGEFORMAT ")[[1]], " ")[[1]][1])
                        } else {
                                result <- NA
                        }
                        
                        
                }
                
                return(result)
        }, error=function(e) NA)

}

get_final_diagnosis <- function(table) {
        
        tb_idx <- grep("FINAL DIAGNOSIS:", table)
        tb_col_idx <- grep("FINAL DIAGNOSIS:", table[[tb_idx]])
        result <- gsub(" IF(.+)$", "",grep("FINAL DIAGNOSIS:", table[[tb_idx]][,tb_col_idx], value = T))
        return(result)
        
}

get_comm_rec <- function(table) {
        
        tb_idx <- grep("COMMENTS", table)
        tb_col_idx <- grep("COMMENTS", table[[tb_idx]])
        result <-  gsub(' DATE \\\\@ \\"M/d/yyyy\\"', "", gsub(' TIME \\\\@ \\"h:mm am/pm\\"', "", gsub(' DATE \\\\@ \\"M/d/yyyy\\"', "",grep("COMMENTS", table[[tb_idx]][,tb_col_idx], value = T))))
        return(result)
        
}

#### Define fields from word tables to extract
word_fields <- c(
        "NPBMPatientInfoLastName",
        "NPBMPatientInfoFirstName",
        "NPBMSiteInfoHospitalNumber",
        "NPBMStudyInfoStudyDate",
        "NPBMCustomPatientInfoSleepLabLocation",
        "NPBMCustomPatientInfoReferring_Physician",
        "NPBMCustomPatientInfoSleep_Specialist",
        "NPBMCustomPatientInfoBedNumber",
        "NPBMCustomPatientInfoDateScored",
        "NPBMPatientInfoDOB",
        "NPBMPatientInfoHeight",
        "NPBPatientInfoLengthUnit",
        "NPBMPatientInfoWeight",
        "NPBPatientInfoWeightUnit",
        "NPBMPatientInfoBMI",
        "NPBMPatientInfoGender",
        "NPBMCustomPatientInfoRecordingTechnician",
        "NPBMCustomPatientInfoScoringTechnician",
        "VAR_LightsOff",
        "VAR_LightsOn",
        "VAR_TotalRecordingTime_Minutes",
        "VAR_TotalSleepTime_Minutes",
        "VAR_SleepEfficiency_Percent",
        "VAR_SleepLatency_Minutes",
        "VAR_REMPeriodCount",
        "VAR_WakeAfterSleepTime_Minutes",
        "VAR_StageN1Time_Minutes",
        "VAR_StageN2Time_Minutes",
        "VAR_StageN3Time_Minutes",
        "VAR_REMTime_Minutes",
        "VAR_REMLatency_Minutes",
        "VAR_StageN1Time_PCT_TST",
        "VAR_StageN2Time_PCT_TST",
        "VAR_StageN3Time_PCT_TST",
        "VAR_REMTime_PCT_TST",
        "VAR_Custom_TotalArousalsIndex_TST",
        "VAR_Custom_TotalArousalCount_TST",
        "VAR_Custom_TotalArousalsIndex_NREM",
        "VAR_TotalArousalCount_NREM",
        "VAR_Custom_LimbMovementIndex_TST",
        "VAR_Custom_LimbMovementCount_TST",
        "VAR_Custom_LimbMovementIndex_NREM",
        "VAR_Custom_LimbMovementCount_NREM",
        "VAR_Custom_LimbMovementIndex_REM",
        "VAR_Custom_LimbMovementCount_REM",
        "VAR_Custom_LimbMovementArousalsIndex_TST",
        "VAR_Custom_LimbMovementArousalCount_TST",
        "VAR_Custom_LimbMovementArousalsIndex_NREM",
        "VAR_Custom_LimbMovementArousalCount_NREM",
        "VAR_Custom_LimbMovementArousalsIndex_REM",
        "VAR_Custom_LimbMovementArousalCount_REM",
        "NPBMCustomPatientInfoCardiacArrhythmias",
        "NPBMCustomPatientInfoBaselineSupine", #Snoring supine
        "NPBMCustomPatientInfoBaselineLateral", #Snoring lateral
        "NPBMCustomPatientInfoBaselineProne" #Snoring prone
)

extract_baseline_3 <- function(paths) {
        
        
        #Get all .docx files
        PSGfiles <- paths
        
        #Add first logical
        first <- T
        
        #Process main extracting script
        i=1
        for (PSGfilename in PSGfiles) {
                
                #PSGfilename <- PSGfiles[1]
                
                message(paste0("Processing ", PSGfilename))
                
                tryCatch(
                        {
                            # Deal with temporaty folders and copying docx to local folder before processing
                            # Copy and rename docx before processsing to avoid error with file names
                            dir.create("temp_reports_b03/", showWarnings = FALSE)
                            file.copy(PSGfilename, to = "temp_reports_b03/")
                            
                            docx <- paste0("temp_reports_b03/", sapply(strsplit(PSGfilename, "/"), "[[", 6))
                            
                            
                            #current_tmp_files <- list.files(tempdir(), full.names = T)
                            
                            message(paste0("Processing: ", sapply(strsplit(PSGfilename, "/"), "[[", 6)))
                            message(paste0(format(100*i/length(PSGfiles), digits = 4),"% done,"))
                            
                            #Get all possible tables
                                #TABLES <-get_tbls(PSGfilename) #This gives a warning
                                TABLES <-sapply(1:docx_tbl_count(docxtractr::read_docx(docx)), docx_extract_tbl, docx=docxtractr::read_docx(docx), header=F)
                                
                                #Process function to extract fields from word tables
                                word_main_values_df <- data.frame(t(sapply(word_fields, find_field, table=TABLES)), stringsAsFactors = F)
                                
                                #Fix fields that need to be fixed and convert to numeric
                                word_main_values_df$NPBMPatientInfoLastName <- strsplit(word_main_values_df$NPBMPatientInfoLastName,", ")[[1]][1]
                                word_main_values_df$NPBMPatientInfoFirstName <- strsplit(word_main_values_df$NPBMPatientInfoFirstName,", ")[[1]][2]
                                word_main_values_df$NPBMPatientInfoHeight <- strsplit(word_main_values_df$NPBMPatientInfoHeight," ")[[1]][1]
                                word_main_values_df$NPBPatientInfoLengthUnit <- strsplit(word_main_values_df$NPBPatientInfoLengthUnit," ")[[1]][2]
                                word_main_values_df$NPBMPatientInfoWeight <- strsplit(word_main_values_df$NPBMPatientInfoWeight," ")[[1]][1]
                                word_main_values_df$NPBPatientInfoWeightUnit <- strsplit(word_main_values_df$NPBPatientInfoWeightUnit," ")[[1]][2]
                                word_main_values_df$NPBMPatientInfoBMI <- trim(gsub("kg/m2", "", word_main_values_df$NPBMPatientInfoBMI))
                                word_main_values_df$VAR_SleepEfficiency_Percent <- gsub("%", "", word_main_values_df$VAR_SleepEfficiency_Percent)
                                word_main_values_df$VAR_StageN1Time_PCT_TST <- gsub("%", "",gsub(")", "", word_main_values_df$VAR_StageN1Time_PCT_TST))
                                word_main_values_df$VAR_StageN2Time_PCT_TST <- gsub("%", "",gsub(")", "", word_main_values_df$VAR_StageN2Time_PCT_TST))
                                word_main_values_df$VAR_StageN3Time_PCT_TST <- gsub("%", "",gsub(")", "", word_main_values_df$VAR_StageN3Time_PCT_TST))
                                word_main_values_df$VAR_REMTime_PCT_TST <- gsub("%", "",gsub(")", "", word_main_values_df$VAR_REMTime_PCT_TST))
                                
                                #Fix as numeric
                                word_main_values_df[,-(which(colnames(word_main_values_df) %in% c("NPBMPatientInfoLastName", "NPBMPatientInfoFirstName","NPBMSiteInfoHospitalNumber","NPBMStudyInfoStudyDate","NPBMCustomPatientInfoSleepLabLocation","NPBMCustomPatientInfoReferring_Physician","NPBMCustomPatientInfoSleep_Specialist","NPBMCustomPatientInfoBedNumber", "NPBMCustomPatientInfoDateScored","NPBMPatientInfoDOB", "NPBPatientInfoLengthUnit","NPBPatientInfoWeightUnit", "NPBMPatientInfoGender", "NPBMCustomPatientInfoRecordingTechnician", "NPBMCustomPatientInfoScoringTechnician", "VAR_LightsOff", "VAR_LightsOn","NPBMCustomPatientInfoCardiacArrhythmias", "NPBMCustomPatientInfoBaselineSupine", "NPBMCustomPatientInfoBaselineLateral", "NPBMCustomPatientInfoBaselineProne")))] <- sapply(word_main_values_df[,-(which(colnames(word_main_values_df) %in% c("NPBMPatientInfoLastName", "NPBMPatientInfoFirstName","NPBMSiteInfoHospitalNumber","NPBMStudyInfoStudyDate","NPBMCustomPatientInfoSleepLabLocation","NPBMCustomPatientInfoReferring_Physician","NPBMCustomPatientInfoSleep_Specialist","NPBMCustomPatientInfoBedNumber", "NPBMCustomPatientInfoDateScored","NPBMPatientInfoDOB", "NPBPatientInfoLengthUnit","NPBPatientInfoWeightUnit", "NPBMPatientInfoGender", "NPBMCustomPatientInfoRecordingTechnician", "NPBMCustomPatientInfoScoringTechnician", "VAR_LightsOff", "VAR_LightsOn","NPBMCustomPatientInfoCardiacArrhythmias", "NPBMCustomPatientInfoBaselineSupine", "NPBMCustomPatientInfoBaselineLateral", "NPBMCustomPatientInfoBaselineProne")))], as.numeric)
                                
                                
                                #Fix columns names
                                colnames(word_main_values_df)[c(1:3,5:10,17,18)] <- paste0("B03_PSG_PHI_",colnames(word_main_values_df)[c(1:3,5:10,17,18)])
                                colnames(word_main_values_df)[-c(1:3,5:10,17,18)] <- paste0("B03_PSG_",colnames(word_main_values_df)[-c(1:3,5:10,17,18)])
                                
                                # Get other relevant values from the word document
                                word_other_values <- c(B03_PSG_study_type="Baseline Study",
                                                       B03_PSG_PHI_clinical_history=gsub(" The patient(.*?)with a history of ", "", grep("CLINICAL HISTORY:", textreadr::read_docx(PSGfilename), value = T)),
                                                       B03_PSG_PHI_clinical_history=gsub(" The patient(.*?)with a history of ", "", grep("CLINICAL HISTORY:", textreadr::read_docx(PSGfilename), value = T)),
                                                       B03_PSG_PHI_final_diagnosis=get_final_diagnosis(TABLES),
                                                       B03_PSG_PHI_comm_rec=get_comm_rec(TABLES)
                                                       
                                )
                                
                                word_other_values_df <-  data.frame(t(word_other_values), stringsAsFactors = F)
                                
                                ### Get fields from Excel embedded tables that appear in the report
                                xl_tb <- PSGfilename
                                
                                #Copy file ################## Fix ths to be universal
                                dir.create("zips_b03/", showWarnings = F)
                                file.copy(from = PSGfilename, to = "zips_b03/")
                                file.rename(paste0("zips_b03/", strsplit(PSGfilename, "/")[[1]][6]), "zips_b03/current.zip")
                                unzip("zips_b03/current.zip", exdir = paste0(getwd(),"/zips_b03"))
                                xls_tables_paths <- grep("Worksheet", dir("zips_b03/word/embeddings/", full.names = T), value = T)
                                
                                
                                # Iterate to load all possible tables
                                xtblist <- NULL
                                
                                for (xt in xls_tables_paths) {
                                        xtblist[[xt]]  <- read.xlsx(xt, sheetIndex = 1)
                                }
                                
                                
                                # Try getting position time in same spreadsheet
                                
                                
                                if (length(which(grepl("Position", xtblist)))==1) {
                                        
                                        xlstb1 <- xtblist[[which(grepl("Position", xtblist))]][,c(8,10,11)] #Position time
                                        
                                } else {
                                        
                                        # Otherwise, in separated
                                        xlstb1 <- xtblist[[which(grepl("Position", xtblist) &  grepl("Time..min.", xtblist))]] #Position time
                                        
                                }
                                
                                
                                
                                
                                xlstb2 <- xtblist[[which(grepl("Respiratory.Events.", xtblist))[1]]] #Respiratory events
                                xlstb3 <- xtblist[[which(grepl("Oxygen.Saturation", xtblist))[1]]] #Oxygen saturation
                                xlstb4 <- xtblist[[which(grepl("Pulse.Rate", xtblist))[1]]] #Pulse rate
                                xlstb5 <- xtblist[[which(grepl("Oxygen.Desats", xtblist))[1]]] #Oxygen desats
                                
                                xls_vars <- data.frame(
                                        
                                        #Position time
                                        B03_PSG_supine_t=as.numeric(as.character(xlstb1[1,2])),
                                        B03_PSG_supine_p=as.numeric(as.character(gsub(")","",gsub("(","",as.character(xlstb1[1,3]), fixed = T), fixed = T))),
                                        B03_PSG_prone_t=as.numeric(as.character(xlstb1[2,2])),
                                        B03_PSG_prone_p=as.numeric(as.character(gsub(")","",gsub("(","",as.character(xlstb1[2,3]), fixed = T), fixed = T))),
                                        B03_PSG_left_t=as.numeric(as.character(xlstb1[3,2])),
                                        B03_PSG_left_p=as.numeric(as.character(gsub(")","",gsub("(","",as.character(xlstb1[3,3]), fixed = T), fixed = T))),
                                        B03_PSG_right_t=as.numeric(as.character(xlstb1[4,2])),
                                        B03_PSG_right_p=as.numeric(as.character(gsub(")","",gsub("(","",as.character(xlstb1[4,3]), fixed = T), fixed = T))),
                                        
                                        #Respiratory events
                                        B03_PSG_cent_ap_count=as.numeric(as.character(xlstb2[1,2])),
                                        B03_PSG_mxd_ap_count=as.numeric(as.character(xlstb2[1,3])),
                                        B03_PSG_obs_ap_count=as.numeric(as.character(xlstb2[1,4])),
                                        B03_PSG_total_ap_count=as.numeric(as.character(xlstb2[1,5])),
                                        B03_PSG_total_hyp_count=as.numeric(as.character(xlstb2[1,6])),
                                        B03_PSG_total_aphyp_count=as.numeric(as.character(xlstb2[1,7])),
                                        
                                        B03_PSG_cent_ap_idx=as.numeric(as.character(xlstb2[2,2])),
                                        B03_PSG_mxd_ap_idx=as.numeric(as.character(xlstb2[2,3])),
                                        B03_PSG_obs_ap_idx=as.numeric(as.character(xlstb2[2,4])),
                                        B03_PSG_total_ap_idx=as.numeric(as.character(xlstb2[2,5])),
                                        B03_PSG_total_hyp_idx=as.numeric(as.character(xlstb2[2,6])),
                                        B03_PSG_total_aphyp_idx=as.numeric(as.character(xlstb2[2,7])),
                                        
                                        B03_PSG_cent_ap_remidx=as.numeric(as.character(xlstb2[3,2])),
                                        B03_PSG_mxd_ap_remidx=as.numeric(as.character(xlstb2[3,3])),
                                        B03_PSG_obs_ap_remidx=as.numeric(as.character(xlstb2[3,4])),
                                        B03_PSG_total_ap_remidx=as.numeric(as.character(xlstb2[3,5])),
                                        B03_PSG_total_hyp_remidx=as.numeric(as.character(xlstb2[3,6])),
                                        B03_PSG_total_aphyp_remidx=as.numeric(as.character(xlstb2[3,7])),
                                        
                                        B03_PSG_cent_ap_nremidx=as.numeric(as.character(xlstb2[4,2])),
                                        B03_PSG_mxd_ap_nremidx=as.numeric(as.character(xlstb2[4,3])),
                                        B03_PSG_obs_ap_nremidx=as.numeric(as.character(xlstb2[4,4])),
                                        B03_PSG_total_ap_nremidx=as.numeric(as.character(xlstb2[4,5])),
                                        B03_PSG_total_hyp_nremidx=as.numeric(as.character(xlstb2[4,6])),
                                        B03_PSG_total_aphyp_nremidx=as.numeric(as.character(xlstb2[4,7])),
                                        
                                        B03_PSG_supine_total_aphyp_idx=as.numeric(strsplit(as.character(xlstb2[1,9]), " (", fixed = T)[[1]][1]),
                                        B03_PSG_supine_total_aphyp_count=as.numeric(gsub(")", "", strsplit(as.character(xlstb2[1,9]), " (", fixed = T)[[1]][2], fixed = T)),
                                        
                                        B03_PSG_prone_total_aphyp_idx=as.numeric(strsplit(as.character(xlstb2[2,9]), " (", fixed = T)[[1]][1]),
                                        B03_PSG_prone_total_aphyp_count=as.numeric(gsub(")", "", strsplit(as.character(xlstb2[2,9]), " (", fixed = T)[[1]][2], fixed = T)),
                                        
                                        B03_PSG_left_total_aphyp_idx=as.numeric(strsplit(as.character(xlstb2[3,9]), " (", fixed = T)[[1]][1]),
                                        B03_PSG_left_total_aphyp_count=as.numeric(gsub(")", "", strsplit(as.character(xlstb2[3,9]), " (", fixed = T)[[1]][2], fixed = T)),
                                        
                                        B03_PSG_right_total_aphyp_idx=as.numeric(strsplit(as.character(xlstb2[4,9]), " (", fixed = T)[[1]][1]),
                                        B03_PSG_right_total_aphyp_count=as.numeric(gsub(")", "", strsplit(as.character(xlstb2[4,9]), " (", fixed = T)[[1]][2], fixed = T)),
                                        
                                        
                                        #Oxygen saturation
                                        B03_PSG_meanSpO2_wake=as.numeric(as.character(xlstb3$Wake[1])),
                                        B03_PSG_minSpO2_wake=as.numeric(as.character(xlstb3$Wake[2])),
                                        B03_PSG_timeSpO2_LT88_wake=as.numeric(as.character(xlstb3$Wake[3])),
                                        B03_PSG_perctime_90.100_wake=as.numeric(as.character(xlstb3$Wake[5])),
                                        B03_PSG_perctime_80.89_wake=as.numeric(as.character(xlstb3$Wake[6])),
                                        B03_PSG_perctime_70.79_wake=as.numeric(as.character(xlstb3$Wake[7])),
                                        B03_PSG_perctime_badO2data_wake=as.numeric(as.character(xlstb3$Wake[8])),
                                        
                                        B03_PSG_meanSpO2_nrem=as.numeric(as.character(xlstb3$NREM[1])),
                                        B03_PSG_minSpO2_nrem=as.numeric(as.character(xlstb3$NREM[2])),
                                        B03_PSG_timeSpO2_LT88_nrem=as.numeric(as.character(xlstb3$NREM[3])),
                                        B03_PSG_perctime_90.100_nrem=as.numeric(as.character(xlstb3$NREM[5])),
                                        B03_PSG_perctime_80.89_nrem=as.numeric(as.character(xlstb3$NREM[6])),
                                        B03_PSG_perctime_70.79_nrem=as.numeric(as.character(xlstb3$NREM[7])),
                                        B03_PSG_perctime_badO2data_nrem=as.numeric(as.character(xlstb3$NREM[8])),
                                        
                                        B03_PSG_meanSpO2_rem=as.numeric(as.character(xlstb3$REM[1])),
                                        B03_PSG_minSpO2_rem=as.numeric(as.character(xlstb3$REM[2])),
                                        B03_PSG_timeSpO2_LT88_rem=as.numeric(as.character(xlstb3$REM[3])),
                                        B03_PSG_perctime_90.100_rem=as.numeric(as.character(xlstb3$REM[5])),
                                        B03_PSG_perctime_80.89_rem=as.numeric(as.character(xlstb3$REM[6])),
                                        B03_PSG_perctime_70.79_rem=as.numeric(as.character(xlstb3$REM[7])),
                                        B03_PSG_perctime_badO2data_rem=as.numeric(as.character(xlstb3$REM[8])),
                                        
                                        B03_PSG_meanSpO2_tst=as.numeric(as.character(xlstb3$TST[1])),
                                        B03_PSG_minSpO2_tst=as.numeric(as.character(xlstb3$TST[2])),
                                        B03_PSG_timeSpO2_LT88_tst=as.numeric(as.character(xlstb3$TST[3])),
                                        B03_PSG_perctime_90.100_tst=as.numeric(as.character(xlstb3$TST[5])),
                                        B03_PSG_perctime_80.89_tst=as.numeric(as.character(xlstb3$TST[6])),
                                        B03_PSG_perctime_70.79_tst=as.numeric(as.character(xlstb3$TST[7])),
                                        B03_PSG_perctime_badO2data_tst=as.numeric(as.character(xlstb3$TST[8])),
                                        
                                        B03_PSG_meanSpO2_tib=as.numeric(as.character(xlstb3$TIB[1])),
                                        B03_PSG_minSpO2_tib=as.numeric(as.character(xlstb3$TIB[2])),
                                        B03_PSG_timeSpO2_LT88_tib=as.numeric(as.character(xlstb3$TIB[3])),
                                        B03_PSG_perctime_90.100_tib=as.numeric(as.character(xlstb3$TIB[5])),
                                        B03_PSG_perctime_80.89_tib=as.numeric(as.character(xlstb3$TIB[6])),
                                        B03_PSG_perctime_70.79_tib=as.numeric(as.character(xlstb3$TIB[7])),
                                        B03_PSG_perctime_badO2data_tib=as.numeric(as.character(xlstb3$TIB[8])),
                                        
                                        #Pulse rate
                                        B03_PSG_maxHR_wake=as.numeric(as.character(xlstb4$Wake[1])),
                                        B03_PSG_meanHR_wake=as.numeric(as.character(xlstb4$Wake[2])),
                                        B03_PSG_minHR_wake=as.numeric(as.character(xlstb4$Wake[3])),
                                        
                                        B03_PSG_maxHR_nrem=as.numeric(as.character(xlstb4$NREM[1])),
                                        B03_PSG_meanHR_nrem=as.numeric(as.character(xlstb4$NREM[2])),
                                        B03_PSG_minHR_nrem=as.numeric(as.character(xlstb4$NREM[4])),
                                        
                                        B03_PSG_maxHR_rem=as.numeric(as.character(xlstb4$REM[1])),
                                        B03_PSG_meanHR_rem=as.numeric(as.character(xlstb4$REM[2])),
                                        B03_PSG_minHR_rem=as.numeric(as.character(xlstb4$REM[3])),
                                        
                                        B03_PSG_maxHR_tst=as.numeric(as.character(xlstb4$TST[1])),
                                        B03_PSG_meanHR_tst=as.numeric(as.character(xlstb4$TST[2])),
                                        B03_PSG_minHR_tst=as.numeric(as.character(xlstb4$TST[3])),
                                        
                                        #Oxygen desats
                                        B03_PSG_desats_tst_idx=as.numeric(strsplit(as.character(xlstb5$Index....Count.[1]), " (", fixed = T)[[1]][1]),
                                        B03_PSG_desats_tst_count=as.numeric(gsub(")","",gsub("(","",as.character(xlstb5$NA.[1]), fixed = T), fixed = T)),
                                        
                                        B03_PSG_desats_nrem_idx=as.numeric(strsplit(as.character(xlstb5$Index....Count.[2]), " (", fixed = T)[[1]][1]),
                                        B03_PSG_desats_nrem_count=as.numeric(gsub(")","",gsub("(","",as.character(xlstb5$NA.[2]), fixed = T), fixed = T)),
                                        
                                        B03_PSG_desats_rem_idx=as.numeric(strsplit(as.character(xlstb5$Index....Count.[3]), " (", fixed = T)[[1]][1]),
                                        B03_PSG_desats_rem_count=as.numeric(gsub(")","",gsub("(","",as.character(xlstb5$NA.[3]), fixed = T), fixed = T)),
                                        stringsAsFactors = F
                                        
                                )
                                
                                
                                
                                
                                
                                #Get other extra variables from sheet index 2 in Excel embedded tables
                                
                                xtblist2 <- NULL
                                
                                for (xt in xls_tables_paths) {
                                        xtblist2[[xt]]  <- read.xlsx(xt, sheetIndex = 2)
                                        
                                        if(ncol(xtblist2[[xt]])!=3) {next()}
                                        
                                        colnames(xtblist2[[xt]]) <- c("Var_name", "value","label")
                                        xtblist2[[xt]]$Var_name <- as.character(xtblist2[[xt]]$Var_name)
                                        xtblist2[[xt]]$label <- as.character(xtblist2[[xt]]$label)
                                }
                                
                                
                                #Combine all spreadsheets with 3 columns (contains relevant PSG data)
                                xtblist2 <- xtblist2[sapply(xtblist2, ncol)==3]
                                for (l_id in 1:length(xtblist2)) {
                                        
                                        xtblist2[[l_id]] <- as.data.frame(sapply((xtblist2[[l_id]]), as.character), stringsAsFactors=F)
                                        
                                        
                                }
                                
                                xlstb2_combined <- as.data.frame(do.call("rbind", xtblist2), stringsAsFactors=F)
                                rownames(xlstb2_combined) <- NULL
                                
                                
                                #Remove repeated rows
                                xlstb2_combined <- unique(xlstb2_combined)
                                xlstb_combined_extraonly <- xlstb2_combined
                                
                                #Filter based on curated list of extra variables - I checked all variables to make sure they are not repetitons
                                #curated_extra <- read.table("~/ConvertReports/curatedExtraVariablesFromxls_Mazzotti08132018.csv", header = F, sep = ",", stringsAsFactors = F)[,1]
                                #xlstb_combined_extraonly <- xlstb2_combined[xlstb2_combined$Var_name %in% curated_extra,]
                                #rownames(xlstb_combined_extraonly) <- NULL
                                
                                #Fix values that are formated as dates in Excel
                                process_excel_date <- function(x) {
                                        if (x=="1900-01-01 00:00:00") {
                                                return("1")
                                        } else {
                                                return(as.character(time_length(ymd_hms(x)-ymd_hms("1899-12-30 00:00:00"), "days")))
                                                
                                        }
                                }
                                
                                #Get unformatted dates
                                xlstb_combined_extraonly$value[grepl(":", xlstb_combined_extraonly$value)] <- as.vector(sapply(xlstb_combined_extraonly$value[grepl(":", xlstb_combined_extraonly$value)], process_excel_date))
                                
                                #Convert to numeric
                                xlstb_combined_extraonly$value <- as.numeric(as.character(xlstb_combined_extraonly$value))
                                
                                #create set of variables to include as extra
                                extrapsgdata <- data.frame(t(xlstb_combined_extraonly$value))
                                colnames(extrapsgdata) <- paste0("B03_PSG_EXTRA_",t(xlstb_combined_extraonly$Var_name))
                                
                                ######### Remove all files from zip_folder to avoid confusion when getting new data
                                file.remove(list.files("zips_b03/", include.dirs = F, full.names = T, recursive = T))
                                
                                ######## Combine ALL data
                                
                                
                                all_fields <- c("B03_PSG_PHI_Filename", "PennSleepID", colnames(word_main_values_df), colnames(word_other_values_df), colnames(xls_vars), colnames(extrapsgdata))
                                
                                #Initiate dataframe with results
                                
                                
                                
                                #If first, create final_df
                                if (first) {
                                        final_df <- data.frame(matrix(NA,1,length(all_fields)))
                                        colnames(final_df) <- all_fields
                                        
                                        final_df[1,] <- c(PSGfilename,
                                                          NA,
                                                          word_main_values_df[1,],
                                                          word_other_values_df[1,],
                                                          xls_vars[1,],
                                                          extrapsgdata[1,])
                                        first <- F
                                } else {
                                        
                                        next_final_df <- data.frame(matrix(NA,1,length(all_fields)))
                                        colnames(next_final_df) <- all_fields
                                        
                                        next_final_df[1,] <- c(PSGfilename,
                                                                NA,
                                                                word_main_values_df[1,],
                                                                word_other_values_df[1,],
                                                                xls_vars[1,],
                                                                extrapsgdata[1,])
                                        
                                        # Merge first with and next
                                        final_df <- dplyr::bind_rows(final_df, next_final_df)
                                        
                                }
                                
                                # Delete corresponding folder
                                # new_tmp_files <- list.files(tempdir(), full.names = T, recursive = T)
                                # tmp_files_toRemove <- new_tmp_files[!(new_tmp_files %in% current_tmp_files)]
                                # tmp_folder_toRemove <- unique(paste(sapply(strsplit(tmp_files_toRemove, "/"), "[[", 1), sapply(strsplit(tmp_files_toRemove, "/"), "[[", 2), sep = "/"))
                                # unlink(tmp_folder_toRemove, recursive = T)
                                 
                                # Delete file from working directory
                                file.remove(docx)  
                                
                                
                                
                        }, error=function(cond) {
                                message("Something went wrong.")
                                message("Here's the original error message:")
                                message(cond)
                                print(PSGfilename)
                                
                                
                        }, finally = print("Done")
                        
                )
                
                i=i+1
                
        }
        
        #Clean data
        #Replace blanks and N/A with NA
        final_df[final_df==""] <- NA
        final_df[final_df=="N/A"] <- NA
        
        #Calculate Age at Study
        final_df$B03_PSG_Age_at_Study <- interval(mdy(final_df$B03_PSG_PHI_NPBMPatientInfoDOB), mdy(final_df$B03_PSG_NPBMStudyInfoStudyDate)) %/% years(1)
        
        #Create identifiable data frame
        identifiable_df <- final_df
        identifiable_df$ProcessedDate <- Sys.Date()
        
        ProcessedTimeID=gsub("-","",gsub(":","",gsub(" ", "", Sys.time())))
        
        identifiable_df$ProcessedTimeID <- ProcessedTimeID
        
        # Per sample QC (not value filter)
        
        QC_df <- data.frame(B03_PSG_PHI_Filename=identifiable_df$B03_PSG_PHI_Filename,
                            # Missing both MRN fields
                            has_missingMRN=is.na(identifiable_df$B03_PSG_PHI_NPBMSiteInfoHospitalNumber),
                            # Gender not (Male, Female, Unknown)
                            has_missingGender=!(tolower(identifiable_df$B03_PSG_NPBMPatientInfoGender) %in% c("male", "female", "unknown")),
                            # Missing age
                            has_missingAge=is.na(identifiable_df$B03_PSG_Age_at_Study),
                            # Missing BMI
                            has_missingBMI=is.na(identifiable_df$B03_PSG_NPBMPatientInfoBMI)
        )
        
        identifiable_df$PerSample_QC <- NA
        identifiable_df$PerSample_QC[rowSums(QC_df[,2:5])>0] <- "FAIL"
        identifiable_df$PerSample_QC[rowSums(QC_df[,2:5])==0] <- "PASS"
        
        QC_df$PerSample_QC <- identifiable_df$PerSample_QC
        write.csv(QC_df, paste0("QC_dataframe_", ProcessedTimeID, ".csv"))
        
        #Save identifiable Rdata file encrypted
        #key <- key_sodium(sodium::keygen())
        saveRDS(identifiable_df, paste0("PennSleepDatabase_Baseline03_", ProcessedTimeID, ".Rdata"))
        #cyphr::encrypt_file("PennSleepDatabase.Rdata", key, "PennSleepDatabase.encrypted")
        
        #Save de-identified version in CSV files - implement a way to check in database if sample was processed and not use the same de-identified IDs
        #identifiable_df$PennSleepID <- paste0("PENNSLEEP00000",seq(1:nrow(identifiable_df)))
        deidentified_df <- select(identifiable_df, -starts_with("PSG_PHI_"))
        
        write.csv(deidentified_df, paste0("PennSleepDatabase_Baseline03_Deidentified",ProcessedTimeID,".csv"), row.names = F)
        
        return(identifiable_df)
        
}

