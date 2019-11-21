library(xml2)
library(xlsx)
library(dplyr)
#library(cyphr)
library(lubridate)
library(textreadr)
library(docxtractr)
library(stringr)



############### Function definitions ##################
#find field

find_field_grep <- function(pattern, table, type=c("col_next", "row_below"), tb.pos=1, n=1, pattern2=pattern, fixed=T) {
        
        
        #For testing
        #pattern = "^REM Events:"
        #table = TABLES
        #type="col_next"
        #tb.pos = 1
        #fixed=F
        #n=1 # this is the number of column or row movements
        
        
        tryCatch({
                
                tb.pos.id <- tb.pos
                
                #Find table
                tb_idx <- grep(pattern, table, fixed = fixed)[tb.pos.id]
                
                tb <- table[[tb_idx]]
                
                tb_col_idx <- grep(pattern2, tb, fixed = fixed)
                tb_col <-  tb[,tb_col_idx]
                tb_field <- grep(pattern2, pull(tb_col), value = T, fixed = fixed)
                
                # Check if still has Excel Tags
                
                
                #type col_next will get values corresponding to next columns
                
                if (type=="col_next") {
                        
                        field_idx_row <- grep(tb_field, pull(tb_col), fixed = fixed)
                        field_idx_col <- tb_col_idx+n
                        
                }
                
                if (type=="row_below") {
                        
                        field_idx_row <- grep(tb_field, tb_col, fixed = fixed)+n
                        field_idx_col <- tb_col_idx
                        
                }
                
                
                
                
                
                result <- tb[field_idx_row,field_idx_col]
                
                
                
                if (length(result)==0) {
                        
                        return(NA)
                        
                } else if (grepl("DOCPROPERTY", result)) {
                        
                        num_fields <- length(strsplit(as.character(result), "MERGEFORMAT ")[[1]])-1
                        
                        if (num_fields==2) {
                                result <- trim(paste(strsplit(strsplit(as.character(result), "MERGEFORMAT ")[[1]],
                                                              "  DOC")[[2]][1], strsplit(strsplit(as.character(result), "MERGEFORMAT ")[[1]],"  DOC")[[3]][1]))
                        } else if (num_fields==1) {
                                
                                result <- trim(strsplit(as.character(result), "MERGEFORMAT ")[[1]][2])
                                
                        } else if (num_fields==0) {
                                
                                if (!grepl("^DOCPROPERTY", result)) { # if does not start with DOCPROPERTY, get very first field
                                        result <- trim(strsplit(strsplit(as.character(result), "MERGEFORMAT ")[[1]], " ")[[1]][1])
                                } else {
                                        result <- NA
                                }
                                
                        }
                        
                        return(result)
                        
                        
                } else {
                        
                        return(pull(result))
                }
                
                
                
        }, error=function(e) NA)
        
}


#Create trim function
trim <- function (x) gsub("^\\s+|\\s+$", "", x)


#### Main function starts

extract_split_1 <- function(paths) {
        
        #paths <- s01_paths[1:10]
        #Get all .docx files
        PSGfiles <- paths
        
        #Add first logical
        first <- T
        
        #Process main extracting script
        i=1
        for (PSGfilename in PSGfiles) {
                
                #PSGfilename <- PSGfiles[1]
                
                tryCatch(
                        {
                            # Deal with temporaty folders and copying docx to local folder before processing
                            # Copy and rename docx before processsing to avoid error with file names
                            dir.create(file.path("temp_reports_s01/"), showWarnings = FALSE)
                            file.copy(PSGfilename, to = "temp_reports_s01/")
                            
                            docx <- paste0("temp_reports_s01/", sapply(strsplit(PSGfilename, "/"), "[[", 6))
                            
                            #current_tmp_files <- list.files(tempdir(), full.names = T)
                            
                            message(paste0("Processing: ", sapply(strsplit(PSGfilename, "/"), "[[", 6)))
                            message(paste0(format(100*i/length(PSGfiles), digits = 4),"% done,"))
                            
                            #Get all possible tables, preparing tables to get
                                TABLES <-sapply(1:docx_tbl_count(docxtractr::read_docx(docx)), docx_extract_tbl, docx=docxtractr::read_docx(docx), header=F)
                                current_docx_summary <- officer::docx_summary(officer::read_docx(docx))
                                PLM_TABLE_id_baseline=min(which(sapply(TABLES, dim)[1,]==4 & sapply(TABLES, dim)[2,]==7))
                                PLM_TABLE_id_treatment=max(which(sapply(TABLES, dim)[1,]==4 & sapply(TABLES, dim)[2,]==7))
                                Arousal_TABLE_id_baseline <- min(which(sapply(TABLES, dim)[1,]==4 & sapply(TABLES, dim)[2,]==3))
                                Arousal_TABLE_id_treatment <- max(which(sapply(TABLES, dim)[1,]==4 & sapply(TABLES, dim)[2,]==3))
                                
                                
                                # Initiate data.frame
                                
                                split_1_df <- data.frame(
                                        
                                        # Patient identifier tables
                                        S01_PHI_Patient_Name=find_field_grep(pattern = "Patient Name", table = TABLES, type = "col_next", tb.pos = 1),
                                        S01_PHI_MRN=find_field_grep(pattern = "MR #", table = TABLES, type = "col_next", tb.pos = 1),
                                        S01_PHI_Sleep_Center_ID=find_field_grep(pattern = "Sleep Center ID", table = TABLES, type = "col_next", tb.pos = "last"),
                                        S01_PSG_Study_date=find_field_grep(pattern = "Study Date", table = TABLES, type = "col_next", tb.pos = 1),
                                        S01_PSG_Sex=find_field_grep(pattern = "Sex", table = TABLES, type = "col_next", tb.pos = 1),
                                        S01_PHI_Date_of_birth=find_field_grep(pattern = "D.O.B.", table = TABLES, type = "col_next", tb.pos = 1),
                                        S01_PSG_Age=as.numeric(find_field_grep(pattern = "Age", table = TABLES, type = "col_next", tb.pos = 1)),
                                        S01_PSG_Blood_Pressure=find_field_grep(pattern = "BP", table = TABLES, type = "col_next", tb.pos = 1),
                                        S01_PSG_Height=find_field_grep(pattern = "Height", table = TABLES, type = "col_next", tb.pos = 1),
                                        S01_PSG_Weight=find_field_grep(pattern = "Weight", table = TABLES, type = "col_next", tb.pos = 1),
                                        S01_PSG_BMI=as.numeric(strsplit(find_field_grep(pattern = "B.M.I.", table = TABLES, type = "col_next", tb.pos = 1), " ")[[1]][1]),
                                        S01_PSG_BMI_unit=strsplit(find_field_grep(pattern = "B.M.I.", table = TABLES, type = "col_next", tb.pos = 1), " ")[[1]][2],
                                        S01_PHI_Referring_Physician=find_field_grep(pattern = "Referring Physician", table = TABLES, type = "col_next", tb.pos = 1),
                                        S01_PHI_Sleep_Specialist=find_field_grep(pattern = "Sleep Specialist", table = TABLES, type = "col_next", tb.pos = 1),
                                        S01_PHI_Location=find_field_grep(pattern = "Location", table = TABLES, type = "col_next", tb.pos = 1),
                                        
                                        # Clinical history and clinical text fields
                                        
                                        S01_PSG_study_type="Split Night Study",
                                        S01_PHI_Clinical_History=paste0("",
                                                                        paste(grep("CLINICAL HISTORY", current_docx_summary$text, value = T),
                                                                              current_docx_summary$text[grep("CLINICAL HISTORY", current_docx_summary$text, ignore.case = T)+1])),
                                        
                                        S01_PSG_Technical_Description=paste0("",
                                                                             paste(grep("technical description", current_docx_summary$text, value = T, ignore.case = T),current_docx_summary$text[grep("technical description", current_docx_summary$text, ignore.case = T)+1])),
                                        
                                        S01_PSG_PHI_Final_Diagnosis=paste0("",
                                                                           paste(grep("final diagnosis", current_docx_summary$text, value = T, ignore.case = T),current_docx_summary$text[grep("final diagnosis", current_docx_summary$text, ignore.case = T)+1])),
                                        
                                        
                                        S01_PSG_PHI_Comments_Recommendations=paste(paste0("",
                                                                                    paste(grep("recommendations", current_docx_summary$text, value = T, ignore.case = T),
                                                                                          current_docx_summary$text[grep("recommendations", current_docx_summary$text, ignore.case = T)+1])), collapse = " "),
                                        S01_PSG_Cardiac_Arrhythmias_Comments=paste0("",
                                                                                    paste(grep("Cardiac Arrhythmias:", current_docx_summary$text, value = T, ignore.case = T),current_docx_summary$text[grep("Cardiac Arrhythmias:", current_docx_summary$text, ignore.case = T)+1])),
                                        
                                        S01_PSG_Recording_Tech_Comments=paste0("",
                                                                               paste(grep("Recording Technician", current_docx_summary$text, value = T, ignore.case = T),current_docx_summary$text[grep("Recording Technician", current_docx_summary$text, ignore.case = T)+1])),
                                        
                                        S01_PSG_Scoring_Tech_Comments=paste0("",
                                                                             paste(grep("Scoring Technician", current_docx_summary$text, value = T, ignore.case = T), current_docx_summary$text[grep("Scoring Technician", current_docx_summary$text, ignore.case = T)+1])),
                                        
                                        
                                        # Start PSG data
                                        S01_PSG_BaselineStartTime_baseline=find_field_grep(pattern = "Start Time", table = TABLES, type = "col_next", tb.pos = 2),
                                        S01_PSG_BaselineEndTime_baseline=find_field_grep(pattern = "End Time ", table = TABLES, type = "col_next", tb.pos = 2),
                                        S01_PSG_TRT_min_baseline=as.numeric(find_field_grep(pattern = "Total Recording Time (minutes)", table = TABLES, type = "col_next", tb.pos = 2)),
                                        S01_PSG_TST_min_baseline=as.numeric(find_field_grep(pattern = "Total Sleep Time (minutes)", table = TABLES, type = "col_next", tb.pos = 2)),
                                        S01_PSG_SleepEfficiency_perc_baseline=as.numeric(find_field_grep(pattern = "Sleep Efficiency (%)", table = TABLES, type = "col_next", tb.pos = 2)),
                                        S01_PSG_SleepOnsetLatency_min_baseline=as.numeric(find_field_grep(pattern = "Sleep Onset Latency (minutes)", table = TABLES, type = "col_next", tb.pos = 2)),
                                        S01_PSG_N_REM_Periods_baseline=as.numeric(find_field_grep(pattern = "Number of REM Periods", table = TABLES, type = "col_next", tb.pos = 2)),
                                        S01_PSG_REM_Latency_baseline=as.numeric(find_field_grep(pattern = "REM Latency", table = TABLES, type = "col_next", tb.pos = 2)),
                                        S01_PSG_Stage_W_min_baseline=as.numeric(find_field_grep(pattern = "Wake", table = TABLES, type = "col_next", tb.pos = 2)),
                                        S01_PSG_Stage_N1_min_baseline=as.numeric(find_field_grep(pattern = "Stage N1|Stage 1", table = TABLES, type = "col_next", tb.pos = 2, fixed = F)),
                                        S01_PSG_Stage_N1_perc_TST_baseline=as.numeric(find_field_grep(pattern = "Stage N1|Stage 1", table = TABLES, type = "col_next", tb.pos = 2, n=2, fixed = F)),
                                        S01_PSG_Stage_N2_min_baseline=as.numeric(find_field_grep(pattern = "Stage N2|Stage 2", table = TABLES, type = "col_next", tb.pos = 2, fixed = F)),
                                        S01_PSG_Stage_N2_perc_TST_baseline=as.numeric(find_field_grep(pattern = "Stage N2|Stage 2", table = TABLES, type = "col_next", tb.pos = 2, n=2, fixed = F)),
                                        S01_PSG_Stage_N3_min_baseline=as.numeric(find_field_grep(pattern = "Stage N3|Stage 3/4", table = TABLES, type = "col_next", tb.pos = 2, fixed = F)),
                                        S01_PSG_Stage_N3_perc_TST_baseline=as.numeric(find_field_grep(pattern = "Stage N3|Stage 3/4", table = TABLES, type = "col_next", tb.pos = 2, n=2, fixed = F)),
                                        S01_PSG_Stage_REM_min_baseline=as.numeric(find_field_grep(pattern = "Total Recording Time (minutes)",
                                                                                         pattern2 = "REM:", table = TABLES, type = "col_next", tb.pos = 2)),
                                        S01_PSG_Stage_REM_perc_TST_baseline=as.numeric(find_field_grep(pattern = "Total Recording Time (minutes)",
                                                                                              pattern2 = "REM:", table = TABLES, type = "col_next", tb.pos = 2, n=2)),
                                        
                                        S01_PSG_Snoring_Level_Supine_baseline =as.character(find_field_grep(pattern = "Snoring Levels",
                                                                                                   pattern2 = "Supine :", table = TABLES, type = "col_next", tb.pos = 1, n=1)),
                                        S01_PSG_Snoring_Level_Lateral_baseline  =as.character(find_field_grep(pattern = "Snoring Levels",
                                                                                                     pattern2 = "Lateral :", table = TABLES, type = "col_next", tb.pos = 1, n=1)),
                                        S01_PSG_Snoring_Level_Prone_baseline  =as.character(find_field_grep(pattern = "Snoring Levels",
                                                                                                   pattern2 = "Prone :", table = TABLES, type = "col_next", tb.pos = 1, n=1)),
                                        
                                        
                                        S01_PSG_CentralApnea_Count_baseline=as.numeric(find_field_grep(pattern = "CENTRAL",
                                                                                              pattern2 = "Number:", table = TABLES, type = "col_next", tb.pos = 1)),
                                        S01_PSG_ObstructiveApnea_Count_baseline=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                  pattern2 = "Number:", table = TABLES, type = "col_next", tb.pos = 1, n=2)),
                                        S01_PSG_MixedApnea_Count_baseline=as.numeric(find_field_grep(pattern = "MIXED",
                                                                                            pattern2 = "Number:", table = TABLES, type = "col_next", tb.pos = 1, n=3)),
                                        S01_PSG_TotalApnea_Count_baseline=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                            pattern2 = "Number:", table = TABLES, type = "col_next", tb.pos = 1, n=4)),
                                        S01_PSG_Hypopnea_Count_baseline=as.numeric(find_field_grep(pattern = "HYPOPNEAS",
                                                                                          pattern2 = "Number:", table = TABLES, type = "col_next", tb.pos = 1, n=5)),
                                        
                                        S01_PSG_CentralApnea_Index_baseline=as.numeric(find_field_grep(pattern = "CENTRAL",
                                                                                              table = TABLES, type = "row_below", tb.pos = 1, n=2)),
                                        S01_PSG_ObstructiveApnea_Index_baseline=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                  table = TABLES, type = "row_below", tb.pos = 1, n=2)),
                                        S01_PSG_MixedApnea_Index_baseline=as.numeric(find_field_grep(pattern = "MIXED",
                                                                                            table = TABLES, type = "row_below", tb.pos = 1, n=2)),
                                        S01_PSG_TotalApnea_Index_baseline=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE", pattern2 = "TOTAL",
                                                                                                     table = TABLES, type = "row_below", tb.pos = 1, n=2)),
                                        S01_PSG_Hypopnea_Index_baseline=as.numeric(find_field_grep(pattern = "HYPOPNEAS",
                                                                                          table = TABLES, type = "row_below", tb.pos = 1, n=2)),
                                        
                                        
                                        S01_PSG_CentralApnea_MeanDuration_sec_baseline=as.numeric(find_field_grep(pattern = "CENTRAL",
                                                                                                         pattern2 = "Mean Duration", table = TABLES, type = "col_next", tb.pos = 1)),
                                        S01_PSG_ObstructiveApnea_MeanDuration_sec_baseline=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                             pattern2 = "Mean Duration", table = TABLES, type = "col_next", tb.pos = 1, n=2)),
                                        S01_PSG_MixedApnea_MeanDuration_sec_baseline=as.numeric(find_field_grep(pattern = "MIXED",
                                                                                                       pattern2 = "Mean Duration", table = TABLES, type = "col_next", tb.pos = 1, n=3)),
                                        S01_PSG_TotalApnea_MeanDuration_sec_baseline=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                       pattern2 = "Mean Duration", table = TABLES, type = "col_next", tb.pos = 1, n=4)),
                                        S01_PSG_Hypopnea_MeanDuration_sec_baseline=as.numeric(find_field_grep(pattern = "HYPOPNEAS",
                                                                                                     pattern2 = "Mean Duration", table = TABLES, type = "col_next", tb.pos = 1, n=5)),
                                        
                                        
                                        
                                        S01_PSG_CentralApnea_LongestDuration_sec_baseline=as.numeric(find_field_grep(pattern = "CENTRAL",
                                                                                                            pattern2 = "Longest Duration", table = TABLES, type = "col_next", tb.pos = 1)),
                                        S01_PSG_ObstructiveApnea_LongestDuration_sec_baseline=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                                pattern2 = "Longest Duration", table = TABLES, type = "col_next", tb.pos = 1, n=2)),
                                        S01_PSG_MixedApnea_LongestDuration_sec_baseline=as.numeric(find_field_grep(pattern = "MIXED",
                                                                                                          pattern2 = "Longest Duration", table = TABLES, type = "col_next", tb.pos = 1, n=3)),
                                        S01_PSG_TotalApnea_LongestDuration_sec_baseline=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                          pattern2 = "Longest Duration", table = TABLES, type = "col_next", tb.pos = 1, n=4)),
                                        S01_PSG_Hypopnea_LongestDuration_sec_baseline=as.numeric(find_field_grep(pattern = "HYPOPNEAS",
                                                                                                        pattern2 = "Longest Duration", table = TABLES, type = "col_next", tb.pos = 1, n=5)),
                                        
                                        
                                        S01_PSG_CentralApnea_REM_Count_baseline=as.numeric(find_field_grep(pattern = "CENTRAL",
                                                                                                  pattern2 = "Occur in REM:", table = TABLES, type = "col_next", tb.pos = 1)),
                                        S01_PSG_ObstructiveApnea_REM_Count_baseline=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                      pattern2 = "Occur in REM:", table = TABLES, type = "col_next", tb.pos = 1, n=2)),
                                        S01_PSG_MixedApnea_REM_Count_baseline=as.numeric(find_field_grep(pattern = "MIXED",
                                                                                                pattern2 = "Occur in REM:", table = TABLES, type = "col_next", tb.pos = 1, n=3)),
                                        S01_PSG_TotalApnea_REM_Count_baseline=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                pattern2 = "Occur in REM:", table = TABLES, type = "col_next", tb.pos = 1, n=4)),
                                        S01_PSG_Hypopnea_REM_Count_baseline=as.numeric(find_field_grep(pattern = "HYPOPNEAS",
                                                                                              pattern2 = "Occur in REM:", table = TABLES, type = "col_next", tb.pos = 1, n=5)),
                                        
                                        
                                        S01_PSG_CentralApnea_NREM_Count_baseline=as.numeric(find_field_grep(pattern = "CENTRAL",
                                                                                                   pattern2 = "Occur in Non-REM:", table = TABLES, type = "col_next", tb.pos = 1)),
                                        S01_PSG_ObstructiveApnea_NREM_Count_baseline=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                       pattern2 = "Occur in Non-REM:", table = TABLES, type = "col_next", tb.pos = 1, n=2)),
                                        S01_PSG_MixedApnea_NREM_Count_baseline=as.numeric(find_field_grep(pattern = "MIXED",
                                                                                                 pattern2 = "Occur in Non-REM:", table = TABLES, type = "col_next", tb.pos = 1, n=3)),
                                        S01_PSG_TotalApnea_NREM_Count_baseline=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                 pattern2 = "Occur in Non-REM:", table = TABLES, type = "col_next", tb.pos = 1, n=4)),
                                        S01_PSG_Hypopnea_NREM_Count_baseline=as.numeric(find_field_grep(pattern = "HYPOPNEAS",
                                                                                               pattern2 = "Occur in Non-REM:", table = TABLES, type = "col_next", tb.pos = 1, n=5)),
                                        
                                        
                                        S01_PSG_CentralApnea_REM_Index_baseline=as.numeric(find_field_grep(pattern = "REM Index:",
                                                                                                  pattern2 = "CENTRAL", table = TABLES, type = "row_below", tb.pos = 1, n=7)),
                                        S01_PSG_ObstructiveApnea_REM_Index_baseline=as.numeric(find_field_grep(pattern = "REM Index:",
                                                                                                      pattern2 = "OBSTRUCTIVE", table = TABLES, type = "row_below", tb.pos = 1, n=7)),
                                        S01_PSG_MixedApnea_REM_Index_baseline=as.numeric(find_field_grep(pattern = "REM Index:",
                                                                                                pattern2 = "MIXED", table = TABLES, type = "row_below", tb.pos = 1, n=7)),
                                        S01_PSG_TotalApnea_REM_Index_baseline=as.numeric(find_field_grep(pattern = "REM Index:",
                                                                                                pattern2 = "TOTAL", table = TABLES, type = "row_below", tb.pos = 1, n=7)),
                                        S01_PSG_Hypopnea_REM_Index_baseline=as.numeric(find_field_grep(pattern = "REM Index:",
                                                                                              pattern2 = "HYPOPNEAS", table = TABLES, type = "row_below", tb.pos = 1, n=7)),
                                        
                                        
                                        S01_PSG_CentralApnea_NREM_Index_baseline=as.numeric(find_field_grep(pattern = "CENTRAL",
                                                                                                   pattern2 = "Non-REM Index:", table = TABLES, type = "col_next", tb.pos = 1)),
                                        S01_PSG_ObstructiveApnea_NREM_Index_baseline=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                       pattern2 = "Non-REM Index:", table = TABLES, type = "col_next", tb.pos = 1, n=2)),
                                        S01_PSG_MixedApnea_NREM_Index_baseline=as.numeric(find_field_grep(pattern = "MIXED",
                                                                                                 pattern2 = "Non-REM Index:", table = TABLES, type = "col_next", tb.pos = 1, n=3)),
                                        S01_PSG_TotalApnea_NREM_Index_baseline=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                 pattern2 = "Non-REM Index:", table = TABLES, type = "col_next", tb.pos = 1, n=4)),
                                        S01_PSG_Hypopnea_NREM_Index_baseline=as.numeric(find_field_grep(pattern = "HYPOPNEAS",
                                                                                               pattern2 = "Non-REM Index:", table = TABLES, type = "col_next", tb.pos = 1, n=5)),
                                        
                                        
                                        S01_PSG_ApneaHypopnea_Index_baseline=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                               pattern2 = "INDEX", table = TABLES, type = "row_below", tb.pos = 1)),
                                        S01_PSG_ApneaHypopnea_Count_baseline=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                               pattern2 = "TOTAL", table = TABLES, type = "row_below", tb.pos = 1)),
                                        S01_PSG_ApneaHypopnea_NREM_Index_baseline=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                    pattern2 = "INDEX", table = TABLES, type = "row_below", tb.pos = 1, n=2)),
                                        S01_PSG_ApneaHypopnea_NREM_Count_baseline=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                    pattern2 = "TOTAL", table = TABLES, type = "row_below", tb.pos = 1, n=2)),
                                        S01_PSG_ApneaHypopnea_REM_Index_baseline=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                   pattern2 = "INDEX", table = TABLES, type = "row_below", tb.pos = 1, n=3)),
                                        S01_PSG_ApneaHypopnea_REM_Count_baseline=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                   pattern2 = "TOTAL", table = TABLES, type = "row_below", tb.pos = 1, n=3)),
                                        S01_PSG_ApneaHypopnea_Supine_Index_baseline=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                      pattern2 = "INDEX", table = TABLES, type = "row_below", tb.pos = 1, n=4)),
                                        S01_PSG_ApneaHypopnea_Supine_Count_baseline=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                      pattern2 = "TOTAL", table = TABLES, type = "row_below", tb.pos = 1, n=4)),
                                        S01_PSG_ApneaHypopnea_Lateral_Index_baseline=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                       pattern2 = "INDEX", table = TABLES, type = "row_below", tb.pos = 1, n=5)),
                                        S01_PSG_ApneaHypopnea_Lateral_Count_baseline=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                       pattern2 = "TOTAL", table = TABLES, type = "row_below", tb.pos = 1, n=5)),
                                        S01_PSG_ApneaHypopnea_Prone_Index_baseline=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                     pattern2 = "INDEX", table = TABLES, type = "row_below", tb.pos = 1, n=6)),
                                        S01_PSG_ApneaHypopnea_Prone_Count_baseline=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                     pattern2 = "TOTAL", table = TABLES, type = "row_below", tb.pos = 1, n=6)),
                                        S01_PSG_ApneaHypopnea_Left_Index_baseline=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                    pattern2 = "INDEX", table = TABLES, type = "row_below", tb.pos = 1, n=7)),
                                        S01_PSG_ApneaHypopnea_Left_Count_baseline=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                    pattern2 = "TOTAL", table = TABLES, type = "row_below", tb.pos = 1, n=7)),
                                        S01_PSG_ApneaHypopnea_Right_Index_baseline=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                     pattern2 = "INDEX", table = TABLES, type = "row_below", tb.pos = 1, n=8)),
                                        S01_PSG_ApneaHypopnea_Right_Count_baseline=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                     pattern2 = "TOTAL", table = TABLES, type = "row_below", tb.pos = 1, n=8)),
                                        
                                        #PLMs and arousal - not well annotated in table (check consitency across files)
                               
                                        S01_PSG_PLM_Index_baseline=as.numeric(TABLES[[PLM_TABLE_id_baseline]][2,2]),
                                        S01_PSG_PLM_NREM_Index_baseline=as.numeric(TABLES[[PLM_TABLE_id_baseline]][3,2]),
                                        S01_PSG_PLM_REM_Index_baseline=as.numeric(TABLES[[PLM_TABLE_id_baseline]][4,2]),
                                        
                                        S01_PSG_PLM_Count_baseline=as.numeric(TABLES[[PLM_TABLE_id_baseline]][2,3]),
                                        S01_PSG_PLM_NREM_Count_baseline=as.numeric(TABLES[[PLM_TABLE_id_baseline]][3,3]),
                                        S01_PSG_PLM_REM_Count_baseline=as.numeric(TABLES[[PLM_TABLE_id_baseline]][4,3]),
                                        
                                        S01_PSG_PLM_wArousal_Index_baseline=as.numeric(TABLES[[PLM_TABLE_id_baseline]][2,4]),
                                        S01_PSG_PLM_wArousal_NREM_Index_baseline=as.numeric(TABLES[[PLM_TABLE_id_baseline]][3,4]),
                                        S01_PSG_PLM_wArousal_REM_Index_baseline=as.numeric(TABLES[[PLM_TABLE_id_baseline]][4,4]),
                                        
                                        S01_PSG_PLM_wArousal_Count_baseline=as.numeric(TABLES[[PLM_TABLE_id_baseline]][2,5]),
                                        S01_PSG_PLM_wArousal_NREM_Count_baseline=as.numeric(TABLES[[PLM_TABLE_id_baseline]][3,5]),
                                        S01_PSG_PLM_wArousal_REM_Count_baseline=as.numeric(TABLES[[PLM_TABLE_id_baseline]][4,5]),
                                        
                                        S01_PSG_PLM_woArousal_Index_baseline=as.numeric(TABLES[[PLM_TABLE_id_baseline]][2,6]),
                                        S01_PSG_PLM_woArousal_NREM_Index_baseline=as.numeric(TABLES[[PLM_TABLE_id_baseline]][3,6]),
                                        S01_PSG_PLM_woArousal_REM_Index_baseline=as.numeric(TABLES[[PLM_TABLE_id_baseline]][4,6]),
                                        
                                        S01_PSG_PLM_woArousal_Count_baseline=as.numeric(TABLES[[PLM_TABLE_id_baseline]][2,7]),
                                        S01_PSG_PLM_woArousal_NREM_Count_baseline=as.numeric(TABLES[[PLM_TABLE_id_baseline]][3,7]),
                                        S01_PSG_PLM_woArousal_REM_Count_baseline=as.numeric(TABLES[[PLM_TABLE_id_baseline]][4,7]),
                                        
                                        
                                        S01_PSG_Arousal_Index_baseline=as.numeric(TABLES[[Arousal_TABLE_id_baseline]][2,2]),
                                        S01_PSG_Arousal_NREM_Index_baseline=as.numeric(TABLES[[Arousal_TABLE_id_baseline]][3,2]),
                                        S01_PSG_Arousal_REM_Index_baseline=as.numeric(TABLES[[Arousal_TABLE_id_baseline]][4,2]),
                                        
                                        S01_PSG_Arousal_Count_baseline=as.numeric(TABLES[[Arousal_TABLE_id_baseline]][2,3]),
                                        S01_PSG_Arousal_NREM_Count_baseline=as.numeric(TABLES[[Arousal_TABLE_id_baseline]][3,3]),
                                        S01_PSG_Arousal_REM_Count_baseline=as.numeric(TABLES[[Arousal_TABLE_id_baseline]][4,3]),
                                        
                                        
                                        S01_PSG_Mean_SaO2_awake_perc_baseline=as.numeric(find_field_grep(pattern = "Mean SaO2",
                                                                                                pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 1, n=1)),
                                        S01_PSG_Mean_SaO2_NREM_perc_baseline=as.numeric(find_field_grep(pattern = "Mean SaO2",
                                                                                               pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 1, n=1)),
                                        S01_PSG_Mean_SaO2_REM_perc_baseline=as.numeric(find_field_grep(pattern = "Mean SaO2",
                                                                                              pattern2 = "Mean SaO2", table = TABLES, type = "col_next", tb.pos = 1, n=3)),
                                        S01_PSG_Mean_SaO2_TST_perc_baseline=as.numeric(find_field_grep(pattern = "Mean SaO2",
                                                                                              pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 1, n=1)),
                                        
                                        S01_PSG_Min_SaO2_awake_perc_baseline=as.numeric(find_field_grep(pattern = "Min. SaO2",
                                                                                               pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 1, n=2)),
                                        S01_PSG_Min_SaO2_NREM_perc_baseline=as.numeric(find_field_grep(pattern = "Min. SaO2",
                                                                                              pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 1, n=2)),
                                        S01_PSG_Min_SaO2_REM_perc_baseline=as.numeric(find_field_grep(pattern = "Min. SaO2",
                                                                                             pattern2 = "Min. SaO2", table = TABLES, type = "col_next", tb.pos = 1, n=3)),
                                        S01_PSG_Min_SaO2_TST_perc_baseline=as.numeric(find_field_grep(pattern = "Min. SaO2",
                                                                                             pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 1, n=2)),
                                        
                                        S01_PSG_Max_SaO2_awake_perc_baseline=as.numeric(find_field_grep(pattern = "Max. SaO2",
                                                                                               pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 1, n=3)),
                                        S01_PSG_Max_SaO2_NREM_perc_baseline=as.numeric(find_field_grep(pattern = "Max. SaO2",
                                                                                              pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 1, n=3)),
                                        S01_PSG_Max_SaO2_REM_perc_baseline=as.numeric(find_field_grep(pattern = "Max. SaO2",
                                                                                             pattern2 = "Max. SaO2", table = TABLES, type = "col_next", tb.pos = 1, n=3)),
                                        S01_PSG_Max_SaO2_TST_perc_baseline=as.numeric(find_field_grep(pattern = "Max. SaO2",
                                                                                             pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 1, n=3)),
                                        
                                        S01_PSG_SaO2_90.100_awake_perc_baseline=as.numeric(find_field_grep(pattern = "90.+100 %",
                                                                                                  pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 1, n=5, fixed=F)),
                                        S01_PSG_SaO2_90.100_NREM_perc_baseline=as.numeric(find_field_grep(pattern = "90.+100 %",
                                                                                                 pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 1, n=5, fixed=F)),
                                        S01_PSG_SaO2_90.100_REM_perc_baseline=as.numeric(find_field_grep(pattern = "90.+100 %",
                                                                                                pattern2 = "90.+100 %", table = TABLES, type = "col_next", tb.pos = 1, n=3, fixed=F)),
                                        S01_PSG_SaO2_90.100_TST_perc_baseline=as.numeric(find_field_grep(pattern = "90.+100 %",
                                                                                                pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 1, n=5, fixed=F)),
                                        
                                        S01_PSG_SaO2_80.89_awake_perc_baseline=as.numeric(find_field_grep(pattern = "80.+89 %",
                                                                                                 pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 1, n=6, fixed=F)),
                                        S01_PSG_SaO2_80.89_NREM_perc_baseline=as.numeric(find_field_grep(pattern = "80.+89 %",
                                                                                                pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 1, n=6, fixed=F)),
                                        S01_PSG_SaO2_80.89_REM_perc_baseline=as.numeric(find_field_grep(pattern = "80.+89 %",
                                                                                               pattern2 = "80.+89 %", table = TABLES, type = "col_next", tb.pos = 1, n=3, fixed=F)),
                                        S01_PSG_SaO2_80.89_TST_perc_baseline=as.numeric(find_field_grep(pattern = "80.+89 %",
                                                                                               pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 1, n=6, fixed=F)),
                                        
                                        S01_PSG_SaO2_70.79_awake_perc_baseline=as.numeric(find_field_grep(pattern = "70.+79 %",
                                                                                                 pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 1, n=7, fixed=F)),
                                        S01_PSG_SaO2_70.79_NREM_perc_baseline=as.numeric(find_field_grep(pattern = "70.+79 %",
                                                                                                pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 1, n=7, fixed=F)),
                                        S01_PSG_SaO2_70.79_REM_perc_baseline=as.numeric(find_field_grep(pattern = "70.+79 %",
                                                                                               pattern2 = "70.+79 %", table = TABLES, type = "col_next", tb.pos = 1, n=3, fixed=F)),
                                        S01_PSG_SaO2_70.79_TST_perc_baseline=as.numeric(find_field_grep(pattern = "70.+79 %",
                                                                                               pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 1, n=7, fixed=F)),
                                        
                                        S01_PSG_SaO2_60.69_awake_perc_baseline=as.numeric(find_field_grep(pattern = "60.+69 %",
                                                                                                 pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 1, n=8, fixed=F)),
                                        S01_PSG_SaO2_60.69_NREM_perc_baseline=as.numeric(find_field_grep(pattern = "60.+69 %",
                                                                                                pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 1, n=8, fixed=F)),
                                        S01_PSG_SaO2_60.69_REM_perc_baseline=as.numeric(find_field_grep(pattern = "60.+69 %",
                                                                                               pattern2 = "60.+69 %", table = TABLES, type = "col_next", tb.pos = 1, n=3, fixed=F)),
                                        S01_PSG_SaO2_60.69_TST_perc_baseline=as.numeric(find_field_grep(pattern = "60.+69 %",
                                                                                               pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 1, n=8, fixed=F)),
                                        
                                        S01_PSG_SaO2_50.59_awake_perc_baseline=as.numeric(find_field_grep(pattern = "50.+59 %",
                                                                                                 pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 1, n=9, fixed=F)),
                                        S01_PSG_SaO2_50.59_NREM_perc_baseline=as.numeric(find_field_grep(pattern = "50.+59 %",
                                                                                                pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 1, n=9, fixed=F)),
                                        S01_PSG_SaO2_50.59_REM_perc_baseline=as.numeric(find_field_grep(pattern = "50.+59 %",
                                                                                               pattern2 = "50.+59 %", table = TABLES, type = "col_next", tb.pos = 1, n=3, fixed=F)),
                                        S01_PSG_SaO2_50.59_TST_perc_baseline=as.numeric(find_field_grep(pattern = "50.+59 %",
                                                                                               pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 1, n=9, fixed=F)),
                                        
                                        S01_PSG_SaO2_Below50_awake_perc_baseline=as.numeric(find_field_grep(pattern = "Below 50 %",
                                                                                                   pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 1, n=10)),
                                        S01_PSG_SaO2_Below50_NREM_perc_baseline=as.numeric(find_field_grep(pattern = "Below 50 %",
                                                                                                  pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 1, n=10)),
                                        S01_PSG_SaO2_Below50REM_perc_baseline=as.numeric(find_field_grep(pattern = "Below 50 %",
                                                                                                pattern2 = "Below 50 %", table = TABLES, type = "col_next", tb.pos = 1, n=3)),
                                        S01_PSG_SaO2_Below50_TST_perc_baseline=as.numeric(find_field_grep(pattern = "Below 50 %",
                                                                                                 pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 1, n=10)),
                                        
                                        
                                        
                                        ########### TREATMENT PART
                                        
                                        # Start PSG data
                                        S01_PSG_TreatmentStartTime_treatment=find_field_grep(pattern = "Start Time", table = TABLES, type = "col_next", tb.pos = 3),
                                        S01_PSG_TreatmentEndTime_treatment=find_field_grep(pattern = "End Time ", table = TABLES, type = "col_next", tb.pos = 3),
                                        S01_PSG_TRT_min_treatment=as.numeric(find_field_grep(pattern = "Total Recording Time (minutes)", table = TABLES, type = "col_next", tb.pos = 3)),
                                        S01_PSG_TST_min_treatment=as.numeric(find_field_grep(pattern = "Total Sleep Time (minutes)", table = TABLES, type = "col_next", tb.pos = 3)),
                                        S01_PSG_SleepEfficiency_perc_treatment=as.numeric(find_field_grep(pattern = "Sleep Efficiency (%)", table = TABLES, type = "col_next", tb.pos = 3)),
                                        S01_PSG_SleepOnsetLatency_min_treatment=as.numeric(find_field_grep(pattern = "Sleep Onset Latency (minutes)", table = TABLES, type = "col_next", tb.pos = 3)),
                                        S01_PSG_N_REM_Periods_treatment=as.numeric(find_field_grep(pattern = "Number of REM Periods", table = TABLES, type = "col_next", tb.pos = 3)),
                                        S01_PSG_REM_Latency_treatment=as.numeric(find_field_grep(pattern = "REM Latency", table = TABLES, type = "col_next", tb.pos = 3)),
                                        S01_PSG_Stage_W_min_treatment=as.numeric(find_field_grep(pattern = "Wake", table = TABLES, type = "col_next", tb.pos = 3)),
                                        S01_PSG_Stage_N1_min_treatment=as.numeric(find_field_grep(pattern = "Stage N1|Stage 1", table = TABLES, type = "col_next", tb.pos = 3, fixed = F)),
                                        S01_PSG_Stage_N1_perc_TST_treatment=as.numeric(find_field_grep(pattern = "Stage N1|Stage 1", table = TABLES, type = "col_next", tb.pos = 3, n=2, fixed = F)),
                                        S01_PSG_Stage_N2_min_treatment=as.numeric(find_field_grep(pattern = "Stage N2|Stage 2", table = TABLES, type = "col_next", tb.pos = 3, fixed = F)),
                                        S01_PSG_Stage_N2_perc_TST_treatment=as.numeric(find_field_grep(pattern = "Stage N2|Stage 2", table = TABLES, type = "col_next", tb.pos = 3, n=2, fixed = F)),
                                        S01_PSG_Stage_N3_min_treatment=as.numeric(find_field_grep(pattern = "Stage N3|Stage 3/4", table = TABLES, type = "col_next", tb.pos = 3, fixed = F)),
                                        S01_PSG_Stage_N3_perc_TST_treatment=as.numeric(find_field_grep(pattern = "Stage N3|Stage 3/4", table = TABLES, type = "col_next", tb.pos = 3, n=2, fixed = F)),
                                        S01_PSG_Stage_REM_min_treatment=as.numeric(find_field_grep(pattern = "Total Recording Time (minutes)",
                                                                                                  pattern2 = "REM:", table = TABLES, type = "col_next", tb.pos = 3)),
                                        S01_PSG_Stage_REM_perc_TST_treatment=as.numeric(find_field_grep(pattern = "Total Recording Time (minutes)",
                                                                                                       pattern2 = "REM:", table = TABLES, type = "col_next", tb.pos = 3, n=2)),
                                        
                                        S01_PSG_Snoring_Level_Supine_treatment=as.character(find_field_grep(pattern = "Snoring Levels",
                                                                                                            pattern2 = "Supine :", table = TABLES, type = "col_next", tb.pos = 2, n=1)),
                                        S01_PSG_Snoring_Level_Lateral_treatment=as.character(find_field_grep(pattern = "Snoring Levels",
                                                                                                              pattern2 = "Lateral :", table = TABLES, type = "col_next", tb.pos = 2, n=1)),
                                        S01_PSG_Snoring_Level_Prone_treatment=as.character(find_field_grep(pattern = "Snoring Levels",
                                                                                                            pattern2 = "Prone :", table = TABLES, type = "col_next", tb.pos = 2, n=1)),
                                        
                                        
                                        S01_PSG_CentralApnea_Count_treatment=as.numeric(find_field_grep(pattern = "CENTRAL",
                                                                                                       pattern2 = "Number:", table = TABLES, type = "col_next", tb.pos = 2)),
                                        S01_PSG_ObstructiveApnea_Count_treatment=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                           pattern2 = "Number:", table = TABLES, type = "col_next", tb.pos = 2, n=2)),
                                        S01_PSG_MixedApnea_Count_treatment=as.numeric(find_field_grep(pattern = "MIXED",
                                                                                                     pattern2 = "Number:", table = TABLES, type = "col_next", tb.pos = 2, n=3)),
                                        S01_PSG_TotalApnea_Count_treatment=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                     pattern2 = "Number:", table = TABLES, type = "col_next", tb.pos = 2, n=4)),
                                        S01_PSG_Hypopnea_Count_treatment=as.numeric(find_field_grep(pattern = "HYPOPNEAS",
                                                                                                   pattern2 = "Number:", table = TABLES, type = "col_next", tb.pos = 2, n=5)),
                                        
                                        S01_PSG_CentralApnea_Index_treatment=as.numeric(find_field_grep(pattern = "CENTRAL",
                                                                                                       table = TABLES, type = "row_below", tb.pos = 2, n=2)),
                                        S01_PSG_ObstructiveApnea_Index_treatment=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                           table = TABLES, type = "row_below", tb.pos = 2, n=2)),
                                        S01_PSG_MixedApnea_Index_treatment=as.numeric(find_field_grep(pattern = "MIXED",
                                                                                                     table = TABLES, type = "row_below", tb.pos = 2, n=2)),
                                        S01_PSG_TotalApnea_Index_treatment=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE", pattern2 = "TOTAL",
                                                                                                     table = TABLES, type = "row_below", tb.pos = 2, n=2)),
                                        S01_PSG_Hypopnea_Index_treatment=as.numeric(find_field_grep(pattern = "HYPOPNEAS",
                                                                                                   table = TABLES, type = "row_below", tb.pos = 2, n=2)),
                                        
                                        
                                        S01_PSG_CentralApnea_MeanDuration_sec_treatment=as.numeric(find_field_grep(pattern = "CENTRAL",
                                                                                                                  pattern2 = "Mean Duration", table = TABLES, type = "col_next", tb.pos = 2)),
                                        S01_PSG_ObstructiveApnea_MeanDuration_sec_treatment=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                                      pattern2 = "Mean Duration", table = TABLES, type = "col_next", tb.pos = 2, n=2)),
                                        S01_PSG_MixedApnea_MeanDuration_sec_treatment=as.numeric(find_field_grep(pattern = "MIXED",
                                                                                                                pattern2 = "Mean Duration", table = TABLES, type = "col_next", tb.pos = 2, n=3)),
                                        S01_PSG_TotalApnea_MeanDuration_sec_treatment=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                                pattern2 = "Mean Duration", table = TABLES, type = "col_next", tb.pos = 2, n=4)),
                                        S01_PSG_Hypopnea_MeanDuration_sec_treatment=as.numeric(find_field_grep(pattern = "HYPOPNEAS",
                                                                                                              pattern2 = "Mean Duration", table = TABLES, type = "col_next", tb.pos = 2, n=5)),
                                        
                                        
                                        
                                        S01_PSG_CentralApnea_LongestDuration_sec_treatment=as.numeric(find_field_grep(pattern = "CENTRAL",
                                                                                                                     pattern2 = "Longest Duration", table = TABLES, type = "col_next", tb.pos = 2)),
                                        S01_PSG_ObstructiveApnea_LongestDuration_sec_treatment=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                                         pattern2 = "Longest Duration", table = TABLES, type = "col_next", tb.pos = 2, n=2)),
                                        S01_PSG_MixedApnea_LongestDuration_sec_treatment=as.numeric(find_field_grep(pattern = "MIXED",
                                                                                                                   pattern2 = "Longest Duration", table = TABLES, type = "col_next", tb.pos = 2, n=3)),
                                        S01_PSG_TotalApnea_LongestDuration_sec_treatment=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                                   pattern2 = "Longest Duration", table = TABLES, type = "col_next", tb.pos = 2, n=4)),
                                        S01_PSG_Hypopnea_LongestDuration_sec_treatment=as.numeric(find_field_grep(pattern = "HYPOPNEAS",
                                                                                                                 pattern2 = "Longest Duration", table = TABLES, type = "col_next", tb.pos = 2, n=5)),
                                        
                                        
                                        S01_PSG_CentralApnea_REM_Count_treatment=as.numeric(find_field_grep(pattern = "CENTRAL",
                                                                                                           pattern2 = "Occur in REM:", table = TABLES, type = "col_next", tb.pos = 2)),
                                        S01_PSG_ObstructiveApnea_REM_Count_treatment=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                               pattern2 = "Occur in REM:", table = TABLES, type = "col_next", tb.pos = 2, n=2)),
                                        S01_PSG_MixedApnea_REM_Count_treatment=as.numeric(find_field_grep(pattern = "MIXED",
                                                                                                         pattern2 = "Occur in REM:", table = TABLES, type = "col_next", tb.pos = 2, n=3)),
                                        S01_PSG_TotalApnea_REM_Count_treatment=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                         pattern2 = "Occur in REM:", table = TABLES, type = "col_next", tb.pos = 2, n=4)),
                                        S01_PSG_Hypopnea_REM_Count_treatment=as.numeric(find_field_grep(pattern = "HYPOPNEAS",
                                                                                                       pattern2 = "Occur in REM:", table = TABLES, type = "col_next", tb.pos = 2, n=5)),
                                        
                                        
                                        S01_PSG_CentralApnea_NREM_Count_treatment=as.numeric(find_field_grep(pattern = "CENTRAL",
                                                                                                            pattern2 = "Occur in Non-REM:", table = TABLES, type = "col_next", tb.pos = 2)),
                                        S01_PSG_ObstructiveApnea_NREM_Count_treatment=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                                pattern2 = "Occur in Non-REM:", table = TABLES, type = "col_next", tb.pos = 2, n=2)),
                                        S01_PSG_MixedApnea_NREM_Count_treatment=as.numeric(find_field_grep(pattern = "MIXED",
                                                                                                          pattern2 = "Occur in Non-REM:", table = TABLES, type = "col_next", tb.pos = 2, n=3)),
                                        S01_PSG_TotalApnea_NREM_Count_treatment=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                          pattern2 = "Occur in Non-REM:", table = TABLES, type = "col_next", tb.pos = 2, n=4)),
                                        S01_PSG_Hypopnea_NREM_Count_treatment=as.numeric(find_field_grep(pattern = "HYPOPNEAS",
                                                                                                        pattern2 = "Occur in Non-REM:", table = TABLES, type = "col_next", tb.pos = 2, n=5)),
                                        
                                        
                                        S01_PSG_CentralApnea_REM_Index_treatment=as.numeric(find_field_grep(pattern = "REM Index:",
                                                                                                           pattern2 = "CENTRAL", table = TABLES, type = "row_below", tb.pos = 2, n=7)),
                                        S01_PSG_ObstructiveApnea_REM_Index_treatment=as.numeric(find_field_grep(pattern = "REM Index:",
                                                                                                               pattern2 = "OBSTRUCTIVE", table = TABLES, type = "row_below", tb.pos = 2, n=7)),
                                        S01_PSG_MixedApnea_REM_Index_treatment=as.numeric(find_field_grep(pattern = "REM Index:",
                                                                                                         pattern2 = "MIXED", table = TABLES, type = "row_below", tb.pos = 2, n=7)),
                                        S01_PSG_TotalApnea_REM_Index_treatment=as.numeric(find_field_grep(pattern = "REM Index:",
                                                                                                         pattern2 = "TOTAL", table = TABLES, type = "row_below", tb.pos = 2, n=7)),
                                        S01_PSG_Hypopnea_REM_Index_treatment=as.numeric(find_field_grep(pattern = "REM Index:",
                                                                                                       pattern2 = "HYPOPNEAS", table = TABLES, type = "row_below", tb.pos = 2, n=7)),
                                        
                                        
                                        S01_PSG_CentralApnea_NREM_Index_treatment=as.numeric(find_field_grep(pattern = "CENTRAL",
                                                                                                            pattern2 = "Non-REM Index:", table = TABLES, type = "col_next", tb.pos = 2)),
                                        S01_PSG_ObstructiveApnea_NREM_Index_treatment=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                                pattern2 = "Non-REM Index:", table = TABLES, type = "col_next", tb.pos = 2, n=2)),
                                        S01_PSG_MixedApnea_NREM_Index_treatment=as.numeric(find_field_grep(pattern = "MIXED",
                                                                                                          pattern2 = "Non-REM Index:", table = TABLES, type = "col_next", tb.pos = 2, n=3)),
                                        S01_PSG_TotalApnea_NREM_Index_treatment=as.numeric(find_field_grep(pattern = "OBSTRUCTIVE",
                                                                                                          pattern2 = "Non-REM Index:", table = TABLES, type = "col_next", tb.pos = 2, n=4)),
                                        S01_PSG_Hypopnea_NREM_Index_treatment=as.numeric(find_field_grep(pattern = "HYPOPNEAS",
                                                                                                        pattern2 = "Non-REM Index:", table = TABLES, type = "col_next", tb.pos = 2, n=5)),
                                        
                                        
                                        S01_PSG_ApneaHypopnea_Index_treatment=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                        pattern2 = "INDEX", table = TABLES, type = "row_below", tb.pos = 2)),
                                        S01_PSG_ApneaHypopnea_Count_treatment=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                        pattern2 = "TOTAL", table = TABLES, type = "row_below", tb.pos = 2)),
                                        S01_PSG_ApneaHypopnea_NREM_Index_treatment=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                             pattern2 = "INDEX", table = TABLES, type = "row_below", tb.pos = 2, n=2)),
                                        S01_PSG_ApneaHypopnea_NREM_Count_treatment=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                             pattern2 = "TOTAL", table = TABLES, type = "row_below", tb.pos = 2, n=2)),
                                        S01_PSG_ApneaHypopnea_REM_Index_treatment=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                            pattern2 = "INDEX", table = TABLES, type = "row_below", tb.pos = 2, n=3)),
                                        S01_PSG_ApneaHypopnea_REM_Count_treatment=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                            pattern2 = "TOTAL", table = TABLES, type = "row_below", tb.pos = 2, n=3)),
                                        S01_PSG_ApneaHypopnea_Supine_Index_treatment=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                               pattern2 = "INDEX", table = TABLES, type = "row_below", tb.pos = 2, n=4)),
                                        S01_PSG_ApneaHypopnea_Supine_Count_treatment=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                               pattern2 = "TOTAL", table = TABLES, type = "row_below", tb.pos = 2, n=4)),
                                        S01_PSG_ApneaHypopnea_Lateral_Index_treatment=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                                pattern2 = "INDEX", table = TABLES, type = "row_below", tb.pos = 2, n=5)),
                                        S01_PSG_ApneaHypopnea_Lateral_Count_treatment=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                                pattern2 = "TOTAL", table = TABLES, type = "row_below", tb.pos = 2, n=5)),
                                        S01_PSG_ApneaHypopnea_Prone_Index_treatment=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                              pattern2 = "INDEX", table = TABLES, type = "row_below", tb.pos = 2, n=6)),
                                        S01_PSG_ApneaHypopnea_Prone_Count_treatment=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                              pattern2 = "TOTAL", table = TABLES, type = "row_below", tb.pos = 2, n=6)),
                                        S01_PSG_ApneaHypopnea_Left_Index_treatment=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                             pattern2 = "INDEX", table = TABLES, type = "row_below", tb.pos = 2, n=7)),
                                        S01_PSG_ApneaHypopnea_Left_Count_treatment=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                             pattern2 = "TOTAL", table = TABLES, type = "row_below", tb.pos = 2, n=7)),
                                        S01_PSG_ApneaHypopnea_Right_Index_treatment=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                              pattern2 = "INDEX", table = TABLES, type = "row_below", tb.pos = 2, n=8)),
                                        S01_PSG_ApneaHypopnea_Right_Count_treatment=as.numeric(find_field_grep(pattern = "Apneas ",
                                                                                                              pattern2 = "TOTAL", table = TABLES, type = "row_below", tb.pos = 2, n=8)),
                                        
                                        #PLMs and arousal - not well annotated in table (check consitency across files)
                                      
                                        S01_PSG_PLM_Index_treatment=as.numeric(TABLES[[PLM_TABLE_id_treatment]][2,2]),
                                        S01_PSG_PLM_NREM_Index_treatment=as.numeric(TABLES[[PLM_TABLE_id_treatment]][3,2]),
                                        S01_PSG_PLM_REM_Index_treatment=as.numeric(TABLES[[PLM_TABLE_id_treatment]][4,2]),
                                        
                                        S01_PSG_PLM_Count_treatment=as.numeric(TABLES[[PLM_TABLE_id_treatment]][2,3]),
                                        S01_PSG_PLM_NREM_Count_treatment=as.numeric(TABLES[[PLM_TABLE_id_treatment]][3,3]),
                                        S01_PSG_PLM_REM_Count_treatment=as.numeric(TABLES[[PLM_TABLE_id_treatment]][4,3]),
                                        
                                        S01_PSG_PLM_wArousal_Index_treatment=as.numeric(TABLES[[PLM_TABLE_id_treatment]][2,4]),
                                        S01_PSG_PLM_wArousal_NREM_Index_treatment=as.numeric(TABLES[[PLM_TABLE_id_treatment]][3,4]),
                                        S01_PSG_PLM_wArousal_REM_Index_treatment=as.numeric(TABLES[[PLM_TABLE_id_treatment]][4,4]),
                                        
                                        S01_PSG_PLM_wArousal_Count_treatment=as.numeric(TABLES[[PLM_TABLE_id_treatment]][2,5]),
                                        S01_PSG_PLM_wArousal_NREM_Count_treatment=as.numeric(TABLES[[PLM_TABLE_id_treatment]][3,5]),
                                        S01_PSG_PLM_wArousal_REM_Count_treatment=as.numeric(TABLES[[PLM_TABLE_id_treatment]][4,5]),
                                        
                                        S01_PSG_PLM_woArousal_Index_treatment=as.numeric(TABLES[[PLM_TABLE_id_treatment]][2,6]),
                                        S01_PSG_PLM_woArousal_NREM_Index_treatment=as.numeric(TABLES[[PLM_TABLE_id_treatment]][3,6]),
                                        S01_PSG_PLM_woArousal_REM_Index_treatment=as.numeric(TABLES[[PLM_TABLE_id_treatment]][4,6]),
                                        
                                        S01_PSG_PLM_woArousal_Count_treatment=as.numeric(TABLES[[PLM_TABLE_id_treatment]][2,7]),
                                        S01_PSG_PLM_woArousal_NREM_Count_treatment=as.numeric(TABLES[[PLM_TABLE_id_treatment]][3,7]),
                                        S01_PSG_PLM_woArousal_REM_Count_treatment=as.numeric(TABLES[[PLM_TABLE_id_treatment]][4,7]),
                                        
                                        
                                        S01_PSG_Arousal_Index_treatment=as.numeric(TABLES[[Arousal_TABLE_id_treatment]][2,2]),
                                        S01_PSG_Arousal_NREM_Index_treatment=as.numeric(TABLES[[Arousal_TABLE_id_treatment]][3,2]),
                                        S01_PSG_Arousal_REM_Index_treatment=as.numeric(TABLES[[Arousal_TABLE_id_treatment]][4,2]),
                                        
                                        S01_PSG_Arousal_Count_treatment=as.numeric(TABLES[[Arousal_TABLE_id_treatment]][2,3]),
                                        S01_PSG_Arousal_NREM_Count_treatment=as.numeric(TABLES[[Arousal_TABLE_id_treatment]][3,3]),
                                        S01_PSG_Arousal_REM_Count_treatment=as.numeric(TABLES[[Arousal_TABLE_id_treatment]][4,3]),
                                        
                                        
                                        S01_PSG_Mean_SaO2_awake_perc_treatment=as.numeric(find_field_grep(pattern = "Mean SaO2",
                                                                                                         pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 2, n=1)),
                                        S01_PSG_Mean_SaO2_NREM_perc_treatment=as.numeric(find_field_grep(pattern = "Mean SaO2",
                                                                                                        pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 2, n=1)),
                                        S01_PSG_Mean_SaO2_REM_perc_treatment=as.numeric(find_field_grep(pattern = "Mean SaO2",
                                                                                                       pattern2 = "Mean SaO2", table = TABLES, type = "col_next", tb.pos = 2, n=3)),
                                        S01_PSG_Mean_SaO2_TST_perc_treatment=as.numeric(find_field_grep(pattern = "Mean SaO2",
                                                                                                       pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 2, n=1)),
                                        
                                        S01_PSG_Min_SaO2_awake_perc_treatment=as.numeric(find_field_grep(pattern = "Min. SaO2",
                                                                                                        pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 2, n=2)),
                                        S01_PSG_Min_SaO2_NREM_perc_treatment=as.numeric(find_field_grep(pattern = "Min. SaO2",
                                                                                                       pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 2, n=2)),
                                        S01_PSG_Min_SaO2_REM_perc_treatment=as.numeric(find_field_grep(pattern = "Min. SaO2",
                                                                                                      pattern2 = "Min. SaO2", table = TABLES, type = "col_next", tb.pos = 2, n=3)),
                                        S01_PSG_Min_SaO2_TST_perc_treatment=as.numeric(find_field_grep(pattern = "Min. SaO2",
                                                                                                      pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 2, n=2)),
                                        
                                        S01_PSG_Max_SaO2_awake_perc_treatment=as.numeric(find_field_grep(pattern = "Max. SaO2",
                                                                                                        pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 2, n=3)),
                                        S01_PSG_Max_SaO2_NREM_perc_treatment=as.numeric(find_field_grep(pattern = "Max. SaO2",
                                                                                                       pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 2, n=3)),
                                        S01_PSG_Max_SaO2_REM_perc_treatment=as.numeric(find_field_grep(pattern = "Max. SaO2",
                                                                                                      pattern2 = "Max. SaO2", table = TABLES, type = "col_next", tb.pos = 2, n=3)),
                                        S01_PSG_Max_SaO2_TST_perc_treatment=as.numeric(find_field_grep(pattern = "Max. SaO2",
                                                                                                      pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 2, n=3)),
                                        
                                        S01_PSG_SaO2_90.100_awake_perc_treatment=as.numeric(find_field_grep(pattern = "90.+100 %",
                                                                                                           pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 2, n=5, fixed=F)),
                                        S01_PSG_SaO2_90.100_NREM_perc_treatment=as.numeric(find_field_grep(pattern = "90.+100 %",
                                                                                                          pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 2, n=5, fixed=F)),
                                        S01_PSG_SaO2_90.100_REM_perc_treatment=as.numeric(find_field_grep(pattern = "90.+100 %",
                                                                                                         pattern2 = "90.+100 %", table = TABLES, type = "col_next", tb.pos = 2, n=3, fixed=F)),
                                        S01_PSG_SaO2_90.100_TST_perc_treatment=as.numeric(find_field_grep(pattern = "90.+100 %",
                                                                                                         pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 2, n=5, fixed=F)),
                                        
                                        S01_PSG_SaO2_80.89_awake_perc_treatment=as.numeric(find_field_grep(pattern = "80.+89 %",
                                                                                                          pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 2, n=6, fixed=F)),
                                        S01_PSG_SaO2_80.89_NREM_perc_treatment=as.numeric(find_field_grep(pattern = "80.+89 %",
                                                                                                         pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 2, n=6, fixed=F)),
                                        S01_PSG_SaO2_80.89_REM_perc_treatment=as.numeric(find_field_grep(pattern = "80.+89 %",
                                                                                                        pattern2 = "80.+89 %", table = TABLES, type = "col_next", tb.pos = 2, n=3, fixed=F)),
                                        S01_PSG_SaO2_80.89_TST_perc_treatment=as.numeric(find_field_grep(pattern = "80.+89 %",
                                                                                                        pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 2, n=6, fixed=F)),
                                        
                                        S01_PSG_SaO2_70.79_awake_perc_treatment=as.numeric(find_field_grep(pattern = "70.+79 %",
                                                                                                          pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 2, n=7, fixed=F)),
                                        S01_PSG_SaO2_70.79_NREM_perc_treatment=as.numeric(find_field_grep(pattern = "70.+79 %",
                                                                                                         pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 2, n=7, fixed=F)),
                                        S01_PSG_SaO2_70.79_REM_perc_treatment=as.numeric(find_field_grep(pattern = "70.+79 %",
                                                                                                        pattern2 = "70.+79 %", table = TABLES, type = "col_next", tb.pos = 2, n=3, fixed=F)),
                                        S01_PSG_SaO2_70.79_TST_perc_treatment=as.numeric(find_field_grep(pattern = "70.+79 %",
                                                                                                        pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 2, n=7, fixed=F)),
                                        
                                        S01_PSG_SaO2_60.69_awake_perc_treatment=as.numeric(find_field_grep(pattern = "60.+69 %",
                                                                                                          pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 2, n=8, fixed=F)),
                                        S01_PSG_SaO2_60.69_NREM_perc_treatment=as.numeric(find_field_grep(pattern = "60.+69 %",
                                                                                                         pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 2, n=8, fixed=F)),
                                        S01_PSG_SaO2_60.69_REM_perc_treatment=as.numeric(find_field_grep(pattern = "60.+69 %",
                                                                                                        pattern2 = "60.+69 %", table = TABLES, type = "col_next", tb.pos = 2, n=3, fixed=F)),
                                        S01_PSG_SaO2_60.69_TST_perc_treatment=as.numeric(find_field_grep(pattern = "60.+69 %",
                                                                                                        pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 2, n=8, fixed=F)),
                                        
                                        S01_PSG_SaO2_50.59_awake_perc_treatment=as.numeric(find_field_grep(pattern = "50.+59 %",
                                                                                                          pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 2, n=9, fixed=F)),
                                        S01_PSG_SaO2_50.59_NREM_perc_treatment=as.numeric(find_field_grep(pattern = "50.+59 %",
                                                                                                         pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 2, n=9, fixed=F)),
                                        S01_PSG_SaO2_50.59_REM_perc_treatment=as.numeric(find_field_grep(pattern = "50.+59 %",
                                                                                                        pattern2 = "50.+59 %", table = TABLES, type = "col_next", tb.pos = 2, n=3, fixed=F)),
                                        S01_PSG_SaO2_50.59_TST_perc_treatment=as.numeric(find_field_grep(pattern = "50.+59 %",
                                                                                                        pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 2, n=9, fixed=F)),
                                        
                                        S01_PSG_SaO2_Below50_awake_perc_treatment=as.numeric(find_field_grep(pattern = "Below 50 %",
                                                                                                            pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 2, n=10)),
                                        S01_PSG_SaO2_Below50_NREM_perc_treatment=as.numeric(find_field_grep(pattern = "Below 50 %",
                                                                                                           pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 2, n=10)),
                                        S01_PSG_SaO2_Below50REM_perc_treatment=as.numeric(find_field_grep(pattern = "Below 50 %",
                                                                                                         pattern2 = "Below 50 %", table = TABLES, type = "col_next", tb.pos = 2, n=3)),
                                        S01_PSG_SaO2_Below50_TST_perc_treatment=as.numeric(find_field_grep(pattern = "Below 50 %",
                                                                                                          pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 2, n=10)),
                                stringsAsFactors = F)
                                
                                ### Titration Data
                                
                                
                                Titration_tb_id <- grep("CPAP Pressures", TABLES)
                                Titration_table <- TABLES[[Titration_tb_id]]
                                
                                colnames(Titration_table) <- Titration_table[1,]
                                Titration_table <- slice(Titration_table, -c(1:2, nrow(Titration_table)))
                                
                                
                                Titration_table_long <- Titration_table %>% tidyr::gather(Parameter, Value, `TRT (min)`:`% time Sao2 < 70`)
 
                                Titration_table_long$Varname <- gsub(" ", "", str_replace_all(paste0(Titration_table_long$Parameter, "_Pressure_", Titration_table_long$`Pressure(cm H2O)`), "[[:punct:]]", "_"))
                                
                                Titration_table_long_clean <- data.frame(t(Titration_table_long[,c("Varname", "Value")]), stringsAsFactors = F)
                                colnames(Titration_table_long_clean) <- paste0("S01_PSG_", Titration_table_long_clean[1,])
                                Titration_table_long_clean <- slice(Titration_table_long_clean,2)
                                
                                
                                ######## Combine ALL data
                                
                                
                                all_fields <- c("S01_PSG_PHI_Filename", "PennSleepID", colnames(split_1_df), colnames(Titration_table_long_clean))
                                
                                #Initiate dataframe with results
                                
                                
                                
                                #If first, create final_df
                                if (first) {
                                        final_df <- data.frame(matrix(NA,1,length(all_fields)))
                                        colnames(final_df) <- all_fields
                                        
                                        final_df[1,] <- c(docx,
                                                          NA,
                                                          split_1_df[1,],
                                                          Titration_table_long_clean[1,])
                                        first <- F
                                } else {
                                        
                                        next_final_df <- data.frame(matrix(NA,1,length(all_fields)))
                                        colnames(next_final_df) <- all_fields
                                        
                                        next_final_df[1,] <- c(docx,
                                                               NA,
                                                               split_1_df[1,],
                                                               Titration_table_long_clean[1,])
                                        
                                        # Merge first with and next
                                        final_df <- dplyr::bind_rows(final_df, next_final_df)
                                        
                                }
                                
                                # Delete corresponding folder
                                # new_tmp_files <- list.files(tempdir(), full.names = T, recursive = T)
                                # tmp_files_toRemove <- new_tmp_files[!(new_tmp_files %in% current_tmp_files)]
                                # tmp_folder_toRemove <- unique(paste("", sapply(strsplit(tmp_files_toRemove, "/"), "[[", 2), sapply(strsplit(tmp_files_toRemove, "/"), "[[", 3), sapply(strsplit(tmp_files_toRemove, "/"), "[[", 4),"", sep = "/"))
                                # unlink(tmp_folder_toRemove, recursive = T)
                                # 
                                # Delete file from working directory
                                file.remove(docx)
                                
                                
                                
                        }, error=function(cond) {
                                message("Something went wrong.")
                                message("Here's the original error message:")
                                message(cond)
                                print(docx)
                                
                                
                        }, finally = print("Done")
                        
                )
                
                i=i+1
                
        }
        
        #Clean data
        #Replace blanks and N/A with NA
        final_df[final_df==""] <- NA
        final_df[final_df=="N/A"] <- NA
        
        #Calculate Age at Study
        final_df$S01_PSG_Age_at_Study <- interval(mdy(final_df$S01_PHI_Date_of_birth), mdy(final_df$S01_PSG_Study_date)) %/% years(1)
        
        #Create identifiable data frame
        identifiable_df <- final_df
        identifiable_df$ProcessedDate <- Sys.Date()
        
        ProcessedTimeID=gsub("-","",gsub(":","",gsub(" ", "", Sys.time())))
        
        identifiable_df$ProcessedTimeID <- ProcessedTimeID
        
        # Per sample QC (not value filter)
        
        QC_df <- data.frame(S01_PSG_PHI_Filename=identifiable_df$S01_PSG_PHI_Filename,
                            # Are there DOCPROPERTY?
                            has_DOCPROPERTY=grepl("DOCPROPERTY", identifiable_df$S01_PHI_Patient_Name),
                            # Missing both MRN fields
                            has_missingMRN=is.na(identifiable_df$S01_PHI_MRN) & is.na(identifiable_df$S01_PHI_Sleep_Center_ID),
                            # Gender not (Male, Female, Unknown)
                            has_missingGender=!(tolower(identifiable_df$S01_PSG_Sex) %in% c("male", "female", "unknown")),
                            # Missing age
                            has_missingAge=is.na(identifiable_df$S01_PSG_Age) & is.na(identifiable_df$S01_PSG_Age_at_Study),
                            # Missing BMI
                            has_missingBMI=is.na(identifiable_df$S01_PSG_BMI)
                            )
        
        identifiable_df$PerSample_QC <- NA
        identifiable_df$PerSample_QC[rowSums(QC_df[,2:6])>0] <- "FAIL"
        identifiable_df$PerSample_QC[rowSums(QC_df[,2:6])==0] <- "PASS"
        
        QC_df$PerSample_QC <- identifiable_df$PerSample_QC
        write.csv(QC_df, paste0("QC_dataframe_", ProcessedTimeID, ".csv"))
        
        
        
        
        #Save identifiable Rdata file encrypted
        #key <- key_sodium(sodium::keygen())
        saveRDS(identifiable_df, paste0("PennSleepDatabase_Split01_", ProcessedTimeID, ".Rdata"))
        #cyphr::encrypt_file("PennSleepDatabase.Rdata", key, "PennSleepDatabase.encrypted")
        
        #Save de-identified version in CSV files - implement a way to check in database if sample was processed and not use the same de-identified IDs
        #identifiable_df$PennSleepID <- paste0("PENNSLEEP00000",seq(1:nrow(identifiable_df)))
        deidentified_df <- select(identifiable_df, -starts_with("PSG_PHI_"))
        
        write.csv(deidentified_df, paste0("PennSleepDatabase_Split01_Deidentified",ProcessedTimeID,".csv"), row.names = F)
        
        return(identifiable_df)
        
        
}

