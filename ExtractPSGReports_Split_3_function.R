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


direct_grep <- function(pattern, pattern2=pattern, table, tb.pos=1, n=1, fixed=T) {
        tryCatch({
                
                tb.pos.id <- tb.pos
                
                #pattern <- "Custom_Diagnostic_TotalArousalsIndex_TST"
                #table <- TABLES
                 
                #Find table
                tb_idx <- grep(pattern, table, fixed = fixed)[tb.pos.id]
                
                tb <- table[[tb_idx]]
                
                tb_col_idx <- grep(pattern2, tb, fixed = fixed)
                tb_col <-  tb[,tb_col_idx]
                tb_field <- grep(pattern2, pull(tb_col), value = T, fixed = fixed)
                
               
                result <- tb_field
                
                
                
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
        "NPBMCustomPatientInfoCardiacArrhythmias",
        "NPBMCustomPatientInfoBaselineSupine", #Snoring supine
        "NPBMCustomPatientInfoBaselineLateral", #Snoring lateral
        "NPBMCustomPatientInfoBaselineProne" #Snoring prone
)

extract_split_3 <- function(paths) {
        
        #paths <- s03_paths[1:10]
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
                                dir.create("temp_reports_s03/", showWarnings = FALSE)
                                file.copy(PSGfilename, to = "temp_reports_s03/")
                                
                                docx <- paste0("temp_reports_s03/", sapply(strsplit(PSGfilename, "/"), "[[", 6))
                                
                                
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
                                
                                # PSG variables - Entire night
                                colnames(word_main_values_df)[colnames(word_main_values_df)=="VAR_LightsOff"] <- "StartTime_entire"
                                colnames(word_main_values_df)[colnames(word_main_values_df)=="VAR_LightsOn"] <- "EndTime_entire"
                                colnames(word_main_values_df)[colnames(word_main_values_df)=="VAR_TotalRecordingTime_Minutes"] <- "TotalRecordingTime_entire"
                                colnames(word_main_values_df)[colnames(word_main_values_df)=="VAR_TotalSleepTime_Minutes"] <- "TotalSleepTime_entire"
                                
                                # PSG Variables - baseline
                                word_main_values_df$StartTime_baseline <- find_field_grep("Start Time:", TABLES, type = "col_next", tb.pos=2, fixed = T)
                                word_main_values_df$EndTime_baseline <- find_field_grep("End Time:", TABLES, type = "col_next", tb.pos=2, fixed = T)
                                word_main_values_df$TotalRecordingTime_baseline <- find_field_grep("Recording Time (min):", TABLES, type = "col_next", tb.pos=2, fixed = T)
                                word_main_values_df$TotalSleepTime_baseline <- find_field_grep("Sleep Time TST (min):", TABLES, type = "col_next", tb.pos=2, fixed = T)
                                word_main_values_df$SleepEfficiency_baseline <- gsub("%","",find_field_grep("Sleep Efficiency (%):", TABLES, type = "col_next", tb.pos=1, fixed = T))
                                word_main_values_df$SleepOnsetLatency_baseline <- find_field_grep("Sleep Onset Latency (min):", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$NumberREMPeriods_baseline <- find_field_grep("Number of REM Periods:", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$REMLatency_baseline <- find_field_grep("REM Latency:", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$WASO_baseline <- find_field_grep("WASO:", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$StageN1_baseline <- find_field_grep("Stage N1:", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$StageN2_baseline <- find_field_grep("Stage N2:", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$StageN3_baseline <- find_field_grep("Stage N3", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$StageREM_baseline <- find_field_grep("REM:", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$StageN1_pct_baseline <- trim(gsub("%)", "", find_field_grep("Stage N1:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=2)))
                                word_main_values_df$StageN2_pct_baseline <- trim(gsub("%)", "", find_field_grep("Stage N2:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=2)))
                                word_main_values_df$StageN3_pct_baseline <- trim(gsub("%)", "", find_field_grep("Stage N3", TABLES, type = "col_next", tb.pos=1, fixed = T, n=2)))
                                word_main_values_df$StageREM_pct_baseline <- trim(gsub("%)", "", find_field_grep("REM:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=2)))
                                
                                ### Extra non-available before values
                                word_main_values_df$Arousal_TST_Idx_baseline <- direct_grep(pattern = "Custom_Diagnostic_TotalArousalsIndex_TST", table = TABLES)
                                word_main_values_df$Arousal_NREM_Idx_baseline <- direct_grep(pattern = "VAR_Custom_Diagnostic_TotalArousalsIndex_NREM", table = TABLES)
                                word_main_values_df$Arousal_REM_Idx_baseline <- direct_grep(pattern = "VAR_Custom_Diagnostic_TotalArousalsIndex_REM", table = TABLES)
                                word_main_values_df$Arousal_TST_count_baseline <- direct_grep(pattern = "VAR_Custom_Diagnostic_TotalArousalCount_TST", table = TABLES)
                                word_main_values_df$Arousal_NREM_count_baseline <- direct_grep(pattern = "VAR_Custom_Diagnostic_TotalArousalCount_NREM", table = TABLES)
                                word_main_values_df$Arousal_REM_count_baseline <- direct_grep(pattern = "VAR_Custom_Diagnostic_TotalArousalCount_REM", table = TABLES)
                                
                                word_main_values_df$PLM_all_TST_Idx_baseline <- direct_grep(pattern = "VAR_Diagnostic_Custom_LimbMovementIndex_TST", table = TABLES)
                                word_main_values_df$PLM_all_NREM_Idx_baseline <- direct_grep(pattern = "VAR_Diagnostic_Custom_LimbMovementIndex_NREM", table = TABLES)
                                word_main_values_df$PLM_all_REM_Idx_baseline <- direct_grep(pattern = "VAR_Diagnostic_Custom_LimbMovementIndex_REM", table = TABLES)
                                word_main_values_df$PLM_all_TST_count_baseline <- direct_grep(pattern = "VAR_Diagnostic_Custom_LimbMovementCount_TST", table = TABLES)
                                word_main_values_df$PLM_all_NREM_count_baseline <- direct_grep(pattern = "VAR_Diagnostic_Custom_LimbMovementCount_NREM", table = TABLES)
                                word_main_values_df$PLM_all_REM_count_baseline <- direct_grep(pattern = "VAR_Diagnostic_Custom_LimbMovementCount_REM", table = TABLES)
                                
                                word_main_values_df$PLM_wArousal_TST_Idx_baseline <- direct_grep(pattern = "VAR_Diagnostic_Custom_LimbMovementArousalsIndex_TST", table = TABLES)
                                word_main_values_df$PLM_wArousal_NREM_Idx_baseline <- direct_grep(pattern = "VAR_Diagnostic_Custom_LimbMovementArousalsIndex_NREM", table = TABLES)
                                word_main_values_df$PLM_wArousal_REM_Idx_baseline <- direct_grep(pattern = "VAR_Diagnostic_Custom_LimbMovementArousalsIndex_REM", table = TABLES)
                                word_main_values_df$PLM_wArousal_TST_count_baseline <- direct_grep(pattern = "VAR_Diagnostic_Custom_LimbMovementArousalCount_TST", table = TABLES)
                                word_main_values_df$PLM_wArousal_NREM_count_baseline <- direct_grep(pattern = "VAR_Diagnostic_Custom_LimbMovementArousalCount_NREM", table = TABLES)
                                word_main_values_df$PLM_wArousal_REM_count_baseline <- direct_grep(pattern = "VAR_Diagnostic_Custom_LimbMovementArousalCount_REM", table = TABLES)
                                
                                
                                
                                
                                #Fix columns names
                                colnames(word_main_values_df)[c(1:3,5:10,17,18)] <- paste0("S03_PSG_PHI_",colnames(word_main_values_df)[c(1:3,5:10,17,18)])
                                colnames(word_main_values_df)[-c(1:3,5:10,17,18)] <- paste0("S03_PSG_",colnames(word_main_values_df)[-c(1:3,5:10,17,18)])
                                
                                
                                #Fix as numeric
                                word_main_values_df[,c(11,13,15,21,22,29:61)] <- sapply(word_main_values_df[,c(11,13,15,21,22,29:61)], as.numeric)
                                
                                
                                
                                # Get other relevant values from the word document
                                word_other_values <- c(S03_PSG_study_type="Split Night Study",
                                                       S03_PSG_PHI_clinical_history=gsub(" The patient(.*?)with a history of ", "", grep("CLINICAL HISTORY:", textreadr::read_docx(PSGfilename), value = T)),
                                                       S03_PSG_PHI_final_diagnosis=grep("FINAL DIAGNOSIS:", textreadr::read_docx(PSGfilename), value = T),
                                                       S03_PSG_PHI_comm_rec=grep("RECOMMENDATIONS:", textreadr::read_docx(PSGfilename), value = T)
                                )
                                
                                word_other_values_df <-  data.frame(t(word_other_values), stringsAsFactors = F)
                                
                                
                                ### Get fields from Excel embedded tables that appear in the report
                                xl_tb <- PSGfilename
                                
                                #Copy file ################## Fix ths to be universal
                                dir.create("zips_s03/", showWarnings = F)
                                file.copy(from = PSGfilename, to = "zips_s03/")
                                file.rename(paste0("zips_s03/", strsplit(PSGfilename, "/")[[1]][6]), "zips_s03/current.zip")
                                unzip("zips_s03/current.zip", exdir = paste0(getwd(),"/zips_s03"))
                                xls_tables_paths <- grep("Worksheet", dir("zips_s03/word/embeddings/", full.names = T), value = T)
                                
                                
                                #Fix values that are formated as dates in Excel
                                process_excel_date <- function(x) {
                                        if (x=="1900-01-01 00:00:00") {
                                                return("1")
                                        } else {
                                                return(as.character(time_length(ymd_hms(x)-ymd_hms("1899-12-30 00:00:00"), "days")))
                                                
                                        }
                                }
                                
                                
                                # Iterate to load all possible tables
                                xtblist <- NULL
                                
                                for (xt in xls_tables_paths) {
                                        xtblist[[xt]]  <- read.xlsx(xt, sheetIndex = 1)
                                }
                                
                                
                                ox_sat_tb.ids <- which(grepl("Oxygen.Saturation", xtblist))
                                
                                # Get second spreadsheet from these
                                
                                xtblist_data <- NULL
                                for (xt_id in ox_sat_tb.ids) {
                                        xtblist_data[[xt_id]]  <- read.xlsx(xls_tables_paths[[xt_id]], sheetIndex = 2, stringsAsFactors=F)
                                }
                                
                                
                                clean_ox_list <- xtblist_data[!sapply(xtblist_data, is.null)]
                                clean_ox_df <- clean_ox_list[[grep("Diagnostic", clean_ox_list)]][,1:2]
                                colnames(clean_ox_df) <- c("var", "value")
                                
                                clean_ox_df <- filter(clean_ox_df, !is.na(value))
                                clean_ox_df$var <- paste0("OxStats_", clean_ox_df$var)
                                
                        
                                
                                ox_stats_diagnostic_names <- clean_ox_df$var
                                ox_stats_diagnostic <- data.frame(t(clean_ox_df$value))
                                
                                colnames(ox_stats_diagnostic) <- ox_stats_diagnostic_names
                                
                                
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
                                
                                #Remove NA on Var_name and starts with <<
                                xlstb_combined_extraonly <- xlstb2_combined %>% filter(!is.na(Var_name), !grepl("^<<", xlstb2_combined$Var_name), !is.na(label))
                                
                                xlstb_combined_extraonly_lightsonoff <- filter(xlstb_combined_extraonly, grepl("Lights", xlstb_combined_extraonly$Var_name)) %>% distinct(.keep_all = T)
                                xlstb_combined_extraonly_others <- filter(xlstb_combined_extraonly, !grepl("Lights", xlstb_combined_extraonly$Var_name)) %>% distinct(.keep_all = T)
                             
                               
                                
                                #Get unformatted dates
                                xlstb_combined_extraonly_others$value[grepl(":", xlstb_combined_extraonly_others$value)] <- as.vector(sapply(xlstb_combined_extraonly_others$value[grepl(":", xlstb_combined_extraonly_others$value)], process_excel_date))
                                
                                #Convert to numeric
                                xlstb_combined_extraonly_others$value <- as.numeric(as.character(xlstb_combined_extraonly_others$value))
                                
                                
                                
                                #create set of variables to include as extra
                                extrapsgdata_lights <- data.frame(t(xlstb_combined_extraonly_lightsonoff$value))
                                colnames(extrapsgdata_lights) <- paste0("S03_PSG_EXTRA_",t(xlstb_combined_extraonly_lightsonoff$Var_name))
                                extrapsgdata_others <- data.frame(t(xlstb_combined_extraonly_others$value))
                                colnames(extrapsgdata_others) <- paste0("S03_PSG_EXTRA_",t(xlstb_combined_extraonly_others$Var_name))
                                
                               
                                ######### Remove all files from zip_folder to avoid confusion when getting new data
                                file.remove(list.files("zips_s03/", include.dirs = F, full.names = T, recursive = T))
                                
                                ######## Combine ALL data
                                
                                
                                all_fields <- c("S03_PSG_PHI_Filename", "PennSleepID", colnames(word_main_values_df), colnames(word_other_values_df), paste0("S03_PSG_", colnames(ox_stats_diagnostic)), colnames(extrapsgdata_lights), colnames(extrapsgdata_others))
                                
                                #Initiate dataframe with results
                                
                                #If first, create final_df
                                if (first) {
                                        final_df <- data.frame(matrix(NA,1,length(all_fields)))
                                        colnames(final_df) <- all_fields
                                        
                                        final_df[1,] <- c(PSGfilename,
                                                          NA,
                                                          word_main_values_df[1,],
                                                          word_other_values_df[1,],
                                                          ox_stats_diagnostic[1,],
                                                          extrapsgdata_lights[1,],
                                                          extrapsgdata_others[1,])
                                        first <- F
                                } else {
                                        
                                        next_final_df <- data.frame(matrix(NA,1,length(all_fields)))
                                        colnames(next_final_df) <- all_fields
                                        
                                        next_final_df[1,] <- c(PSGfilename,
                                                                NA,
                                                                word_main_values_df[1,],
                                                                word_other_values_df[1,],
                                                               ox_stats_diagnostic[1,],
                                                                extrapsgdata_lights[1,],
                                                                extrapsgdata_others[1,])
                                        
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
        final_df$S03_PSG_Age_at_Study <- interval(mdy(final_df$S03_PSG_PHI_NPBMPatientInfoDOB), mdy(final_df$S03_PSG_NPBMStudyInfoStudyDate)) %/% years(1)
        
        #Create identifiable data frame
        identifiable_df <- final_df
        identifiable_df$ProcessedDate <- Sys.Date()
        
        ProcessedTimeID=gsub("-","",gsub(":","",gsub(" ", "", Sys.time())))
        
        identifiable_df$ProcessedTimeID <- ProcessedTimeID
        
        # Per sample QC (not value filter)
        
        QC_df <- data.frame(S03_PSG_PHI_Filename=identifiable_df$S03_PSG_PHI_Filename,
                            # Missing both MRN fields
                            has_missingMRN=is.na(identifiable_df$S03_PSG_PHI_NPBMSiteInfoHospitalNumber),
                            # Gender not (Male, Female, Unknown)
                            has_missingGender=!(tolower(identifiable_df$S03_PSG_NPBMPatientInfoGender) %in% c("male", "female", "unknown")),
                            # Missing age
                            has_missingAge=is.na(identifiable_df$S03_PSG_Age_at_Study),
                            # Missing BMI
                            has_missingBMI=is.na(identifiable_df$S03_PSG_NPBMPatientInfoBMI)
        )
        
        identifiable_df$PerSample_QC <- NA
        identifiable_df$PerSample_QC[rowSums(QC_df[,2:5])>0] <- "FAIL"
        identifiable_df$PerSample_QC[rowSums(QC_df[,2:5])==0] <- "PASS"
        
        QC_df$PerSample_QC <- identifiable_df$PerSample_QC
        write.csv(QC_df, paste0("QC_dataframe_", ProcessedTimeID, ".csv"))
        
        #Save identifiable Rdata file encrypted
        #key <- key_sodium(sodium::keygen())
        saveRDS(identifiable_df, paste0("PennSleepDatabase_Split03_", ProcessedTimeID, ".Rdata"))
        #cyphr::encrypt_file("PennSleepDatabase.Rdata", key, "PennSleepDatabase.encrypted")
        
        #Save de-identified version in CSV files - implement a way to check in database if sample was processed and not use the same de-identified IDs
        #identifiable_df$PennSleepID <- paste0("PENNSLEEP00000",seq(1:nrow(identifiable_df)))
        deidentified_df <- select(identifiable_df, -starts_with("PSG_PHI_"))
        
        write.csv(deidentified_df, paste0("PennSleepDatabase_Split03_Deidentified",ProcessedTimeID,".csv"), row.names = F)
        
        return(identifiable_df)
        
}

