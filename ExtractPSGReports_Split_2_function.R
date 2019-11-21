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

get_final_diagnosis <- function(table) {
        
        tb_idx <- grep("FINAL DIAGNOSIS:", table)
        
        if (length(tb_idx)==0) {return(NA)}
        
        tb_col_idx <- grep("FINAL DIAGNOSIS:", table[[tb_idx]])
        result <- gsub(" IF(.+)$", "",grep("FINAL DIAGNOSIS:", pull(table[[tb_idx]][,tb_col_idx]), value = T))
        return(result)
        
}

get_comm_rec <- function(table) {
        
        tb_idx <- grep("COMMENTS", table)
        
        if (length(tb_idx)==0) {return(NA)}
        
        tb_col_idx <- grep("COMMENTS", table[[tb_idx]])
        result <-  gsub(' DATE \\\\@ \\"M/d/yyyy\\"', "", gsub(' TIME \\\\@ \\"h:mm am/pm\\"', "", gsub(' DATE \\\\@ \\"M/d/yyyy\\"', "",grep("COMMENTS", pull(table[[tb_idx]][,tb_col_idx]), value = T))))
        return(result)
        
}

#### Define fields from word tables to extract
word_fields <- c(
        "NPBMPatientInfoLastName",
        "NPBMPatientInfoFirstName",
        "NPBMSiteInfoHospitalNumber",
        "NPBMStudyInfoStudyDate",
        "NPBMCustomPatientInfoReferring_Physician",
        "NPBMCustomPatientInfoSleep_Specialist",
        "NPBMSiteInfoSubjectCode",
        "NPBMPatientInfoDOB",
        "NPBMPatientInfoHeight",
        "NPBPatientInfoLengthUnit",
        "NPBMPatientInfoWeight",
        "NPBPatientInfoWeightUnit",
        "NPBMPatientInfoBMI",
        "NPBMPatientInfoGender",
        "NPBMCustomPatientInfoBloodPressure",
        "NPBMCustomPatientInfoBaselineSupine", #Snoring supine
        "NPBMCustomPatientInfoBaselineLateral", #Snoring lateral
        "NPBMCustomPatientInfoBaselineProne" #Snoring prone
)

extract_split_2 <- function(paths) {
        
        #paths <- s02_paths[1]
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
                            dir.create("temp_reports_s02/", showWarnings = FALSE)
                            file.copy(PSGfilename, to = "temp_reports_s02/")
                            
                            #docx <- paste0("temp_reports_s02/", sapply(strsplit(PSGfilename, "/"), "[[", 3))
                            docx <- paste0("temp_reports_s02/", sapply(strsplit(PSGfilename, "/"), "[[", 6))
                            
                            
                            
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
                                word_main_values_df$NPBPatientInfoLengthUnit <- strsplit(word_main_values_df$NPBPatientInfoLengthUnit," ")[[1]][length(strsplit(word_main_values_df$NPBPatientInfoLengthUnit," ")[[1]])]
                                word_main_values_df$NPBMPatientInfoWeight <- strsplit(word_main_values_df$NPBMPatientInfoWeight," ")[[1]][1]
                                word_main_values_df$NPBPatientInfoWeightUnit <- strsplit(word_main_values_df$NPBPatientInfoWeightUnit," ")[[1]][length(strsplit(word_main_values_df$NPBPatientInfoWeightUnit," ")[[1]])]
                                word_main_values_df$NPBMPatientInfoBMI <- trim(gsub("kg/m2", "", word_main_values_df$NPBMPatientInfoBMI))
                                
                                # PSG variables - Entire night
                                word_main_values_df$StartTime_entire <- find_field_grep("Start Time :", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$EndTime_entire <- find_field_grep("End Time :", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$TotalRecordingTime_entire <- find_field_grep("Total Recording Time (minutes) :", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$TotalSleepTime_entire <- find_field_grep("Total Sleep Time (minutes) :", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$SleepEfficiency_entire <- find_field_grep("Sleep Efficiency (%)", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$SleepOnsetLatency_entire <- find_field_grep("Sleep Onset Latency (minutes) :", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$NumberREMPeriods_entire <- find_field_grep("Number of REM Periods :", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$REMLatency_entire <- find_field_grep("REM Latency :", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$WASO_entire <- find_field_grep("WASO:", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$StageN1_entire <- find_field_grep("Stage N1:", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$StageN2_entire <- find_field_grep("Stage N2:", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$StageN3_entire <- find_field_grep("Stage N3", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$StageREM_entire <- find_field_grep("REM:", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$StageN1_pct_entire <- trim(gsub("%", "", find_field_grep("Stage N1:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=2)))
                                word_main_values_df$StageN2_pct_entire <- trim(gsub("%", "", find_field_grep("Stage N2:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=2)))
                                word_main_values_df$StageN3_pct_entire <- trim(gsub("%", "", find_field_grep("Stage N3", TABLES, type = "col_next", tb.pos=1, fixed = T, n=2)))
                                word_main_values_df$StageREM_pct_entire <- trim(gsub("%", "", find_field_grep("REM:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=2)))
                                
                                
                                # PSG variables - Baseline
                                word_main_values_df$StartTime_baseline <- find_field_grep("Start Time :", TABLES, type = "col_next", tb.pos=2, fixed = T)
                                word_main_values_df$EndTime_baseline <- find_field_grep("End Time :", TABLES, type = "col_next", tb.pos=2, fixed = T)
                                word_main_values_df$TotalRecordingTime_baseline <- find_field_grep("Total Recording Time (minutes) :", TABLES, type = "col_next", tb.pos=2, fixed = T)
                                word_main_values_df$TotalSleepTime_baseline <- find_field_grep("Total Sleep Time (minutes) :", TABLES, type = "col_next", tb.pos=2, fixed = T)
                                word_main_values_df$SleepEfficiency_baseline <- find_field_grep("Sleep Efficiency (%)", TABLES, type = "col_next", tb.pos=2, fixed = T)
                                word_main_values_df$SleepOnsetLatency_baseline <- find_field_grep("Sleep Onset Latency (minutes) :", TABLES, type = "col_next", tb.pos=2, fixed = T)
                                word_main_values_df$NumberREMPeriods_baseline <- find_field_grep("Number of REM Periods :", TABLES, type = "col_next", tb.pos=2, fixed = T)
                                word_main_values_df$REMLatency_baseline <- find_field_grep("REM Latency :", TABLES, type = "col_next", tb.pos=2, fixed = T)
                                word_main_values_df$WASO_baseline <- find_field_grep("WASO:", TABLES, type = "col_next", tb.pos=2, fixed = T)
                                word_main_values_df$StageN1_baseline <- find_field_grep("Stage N1:", TABLES, type = "col_next", tb.pos=2, fixed = T)
                                word_main_values_df$StageN2_baseline <- find_field_grep("Stage N2:", TABLES, type = "col_next", tb.pos=2, fixed = T)
                                word_main_values_df$StageN3_baseline <- find_field_grep("Stage N3", TABLES, type = "col_next", tb.pos=2, fixed = T)
                                word_main_values_df$StageREM_baseline <- find_field_grep("REM:", TABLES, type = "col_next", tb.pos=2, fixed = T)
                                word_main_values_df$StageN1_pct_baseline <- trim(gsub("%", "", find_field_grep("Stage N1:", TABLES, type = "col_next", tb.pos=2, fixed = T, n=2)))
                                word_main_values_df$StageN2_pct_baseline <- trim(gsub("%", "", find_field_grep("Stage N2:", TABLES, type = "col_next", tb.pos=2, fixed = T, n=2)))
                                word_main_values_df$StageN3_pct_baseline <- trim(gsub("%", "", find_field_grep("Stage N3", TABLES, type = "col_next", tb.pos=2, fixed = T, n=2)))
                                word_main_values_df$StageREM_pct_baseline <- trim(gsub("%", "", find_field_grep("REM:", TABLES, type = "col_next", tb.pos=2, fixed = T, n=2)))
                                
                                # Apnea & Hypopnea Events - Baseline
                                word_main_values_df$N_Central_Events_baseline <- find_field_grep(pattern = "CENTRAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=1)
                                word_main_values_df$Idx_Central_Events_baseline <- find_field_grep(pattern = "CENTRAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=2)
                                word_main_values_df$MeanDur_Central_Events_baseline <- find_field_grep(pattern = "CENTRAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=3)
                                word_main_values_df$LongestDur_Central_Events_baseline <- find_field_grep(pattern = "CENTRAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=4)
                                word_main_values_df$N_Central_Events_REM_baseline <- find_field_grep(pattern = "CENTRAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=5)
                                word_main_values_df$N_Central_Events_NREM_baseline <- find_field_grep(pattern = "CENTRAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=6)
                                word_main_values_df$Idx_Central_Events_REM_baseline <- find_field_grep(pattern = "CENTRAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=7)
                                word_main_values_df$Idx_Central_Events_NREM_baseline <- find_field_grep(pattern = "CENTRAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=8)
                                
                                word_main_values_df$N_Obstructive_Events_baseline <- find_field_grep(pattern = "OBSTRUCTIVE", TABLES, type = "row_below", tb.pos=1, fixed = T, n=1)
                                word_main_values_df$Idx_Obstructive_Events_baseline <- find_field_grep(pattern = "OBSTRUCTIVE", TABLES, type = "row_below", tb.pos=1, fixed = T, n=2)
                                word_main_values_df$MeanDur_Obstructive_Events_baseline <- find_field_grep(pattern = "OBSTRUCTIVE", TABLES, type = "row_below", tb.pos=1, fixed = T, n=3)
                                word_main_values_df$LongestDur_Obstructive_Events_baseline <- find_field_grep(pattern = "OBSTRUCTIVE", TABLES, type = "row_below", tb.pos=1, fixed = T, n=4)
                                word_main_values_df$N_Obstructive_Events_REM_baseline <- find_field_grep(pattern = "OBSTRUCTIVE", TABLES, type = "row_below", tb.pos=1, fixed = T, n=5)
                                word_main_values_df$N_Obstructive_Events_NREM_baseline <- find_field_grep(pattern = "OBSTRUCTIVE", TABLES, type = "row_below", tb.pos=1, fixed = T, n=6)
                                word_main_values_df$Idx_Obstructive_Events_REM_baseline <- find_field_grep(pattern = "OBSTRUCTIVE", TABLES, type = "row_below", tb.pos=1, fixed = T, n=7)
                                word_main_values_df$Idx_Obstructive_Events_NREM_baseline <- find_field_grep(pattern = "OBSTRUCTIVE", TABLES, type = "row_below", tb.pos=1, fixed = T, n=8)
                                
                                word_main_values_df$N_Mixed_Events_baseline <- find_field_grep(pattern = "MIXED", TABLES, type = "row_below", tb.pos=1, fixed = T, n=1)
                                word_main_values_df$Idx_Mixed_Events_baseline <- find_field_grep(pattern = "MIXED", TABLES, type = "row_below", tb.pos=1, fixed = T, n=2)
                                word_main_values_df$MeanDur_Mixed_Events_baseline <- find_field_grep(pattern = "MIXED", TABLES, type = "row_below", tb.pos=1, fixed = T, n=3)
                                word_main_values_df$LongestDur_Mixed_Events_baseline <- find_field_grep(pattern = "MIXED", TABLES, type = "row_below", tb.pos=1, fixed = T, n=4)
                                word_main_values_df$N_Mixed_Events_REM_baseline <- find_field_grep(pattern = "MIXED", TABLES, type = "row_below", tb.pos=1, fixed = T, n=5)
                                word_main_values_df$N_Mixed_Events_NREM_baseline <- find_field_grep(pattern = "MIXED", TABLES, type = "row_below", tb.pos=1, fixed = T, n=6)
                                word_main_values_df$Idx_Mixed_Events_REM_baseline <- find_field_grep(pattern = "MIXED", TABLES, type = "row_below", tb.pos=1, fixed = T, n=7)
                                word_main_values_df$Idx_Mixed_Events_NREM_baseline <- find_field_grep(pattern = "MIXED", TABLES, type = "row_below", tb.pos=1, fixed = T, n=8)
                                
                                word_main_values_df$N_Apnea_Events_baseline <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=1)
                                word_main_values_df$Idx_Apnea_Events_baseline <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=2)
                                word_main_values_df$MeanDur_Apnea_Events_baseline <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=3)
                                word_main_values_df$LongestDur_Apnea_Events_baseline <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=4)
                                word_main_values_df$N_Apnea_Events_REM_baseline <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=5)
                                word_main_values_df$N_Apnea_Events_NREM_baseline <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=6)
                                word_main_values_df$Idx_Apnea_Events_REM_baseline <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=7)
                                word_main_values_df$Idx_Apnea_Events_NREM_baseline <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=8)
                                
                                word_main_values_df$N_Hypopneas_Events_baseline <- find_field_grep(pattern = "HYPOPNEAS", TABLES, type = "row_below", tb.pos=1, fixed = T, n=1)
                                word_main_values_df$Idx_Hypopneas_Events_baseline <- find_field_grep(pattern = "HYPOPNEAS", TABLES, type = "row_below", tb.pos=1, fixed = T, n=2)
                                word_main_values_df$MeanDur_Hypopneas_Events_baseline <- find_field_grep(pattern = "HYPOPNEAS", TABLES, type = "row_below", tb.pos=1, fixed = T, n=3)
                                word_main_values_df$LongestDur_Hypopneas_Events_baseline <- find_field_grep(pattern = "HYPOPNEAS", TABLES, type = "row_below", tb.pos=1, fixed = T, n=4)
                                word_main_values_df$N_Hypopneas_Events_REM_baseline <- find_field_grep(pattern = "HYPOPNEAS", TABLES, type = "row_below", tb.pos=1, fixed = T, n=5)
                                word_main_values_df$N_Hypopneas_Events_NREM_baseline <- find_field_grep(pattern = "HYPOPNEAS", TABLES, type = "row_below", tb.pos=1, fixed = T, n=6)
                                word_main_values_df$Idx_Hypopneas_Events_REM_baseline <- find_field_grep(pattern = "HYPOPNEAS", TABLES, type = "row_below", tb.pos=1, fixed = T, n=7)
                                word_main_values_df$Idx_Hypopneas_Events_NREM_baseline <- find_field_grep(pattern = "HYPOPNEAS", TABLES, type = "row_below", tb.pos=1, fixed = T, n=8)
                                
                                # Respiratory Events and Body Position
                                word_main_values_df$Total_AHI_baseline <- find_field_grep(pattern = "INDEX", TABLES, type = "row_below", tb.pos=1, fixed = T, n=1)
                                word_main_values_df$Total_AH_count_baseline <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=1)
                                word_main_values_df$NREM_AHI_baseline <- find_field_grep(pattern = "INDEX", TABLES, type = "row_below", tb.pos=1, fixed = T, n=2)
                                word_main_values_df$NREM_AH_count_baseline <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=2)
                                word_main_values_df$REM_AHI_baseline <- find_field_grep(pattern = "INDEX", TABLES, type = "row_below", tb.pos=1, fixed = T, n=3)
                                word_main_values_df$REM_AH_count_baseline <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=3)
                                word_main_values_df$Supine_AHI_baseline <- find_field_grep(pattern = "INDEX", TABLES, type = "row_below", tb.pos=1, fixed = T, n=4)
                                word_main_values_df$Supine_AH_count_baseline <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=4)
                                word_main_values_df$Lateral_AHI_baseline <- find_field_grep(pattern = "INDEX", TABLES, type = "row_below", tb.pos=1, fixed = T, n=5)
                                word_main_values_df$Lateral_AH_count_baseline <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=5)
                                word_main_values_df$Prone_AHI_baseline <- find_field_grep(pattern = "INDEX", TABLES, type = "row_below", tb.pos=1, fixed = T, n=6)
                                word_main_values_df$Prone_AH_count_baseline <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=6)
                                word_main_values_df$Left_AHI_baseline <- find_field_grep(pattern = "INDEX", TABLES, type = "row_below", tb.pos=1, fixed = T, n=7)
                                word_main_values_df$Left_AH_count_baseline <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=7)
                                word_main_values_df$Right_AHI_baseline <- find_field_grep(pattern = "INDEX", TABLES, type = "row_below", tb.pos=1, fixed = T, n=8)
                                word_main_values_df$Right_AH_count_baseline <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=8)
                                
                                #PLM Events 
                                word_main_values_df$PLM_Total_Idx_baseline <- find_field_grep(pattern = "Total Events:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=1)
                                word_main_values_df$PLM_Total_Idx_NREM_baseline <- find_field_grep(pattern = "Non-REM Events:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=1)
                                word_main_values_df$PLM_Total_Idx_REM_baseline <- strsplit(pull(TABLES[[grep("REM Events:", TABLES)[1]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),2]), " ")[[1]][length(strsplit(pull(TABLES[[grep("REM Events:", TABLES)[1]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),2]), " ")[[1]])]
                                
                                word_main_values_df$PLM_Total_counts_baseline <- find_field_grep(pattern = "Total Events:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=2)
                                word_main_values_df$PLM_Total_counts_NREM_baseline <- find_field_grep(pattern = "Non-REM Events:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=2)
                                word_main_values_df$PLM_Total_counts_REM_baseline <- strsplit(pull(TABLES[[grep("REM Events:", TABLES)[1]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),3]), " ")[[1]][length(strsplit(pull(TABLES[[grep("REM Events:", TABLES)[1]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),3]), " ")[[1]])]
                                
                                word_main_values_df$PLM_withArousal_Idx_baseline <- find_field_grep(pattern = "Total Events:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=3)
                                word_main_values_df$PLM_withArousal_Idx_NREM_baseline <- find_field_grep(pattern = "Non-REM Events:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=3)
                                word_main_values_df$PLM_withArousal_Idx_REM_baseline <-strsplit(pull(TABLES[[grep("REM Events:", TABLES)[1]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),4]), " ")[[1]][length(strsplit(pull(TABLES[[grep("REM Events:", TABLES)[1]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),4]), " ")[[1]])]
                                
                                word_main_values_df$PLM_withArousal_counts_baseline <- find_field_grep(pattern = "Total Events:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=4)
                                word_main_values_df$PLM_withArousal_counts_NREM_baseline <- find_field_grep(pattern = "Non-REM Events:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=4)
                                word_main_values_df$PLM_withArousal_counts_REM_baseline <- strsplit(pull(TABLES[[grep("REM Events:", TABLES)[1]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),5]), " ")[[1]][length(strsplit(pull(TABLES[[grep("REM Events:", TABLES)[1]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),5]), " ")[[1]])]
                                
                                word_main_values_df$PLM_withoutArousal_Idx_baseline <- find_field_grep(pattern = "Total Events:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=5)
                                word_main_values_df$PLM_withoutArousal_Idx_NREM_baseline <- find_field_grep(pattern = "Non-REM Events:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=5)
                                word_main_values_df$PLM_withoutArousal_Idx_REM_baseline <- strsplit(pull(TABLES[[grep("REM Events:", TABLES)[1]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),6]), " ")[[1]][length(strsplit(pull(TABLES[[grep("REM Events:", TABLES)[1]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),6]), " ")[[1]])]
                                
                                word_main_values_df$PLM_withoutArousal_counts_baseline <- find_field_grep(pattern = "Total Events:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=6)
                                word_main_values_df$PLM_withoutArousal_counts_NREM_baseline <- find_field_grep(pattern = "Non-REM Events:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=6)
                                word_main_values_df$PLM_withoutArousal_counts_REM_baseline <- strsplit(pull(TABLES[[grep("REM Events:", TABLES)[1]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),7]), " ")[[1]][length(strsplit(pull(TABLES[[grep("REM Events:", TABLES)[1]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),7]), " ")[[1]])]
                                
                                # All arousals
                                word_main_values_df$Arousal_Idx_baseline <- find_field_grep(pattern = "Total Events:", TABLES, type = "col_next", tb.pos=2, fixed = T, n=1)
                                word_main_values_df$Arousal_Idx_NREM_baseline <- find_field_grep(pattern = "Non-REM Events:", TABLES, type = "col_next", tb.pos=2, fixed = T, n=1)
                                word_main_values_df$Arousal_Idx_REM_baseline <- strsplit(pull(TABLES[[grep("REM Events:", TABLES)[2]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[2]]]$V1),2]), " ")[[1]][length(strsplit(pull(TABLES[[grep("REM Events:", TABLES)[2]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[2]]]$V1),2]), " ")[[1]])]
                                
                                word_main_values_df$Arousal_counts_baseline <- find_field_grep(pattern = "Total Events:", TABLES, type = "col_next", tb.pos=2, fixed = T, n=2)
                                word_main_values_df$Arousal_counts_NREM_baseline <- find_field_grep(pattern = "Non-REM Events:", TABLES, type = "col_next", tb.pos=2, fixed = T, n=2)
                                word_main_values_df$Arousal_counts_REM_baseline <- strsplit(pull(TABLES[[grep("REM Events:", TABLES)[2]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[2]]]$V1),3]), " ")[[1]][length(strsplit(pull(TABLES[[grep("REM Events:", TABLES)[2]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[2]]]$V1),3]), " ")[[1]])]
                                
                                
                                # Oxygen Saturation
                                word_main_values_df$Mean_SaO2_awake_perc_baseline=as.numeric(find_field_grep(pattern = "Mean SaO2",
                                                                                                 pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 1, n=1))
                                word_main_values_df$Mean_SaO2_NREM_perc_baseline=as.numeric(find_field_grep(pattern = "Mean SaO2",
                                                                                                pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 1, n=1))
                                word_main_values_df$Mean_SaO2_REM_perc_baseline=as.numeric(find_field_grep(pattern = "Mean SaO2",
                                                                                               pattern2 = "Mean SaO2", table = TABLES, type = "col_next", tb.pos = 1, n=3))
                                word_main_values_df$Mean_SaO2_TST_perc_baseline=as.numeric(find_field_grep(pattern = "Mean SaO2",
                                                                                               pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 1, n=1))
                                
                                word_main_values_df$Min_SaO2_awake_perc_baseline=as.numeric(find_field_grep(pattern = "Min. SaO2",
                                                                                                pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 1, n=2))
                                word_main_values_df$Min_SaO2_NREM_perc_baseline=as.numeric(find_field_grep(pattern = "Min. SaO2",
                                                                                               pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 1, n=2))
                                word_main_values_df$Min_SaO2_REM_perc_baseline=as.numeric(find_field_grep(pattern = "Min. SaO2",
                                                                                              pattern2 = "Min. SaO2", table = TABLES, type = "col_next", tb.pos = 1, n=3))
                                word_main_values_df$Min_SaO2_TST_perc_baseline=as.numeric(find_field_grep(pattern = "Min. SaO2",
                                                                                              pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 1, n=2))
                                
                                word_main_values_df$Max_SaO2_awake_perc_baseline=as.numeric(find_field_grep(pattern = "Max. SaO2",
                                                                                                pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 1, n=3))
                                word_main_values_df$Max_SaO2_NREM_perc_baseline=as.numeric(find_field_grep(pattern = "Max. SaO2",
                                                                                               pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 1, n=3))
                                word_main_values_df$Max_SaO2_REM_perc_baseline=as.numeric(find_field_grep(pattern = "Max. SaO2",
                                                                                              pattern2 = "Max. SaO2", table = TABLES, type = "col_next", tb.pos = 1, n=3))
                                word_main_values_df$Max_SaO2_TST_perc_baseline=as.numeric(find_field_grep(pattern = "Max. SaO2",
                                                                                              pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 1, n=3))
                                
                                word_main_values_df$SaO2_90.100_awake_perc_baseline=as.numeric(gsub("%","",find_field_grep(pattern = "90.+100 %",
                                                                                                   pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 1, n=5, fixed=F)))
                                word_main_values_df$SaO2_90.100_NREM_perc_baseline=as.numeric(gsub("%","",find_field_grep(pattern = "90.+100 %",
                                                                                                  pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 1, n=5, fixed=F)))
                                word_main_values_df$SaO2_90.100_REM_perc_baseline=as.numeric(gsub("%","",find_field_grep(pattern = "90.+100 %",
                                                                                                 pattern2 = "90.+100 %", table = TABLES, type = "col_next", tb.pos = 1, n=3, fixed=F)))
                                word_main_values_df$SaO2_90.100_TST_perc_baseline=as.numeric(gsub("%","",find_field_grep(pattern = "90.+100 %",
                                                                                                 pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 1, n=5, fixed=F)))
                                
                                word_main_values_df$SaO2_80.89_awake_perc_baseline=as.numeric(gsub("%","",find_field_grep(pattern = "80.+89 %",
                                                                                                  pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 1, n=6, fixed=F)))
                                word_main_values_df$SaO2_80.89_NREM_perc_baseline=as.numeric(gsub("%","",find_field_grep(pattern = "80.+89 %",
                                                                                                 pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 1, n=6, fixed=F)))
                                word_main_values_df$SaO2_80.89_REM_perc_baseline=as.numeric(gsub("%","",find_field_grep(pattern = "80.+89 %",
                                                                                                pattern2 = "80.+89 %", table = TABLES, type = "col_next", tb.pos = 1, n=3, fixed=F)))
                                word_main_values_df$SaO2_80.89_TST_perc_baseline=as.numeric(gsub("%","",find_field_grep(pattern = "80.+89 %",
                                                                                                pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 1, n=6, fixed=F)))
                                
                                word_main_values_df$SaO2_70.79_awake_perc_baseline=as.numeric(gsub("%","",find_field_grep(pattern = "70.+79 %",
                                                                                                  pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 1, n=7, fixed=F)))
                                word_main_values_df$SaO2_70.79_NREM_perc_baseline=as.numeric(gsub("%","",find_field_grep(pattern = "70.+79 %",
                                                                                                 pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 1, n=7, fixed=F)))
                                word_main_values_df$SaO2_70.79_REM_perc_baseline=as.numeric(gsub("%","",find_field_grep(pattern = "70.+79 %",
                                                                                                pattern2 = "70.+79 %", table = TABLES, type = "col_next", tb.pos = 1, n=3, fixed=F)))
                                word_main_values_df$SaO2_70.79_TST_perc_baseline=as.numeric(gsub("%","",find_field_grep(pattern = "70.+79 %",
                                                                                                pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 1, n=7, fixed=F)))
                                
                                word_main_values_df$SaO2_60.69_awake_perc_baseline=as.numeric(gsub("%","",find_field_grep(pattern = "60.+69 %",
                                                                                                  pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 1, n=8, fixed=F)))
                                word_main_values_df$SaO2_60.69_NREM_perc_baseline=as.numeric(gsub("%","",find_field_grep(pattern = "60.+69 %",
                                                                                                 pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 1, n=8, fixed=F)))
                                word_main_values_df$SaO2_60.69_REM_perc_baseline=as.numeric(gsub("%","",find_field_grep(pattern = "60.+69 %",
                                                                                                pattern2 = "60.+69 %", table = TABLES, type = "col_next", tb.pos = 1, n=3, fixed=F)))
                                word_main_values_df$SaO2_60.69_TST_perc_baseline=as.numeric(gsub("%","",find_field_grep(pattern = "60.+69 %",
                                                                                                pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 1, n=8, fixed=F)))
                                
                                word_main_values_df$SaO2_50.59_awake_perc_baseline=as.numeric(gsub("%","",find_field_grep(pattern = "50.+59 %",
                                                                                                  pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 1, n=9, fixed=F)))
                                word_main_values_df$SaO2_50.59_NREM_perc_baseline=as.numeric(gsub("%","",find_field_grep(pattern = "50.+59 %",
                                                                                                 pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 1, n=9, fixed=F)))
                                word_main_values_df$SaO2_50.59_REM_perc_baseline=as.numeric(gsub("%","",find_field_grep(pattern = "50.+59 %",
                                                                                                pattern2 = "50.+59 %", table = TABLES, type = "col_next", tb.pos = 1, n=3, fixed=F)))
                                word_main_values_df$SaO2_50.59_TST_perc_baseline=as.numeric(gsub("%","",find_field_grep(pattern = "50.+59 %",
                                                                                                pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 1, n=9, fixed=F)))
                                
                                word_main_values_df$SaO2_Below50_awake_perc_baseline=as.numeric(gsub("%","",find_field_grep(pattern = "Below 50 %",
                                                                                                    pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos = 1, n=10)))
                                word_main_values_df$SaO2_Below50_NREM_perc_baseline=as.numeric(gsub("%","",find_field_grep(pattern = "Below 50 %",
                                                                                                   pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos = 1, n=10)))
                                word_main_values_df$SaO2_Below50REM_perc_baseline=as.numeric(gsub("%","",find_field_grep(pattern = "Below 50 %",
                                                                                                 pattern2 = "Below 50 %", table = TABLES, type = "col_next", tb.pos = 1, n=3)))
                                word_main_values_df$SaO2_Below50_TST_perc_baseline=as.numeric(gsub("%","",find_field_grep(pattern = "Below 50 %",
                                                                                                  pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos = 1, n=10)))
                                
                                
                    
                              ########## Treatment Part
                                # PSG variables - Treatment
                                word_main_values_df$StartTime_treatment <- find_field_grep("Start Time :", TABLES, type = "col_next", tb.pos=3, fixed = T)
                                word_main_values_df$EndTime_treatment <- find_field_grep("End Time :", TABLES, type = "col_next", tb.pos=3, fixed = T)
                                word_main_values_df$TotalRecordingTime_treatment <- find_field_grep("Total Recording Time (minutes) :", TABLES, type = "col_next", tb.pos=3, fixed = T)
                                word_main_values_df$TotalSleepTime_treatment <- find_field_grep("Total Sleep Time (minutes) :", TABLES, type = "col_next", tb.pos=3, fixed = T)
                                word_main_values_df$SleepEfficiency_treatment <- find_field_grep("Sleep Efficiency (%)", TABLES, type = "col_next", tb.pos=3, fixed = T)
                                word_main_values_df$SleepOnsetLatency_treatment <- find_field_grep("Sleep Onset Latency (minutes) :", TABLES, type = "col_next", tb.pos=3, fixed = T)
                                word_main_values_df$NumberREMPeriods_treatment <- find_field_grep("Number of REM Periods :", TABLES, type = "col_next", tb.pos=3, fixed = T)
                                word_main_values_df$REMLatency_treatment <- find_field_grep("REM Latency :", TABLES, type = "col_next", tb.pos=3, fixed = T)
                                word_main_values_df$WASO_treatment <- find_field_grep("WASO:", TABLES, type = "col_next", tb.pos=3, fixed = T)
                                word_main_values_df$StageN1_treatment <- find_field_grep("Stage N1:", TABLES, type = "col_next", tb.pos=3, fixed = T)
                                word_main_values_df$StageN2_treatment <- find_field_grep("Stage N2:", TABLES, type = "col_next", tb.pos=3, fixed = T)
                                word_main_values_df$StageN3_treatment <- find_field_grep("Stage N3", TABLES, type = "col_next", tb.pos=3, fixed = T)
                                word_main_values_df$StageREM_treatment <- find_field_grep("REM:", TABLES, type = "col_next", tb.pos=4, fixed = T)
                                word_main_values_df$StageN1_pct_treatment <- trim(gsub("%", "", find_field_grep("Stage N1:", TABLES, type = "col_next", tb.pos=3, fixed = T, n=2)))
                                word_main_values_df$StageN2_pct_treatment <- trim(gsub("%", "", find_field_grep("Stage N2:", TABLES, type = "col_next", tb.pos=3, fixed = T, n=2)))
                                word_main_values_df$StageN3_pct_treatment <- trim(gsub("%", "", find_field_grep("Stage N3", TABLES, type = "col_next", tb.pos=3, fixed = T, n=2)))
                                word_main_values_df$StageREM_pct_treatment <- trim(gsub("%", "", find_field_grep("REM:", TABLES, type = "col_next", tb.pos=4, fixed = T, n=2)))
                                
                                # Apnea & Hypopnea Events - Treatment
                                word_main_values_df$N_Central_Events_treatment <- find_field_grep(pattern = "CENTRAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=1)
                                word_main_values_df$Idx_Central_Events_treatment <- find_field_grep(pattern = "CENTRAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=2)
                                word_main_values_df$MeanDur_Central_Events_treatment <- find_field_grep(pattern = "CENTRAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=3)
                                word_main_values_df$LongestDur_Central_Events_treatment <- find_field_grep(pattern = "CENTRAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=4)
                                word_main_values_df$N_Central_Events_REM_treatment <- find_field_grep(pattern = "CENTRAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=5)
                                word_main_values_df$N_Central_Events_NREM_treatment <- find_field_grep(pattern = "CENTRAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=6)
                                word_main_values_df$Idx_Central_Events_REM_treatment <- find_field_grep(pattern = "CENTRAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=7)
                                word_main_values_df$Idx_Central_Events_NREM_treatment <- find_field_grep(pattern = "CENTRAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=8)
                                
                                word_main_values_df$N_Obstructive_Events_treatment <- find_field_grep(pattern = "OBSTRUCTIVE", TABLES, type = "row_below", tb.pos=2, fixed = T, n=1)
                                word_main_values_df$Idx_Obstructive_Events_treatment <- find_field_grep(pattern = "OBSTRUCTIVE", TABLES, type = "row_below", tb.pos=2, fixed = T, n=2)
                                word_main_values_df$MeanDur_Obstructive_Events_treatment <- find_field_grep(pattern = "OBSTRUCTIVE", TABLES, type = "row_below", tb.pos=2, fixed = T, n=3)
                                word_main_values_df$LongestDur_Obstructive_Events_treatment <- find_field_grep(pattern = "OBSTRUCTIVE", TABLES, type = "row_below", tb.pos=2, fixed = T, n=4)
                                word_main_values_df$N_Obstructive_Events_REM_treatment <- find_field_grep(pattern = "OBSTRUCTIVE", TABLES, type = "row_below", tb.pos=2, fixed = T, n=5)
                                word_main_values_df$N_Obstructive_Events_NREM_treatment <- find_field_grep(pattern = "OBSTRUCTIVE", TABLES, type = "row_below", tb.pos=2, fixed = T, n=6)
                                word_main_values_df$Idx_Obstructive_Events_REM_treatment <- find_field_grep(pattern = "OBSTRUCTIVE", TABLES, type = "row_below", tb.pos=2, fixed = T, n=7)
                                word_main_values_df$Idx_Obstructive_Events_NREM_treatment <- find_field_grep(pattern = "OBSTRUCTIVE", TABLES, type = "row_below", tb.pos=2, fixed = T, n=8)
                                
                                word_main_values_df$N_Mixed_Events_treatment <- find_field_grep(pattern = "MIXED", TABLES, type = "row_below", tb.pos=2, fixed = T, n=1)
                                word_main_values_df$Idx_Mixed_Events_treatment <- find_field_grep(pattern = "MIXED", TABLES, type = "row_below", tb.pos=2, fixed = T, n=2)
                                word_main_values_df$MeanDur_Mixed_Events_treatment <- find_field_grep(pattern = "MIXED", TABLES, type = "row_below", tb.pos=2, fixed = T, n=3)
                                word_main_values_df$LongestDur_Mixed_Events_treatment <- find_field_grep(pattern = "MIXED", TABLES, type = "row_below", tb.pos=2, fixed = T, n=4)
                                word_main_values_df$N_Mixed_Events_REM_treatment <- find_field_grep(pattern = "MIXED", TABLES, type = "row_below", tb.pos=2, fixed = T, n=5)
                                word_main_values_df$N_Mixed_Events_NREM_treatment <- find_field_grep(pattern = "MIXED", TABLES, type = "row_below", tb.pos=2, fixed = T, n=6)
                                word_main_values_df$Idx_Mixed_Events_REM_treatment <- find_field_grep(pattern = "MIXED", TABLES, type = "row_below", tb.pos=2, fixed = T, n=7)
                                word_main_values_df$Idx_Mixed_Events_NREM_treatment <- find_field_grep(pattern = "MIXED", TABLES, type = "row_below", tb.pos=2, fixed = T, n=8)
                                
                                word_main_values_df$N_Apnea_Events_treatment <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=1)
                                word_main_values_df$Idx_Apnea_Events_treatment <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=2)
                                word_main_values_df$MeanDur_Apnea_Events_treatment <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=3)
                                word_main_values_df$LongestDur_Apnea_Events_treatment <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=4)
                                word_main_values_df$N_Apnea_Events_REM_treatment <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=5)
                                word_main_values_df$N_Apnea_Events_NREM_treatment <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=6)
                                word_main_values_df$Idx_Apnea_Events_REM_treatment <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=7)
                                word_main_values_df$Idx_Apnea_Events_NREM_treatment <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=8)
                                
                                word_main_values_df$N_Hypopneas_Events_treatment <- find_field_grep(pattern = "HYPOPNEAS", TABLES, type = "row_below", tb.pos=2, fixed = T, n=1)
                                word_main_values_df$Idx_Hypopneas_Events_treatment <- find_field_grep(pattern = "HYPOPNEAS", TABLES, type = "row_below", tb.pos=2, fixed = T, n=2)
                                word_main_values_df$MeanDur_Hypopneas_Events_treatment <- find_field_grep(pattern = "HYPOPNEAS", TABLES, type = "row_below", tb.pos=2, fixed = T, n=3)
                                word_main_values_df$LongestDur_Hypopneas_Events_treatment <- find_field_grep(pattern = "HYPOPNEAS", TABLES, type = "row_below", tb.pos=2, fixed = T, n=4)
                                word_main_values_df$N_Hypopneas_Events_REM_treatment <- find_field_grep(pattern = "HYPOPNEAS", TABLES, type = "row_below", tb.pos=2, fixed = T, n=5)
                                word_main_values_df$N_Hypopneas_Events_NREM_treatment <- find_field_grep(pattern = "HYPOPNEAS", TABLES, type = "row_below", tb.pos=2, fixed = T, n=6)
                                word_main_values_df$Idx_Hypopneas_Events_REM_treatment <- find_field_grep(pattern = "HYPOPNEAS", TABLES, type = "row_below", tb.pos=2, fixed = T, n=7)
                                word_main_values_df$Idx_Hypopneas_Events_NREM_treatment <- find_field_grep(pattern = "HYPOPNEAS", TABLES, type = "row_below", tb.pos=2, fixed = T, n=8)
                                
                                # Respiratory Events and Body Position
                                word_main_values_df$Total_AHI_treatment <- find_field_grep(pattern = "INDEX", TABLES, type = "row_below", tb.pos=4, fixed = T, n=1)
                                word_main_values_df$Total_AH_count_treatment <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=6, fixed = T, n=1)
                                word_main_values_df$NREM_AHI_treatment <- find_field_grep(pattern = "INDEX", TABLES, type = "row_below", tb.pos=4, fixed = T, n=2)
                                word_main_values_df$NREM_AH_count_treatment <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=6, fixed = T, n=2)
                                word_main_values_df$REM_AHI_treatment <- find_field_grep(pattern = "INDEX", TABLES, type = "row_below", tb.pos=4, fixed = T, n=3)
                                word_main_values_df$REM_AH_count_treatment <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=6, fixed = T, n=3)
                                word_main_values_df$Supine_AHI_treatment <- find_field_grep(pattern = "INDEX", TABLES, type = "row_below", tb.pos=4, fixed = T, n=4)
                                word_main_values_df$Supine_AH_count_treatment <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=6, fixed = T, n=4)
                                word_main_values_df$Lateral_AHI_treatment <- find_field_grep(pattern = "INDEX", TABLES, type = "row_below", tb.pos=4, fixed = T, n=5)
                                word_main_values_df$Lateral_AH_count_treatment <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=6, fixed = T, n=5)
                                word_main_values_df$Prone_AHI_treatment <- find_field_grep(pattern = "INDEX", TABLES, type = "row_below", tb.pos=4, fixed = T, n=6)
                                word_main_values_df$Prone_AH_count_treatment <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=6, fixed = T, n=6)
                                word_main_values_df$Left_AHI_treatment <- find_field_grep(pattern = "INDEX", TABLES, type = "row_below", tb.pos=4, fixed = T, n=7)
                                word_main_values_df$Left_AH_count_treatment <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=6, fixed = T, n=7)
                                word_main_values_df$Right_AHI_treatment <- find_field_grep(pattern = "INDEX", TABLES, type = "row_below", tb.pos=4, fixed = T, n=8)
                                word_main_values_df$Right_AH_count_treatment <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=6, fixed = T, n=8)
                                
                                #PLM Events 
                                word_main_values_df$PLM_Total_Idx_treatment <- find_field_grep(pattern = "Total Events:", TABLES, type = "col_next", tb.pos=3, fixed = T, n=1)
                                word_main_values_df$PLM_Total_Idx_NREM_treatment <- find_field_grep(pattern = "Non-REM Events:", TABLES, type = "col_next", tb.pos=3, fixed = T, n=1)
                                word_main_values_df$PLM_Total_Idx_REM_treatment <- strsplit(pull(TABLES[[grep("REM Events:", TABLES)[3]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[3]]]$V1),2]), " ")[[1]][length(strsplit(pull(TABLES[[grep("REM Events:", TABLES)[3]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[3]]]$V1),2]), " ")[[1]])]
                                
                                word_main_values_df$PLM_Total_counts_treatment <- find_field_grep(pattern = "Total Events:", TABLES, type = "col_next", tb.pos=3, fixed = T, n=2)
                                word_main_values_df$PLM_Total_counts_NREM_treatment <- find_field_grep(pattern = "Non-REM Events:", TABLES, type = "col_next", tb.pos=3, fixed = T, n=2)
                                word_main_values_df$PLM_Total_counts_REM_treatment <- strsplit(pull(TABLES[[grep("REM Events:", TABLES)[3]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[3]]]$V1),3]), " ")[[1]][length(strsplit(pull(TABLES[[grep("REM Events:", TABLES)[3]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[3]]]$V1),3]), " ")[[1]])]
                                
                                word_main_values_df$PLM_withArousal_Idx_treatment <- find_field_grep(pattern = "Total Events:", TABLES, type = "col_next", tb.pos=3, fixed = T, n=3)
                                word_main_values_df$PLM_withArousal_Idx_NREM_treatment <- find_field_grep(pattern = "Non-REM Events:", TABLES, type = "col_next", tb.pos=3, fixed = T, n=3)
                                word_main_values_df$PLM_withArousal_Idx_REM_treatment <-strsplit(pull(TABLES[[grep("REM Events:", TABLES)[3]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[3]]]$V1),4]), " ")[[1]][length(strsplit(pull(TABLES[[grep("REM Events:", TABLES)[3]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[3]]]$V1),4]), " ")[[1]])]
                                
                                word_main_values_df$PLM_withArousal_counts_treatment <- find_field_grep(pattern = "Total Events:", TABLES, type = "col_next", tb.pos=3, fixed = T, n=4)
                                word_main_values_df$PLM_withArousal_counts_NREM_treatment <- find_field_grep(pattern = "Non-REM Events:", TABLES, type = "col_next", tb.pos=3, fixed = T, n=4)
                                word_main_values_df$PLM_withArousal_counts_REM_treatment <- strsplit(pull(TABLES[[grep("REM Events:", TABLES)[3]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[3]]]$V1),5]), " ")[[1]][length(strsplit(pull(TABLES[[grep("REM Events:", TABLES)[3]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[3]]]$V1),5]), " ")[[1]])]
                                
                                word_main_values_df$PLM_withoutArousal_Idx_treatment <- find_field_grep(pattern = "Total Events:", TABLES, type = "col_next", tb.pos=3, fixed = T, n=5)
                                word_main_values_df$PLM_withoutArousal_Idx_NREM_treatment <- find_field_grep(pattern = "Non-REM Events:", TABLES, type = "col_next", tb.pos=3, fixed = T, n=5)
                                word_main_values_df$PLM_withoutArousal_Idx_REM_treatment <- strsplit(pull(TABLES[[grep("REM Events:", TABLES)[3]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[3]]]$V1),6]), " ")[[1]][length(strsplit(pull(TABLES[[grep("REM Events:", TABLES)[3]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[3]]]$V1),6]), " ")[[1]])]
                                
                                word_main_values_df$PLM_withoutArousal_counts_treatment <- find_field_grep(pattern = "Total Events:", TABLES, type = "col_next", tb.pos=3, fixed = T, n=6)
                                word_main_values_df$PLM_withoutArousal_counts_NREM_treatment <- find_field_grep(pattern = "Non-REM Events:", TABLES, type = "col_next", tb.pos=3, fixed = T, n=6)
                                word_main_values_df$PLM_withoutArousal_counts_REM_treatment <- strsplit(pull(TABLES[[grep("REM Events:", TABLES)[3]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),7]), " ")[[1]][length(strsplit(pull(TABLES[[grep("REM Events:", TABLES)[3]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[3]]]$V1),7]), " ")[[1]])]
                                
                                # All arousals
                                word_main_values_df$Arousal_Idx_treatment <- find_field_grep(pattern = "Total Events:", TABLES, type = "col_next", tb.pos=4, fixed = T, n=1)
                                word_main_values_df$Arousal_Idx_NREM_treatment <- find_field_grep(pattern = "Non-REM Events:", TABLES, type = "col_next", tb.pos=4, fixed = T, n=1)
                                word_main_values_df$Arousal_Idx_REM_treatment <- strsplit(pull(TABLES[[grep("REM Events:", TABLES)[4]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[4]]]$V1),2]), " ")[[1]][length(strsplit(pull(TABLES[[grep("REM Events:", TABLES)[4]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[4]]]$V1),2]), " ")[[1]])]
                                
                                word_main_values_df$Arousal_counts_treatment <- find_field_grep(pattern = "Total Events:", TABLES, type = "col_next", tb.pos=4, fixed = T, n=2)
                                word_main_values_df$Arousal_counts_NREM_treatment <- find_field_grep(pattern = "Non-REM Events:", TABLES, type = "col_next", tb.pos=4, fixed = T, n=2)
                                word_main_values_df$Arousal_counts_REM_treatment <- strsplit(pull(TABLES[[grep("REM Events:", TABLES)[4]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[4]]]$V1),3]), " ")[[1]][length(strsplit(pull(TABLES[[grep("REM Events:", TABLES)[4]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[4]]]$V1),3]), " ")[[1]])]
                                
                                
                                # Oxygen Saturation
                                word_main_values_df$Mean_SaO2_awake_perc_treatment=as.numeric(find_field_grep(pattern = "Mean SaO2",
                                                                                                             pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos=2, n=1))
                                word_main_values_df$Mean_SaO2_NREM_perc_treatment=as.numeric(find_field_grep(pattern = "Mean SaO2",
                                                                                                            pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos=2, n=1))
                                word_main_values_df$Mean_SaO2_REM_perc_treatment=as.numeric(find_field_grep(pattern = "Mean SaO2",
                                                                                                           pattern2 = "Mean SaO2", table = TABLES, type = "col_next", tb.pos=2, n=3))
                                word_main_values_df$Mean_SaO2_TST_perc_treatment=as.numeric(find_field_grep(pattern = "Mean SaO2",
                                                                                                           pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos=2, n=1))
                                
                                word_main_values_df$Min_SaO2_awake_perc_treatment=as.numeric(find_field_grep(pattern = "Min. SaO2",
                                                                                                            pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos=2, n=2))
                                word_main_values_df$Min_SaO2_NREM_perc_treatment=as.numeric(find_field_grep(pattern = "Min. SaO2",
                                                                                                           pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos=2, n=2))
                                word_main_values_df$Min_SaO2_REM_perc_treatment=as.numeric(find_field_grep(pattern = "Min. SaO2",
                                                                                                          pattern2 = "Min. SaO2", table = TABLES, type = "col_next", tb.pos=2, n=3))
                                word_main_values_df$Min_SaO2_TST_perc_treatment=as.numeric(find_field_grep(pattern = "Min. SaO2",
                                                                                                          pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos=2, n=2))
                                
                                word_main_values_df$Max_SaO2_awake_perc_treatment=as.numeric(find_field_grep(pattern = "Max. SaO2",
                                                                                                            pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos=2, n=3))
                                word_main_values_df$Max_SaO2_NREM_perc_treatment=as.numeric(find_field_grep(pattern = "Max. SaO2",
                                                                                                           pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos=2, n=3))
                                word_main_values_df$Max_SaO2_REM_perc_treatment=as.numeric(find_field_grep(pattern = "Max. SaO2",
                                                                                                          pattern2 = "Max. SaO2", table = TABLES, type = "col_next", tb.pos=2, n=3))
                                word_main_values_df$Max_SaO2_TST_perc_treatment=as.numeric(find_field_grep(pattern = "Max. SaO2",
                                                                                                          pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos=2, n=3))
                                
                                word_main_values_df$SaO2_90.100_awake_perc_treatment=as.numeric(gsub("%","",find_field_grep(pattern = "90.+100 %",
                                                                                                                           pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos=2, n=5, fixed=F)))
                                word_main_values_df$SaO2_90.100_NREM_perc_treatment=as.numeric(gsub("%","",find_field_grep(pattern = "90.+100 %",
                                                                                                                          pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos=2, n=5, fixed=F)))
                                word_main_values_df$SaO2_90.100_REM_perc_treatment=as.numeric(gsub("%","",find_field_grep(pattern = "90.+100 %",
                                                                                                                         pattern2 = "90.+100 %", table = TABLES, type = "col_next", tb.pos=2, n=3, fixed=F)))
                                word_main_values_df$SaO2_90.100_TST_perc_treatment=as.numeric(gsub("%","",find_field_grep(pattern = "90.+100 %",
                                                                                                                         pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos=2, n=5, fixed=F)))
                                
                                word_main_values_df$SaO2_80.89_awake_perc_treatment=as.numeric(gsub("%","",find_field_grep(pattern = "80.+89 %",
                                                                                                                          pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos=2, n=6, fixed=F)))
                                word_main_values_df$SaO2_80.89_NREM_perc_treatment=as.numeric(gsub("%","",find_field_grep(pattern = "80.+89 %",
                                                                                                                         pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos=2, n=6, fixed=F)))
                                word_main_values_df$SaO2_80.89_REM_perc_treatment=as.numeric(gsub("%","",find_field_grep(pattern = "80.+89 %",
                                                                                                                        pattern2 = "80.+89 %", table = TABLES, type = "col_next", tb.pos=2, n=3, fixed=F)))
                                word_main_values_df$SaO2_80.89_TST_perc_treatment=as.numeric(gsub("%","",find_field_grep(pattern = "80.+89 %",
                                                                                                                        pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos=2, n=6, fixed=F)))
                                
                                word_main_values_df$SaO2_70.79_awake_perc_treatment=as.numeric(gsub("%","",find_field_grep(pattern = "70.+79 %",
                                                                                                                          pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos=2, n=7, fixed=F)))
                                word_main_values_df$SaO2_70.79_NREM_perc_treatment=as.numeric(gsub("%","",find_field_grep(pattern = "70.+79 %",
                                                                                                                         pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos=2, n=7, fixed=F)))
                                word_main_values_df$SaO2_70.79_REM_perc_treatment=as.numeric(gsub("%","",find_field_grep(pattern = "70.+79 %",
                                                                                                                        pattern2 = "70.+79 %", table = TABLES, type = "col_next", tb.pos=2, n=3, fixed=F)))
                                word_main_values_df$SaO2_70.79_TST_perc_treatment=as.numeric(gsub("%","",find_field_grep(pattern = "70.+79 %",
                                                                                                                        pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos=2, n=7, fixed=F)))
                                
                                word_main_values_df$SaO2_60.69_awake_perc_treatment=as.numeric(gsub("%","",find_field_grep(pattern = "60.+69 %",
                                                                                                                          pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos=2, n=8, fixed=F)))
                                word_main_values_df$SaO2_60.69_NREM_perc_treatment=as.numeric(gsub("%","",find_field_grep(pattern = "60.+69 %",
                                                                                                                         pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos=2, n=8, fixed=F)))
                                word_main_values_df$SaO2_60.69_REM_perc_treatment=as.numeric(gsub("%","",find_field_grep(pattern = "60.+69 %",
                                                                                                                        pattern2 = "60.+69 %", table = TABLES, type = "col_next", tb.pos=2, n=3, fixed=F)))
                                word_main_values_df$SaO2_60.69_TST_perc_treatment=as.numeric(gsub("%","",find_field_grep(pattern = "60.+69 %",
                                                                                                                        pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos=2, n=8, fixed=F)))
                                
                                word_main_values_df$SaO2_50.59_awake_perc_treatment=as.numeric(gsub("%","",find_field_grep(pattern = "50.+59 %",
                                                                                                                          pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos=2, n=9, fixed=F)))
                                word_main_values_df$SaO2_50.59_NREM_perc_treatment=as.numeric(gsub("%","",find_field_grep(pattern = "50.+59 %",
                                                                                                                         pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos=2, n=9, fixed=F)))
                                word_main_values_df$SaO2_50.59_REM_perc_treatment=as.numeric(gsub("%","",find_field_grep(pattern = "50.+59 %",
                                                                                                                        pattern2 = "50.+59 %", table = TABLES, type = "col_next", tb.pos=2, n=3, fixed=F)))
                                word_main_values_df$SaO2_50.59_TST_perc_treatment=as.numeric(gsub("%","",find_field_grep(pattern = "50.+59 %",
                                                                                                                        pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos=2, n=9, fixed=F)))
                                
                                word_main_values_df$SaO2_Below50_awake_perc_treatment=as.numeric(gsub("%","",find_field_grep(pattern = "Below 50 %",
                                                                                                                            pattern2 = "AWAKE", table = TABLES, type = "row_below", tb.pos=2, n=10)))
                                word_main_values_df$SaO2_Below50_NREM_perc_treatment=as.numeric(gsub("%","",find_field_grep(pattern = "Below 50 %",
                                                                                                                           pattern2 = "NREM", table = TABLES, type = "row_below", tb.pos=2, n=10)))
                                word_main_values_df$SaO2_Below50REM_perc_treatment=as.numeric(gsub("%","",find_field_grep(pattern = "Below 50 %",
                                                                                                                         pattern2 = "Below 50 %", table = TABLES, type = "col_next", tb.pos=2, n=3)))
                                word_main_values_df$SaO2_Below50_TST_perc_treatment=as.numeric(gsub("%","",find_field_grep(pattern = "Below 50 %",
                                                                                                                          pattern2 = "Total Sleep Time", table = TABLES, type = "row_below", tb.pos=2, n=10)))
                                
                                
                                #Fix as numeric
                                word_main_values_df[,c(9,11,13,21:35,38:168,171:301)] <- sapply(word_main_values_df[,c(9,11,13,21:35,38:168,171:301)], as.numeric)
                                
                                #Fix columns names
                                colnames(word_main_values_df)[c(1:3,5:8)] <- paste0("S02_PSG_PHI_",colnames(word_main_values_df)[c(1:3,5:8)])
                                colnames(word_main_values_df)[-c(1:3,5:8)] <- paste0("S02_PSG_",colnames(word_main_values_df)[-c(1:3,5:8)])
                                
                                # Get other relevant values from the word document
                                word_main_values_df$S02_PSG_study_type <- "Split Night Study"
                                word_main_values_df$S02_PSG_PHI_clinical_history <- grep("CLINICAL HISTORY:", textreadr::read_docx(PSGfilename), value = T)
                                word_main_values_df$S02_PSG_PHI_final_diagnosis <- grep("FINAL DIAGNOSIS:", textreadr::read_docx(PSGfilename), value = T)
                                word_main_values_df$S02_PSG_PHI_comm_rec <- grep("RECOMMENDATIONS:", textreadr::read_docx(PSGfilename), value = T)
                                
                                
                                
                                
                                
                                
                                ### Get fields from Excel embedded tables that appear in the report
                                xl_tb <- PSGfilename
                                
                                #Copy file ################## Fix ths to be universal
                                dir.create("zips_s02/", showWarnings = F)
                                file.copy(from = PSGfilename, to = "zips_s02/")
                                file.rename(paste0("zips_s02/", strsplit(PSGfilename, "/")[[1]][6]), "zips_s02/current.zip")
                                unzip("zips_s02/current.zip", exdir = paste0(getwd(),"/zips_s02"))
                                xls_tables_paths <- grep("Worksheet", dir("zips_s02/word/embeddings/", full.names = T), value = T)
                                
                                
                                # Iterate to load all possible tables - not applicable to Split 02
                                xtblist <- NULL
                                
                                for (xt in xls_tables_paths) {
                                        xtblist[[xt]]  <- read.xlsx(xt, sheetIndex = 1)
                                }
                                
                                
                              
                               
                                
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
                                
                                xtblist2 <- lapply(xtblist2, function(x) {
                                        x$value <- as.character(x$value)
                                        return(x) } )
                                
                                xlstb2_combined <- bind_rows(xtblist2)
                                
                                #Remove repeated rows
                                xlstb2_combined <- unique(xlstb2_combined)
                                xlstb_combined_extraonly <- xlstb2_combined %>% filter(!is.na(label), !grepl("^<<", xlstb2_combined$Var_name))
                                
                                
                                
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
                                colnames(extrapsgdata) <- paste0("S02_PSG_EXTRA_",t(xlstb_combined_extraonly$Var_name))
                                
                                ######### Remove all files from zip_folder to avoid confusion when getting new data
                                file.remove(list.files("zips_s02/", include.dirs = F, full.names = T, recursive = T))
                                
                                ######## Combine ALL data
                                
                                all_fields <- c("S02_PSG_PHI_Filename", "PennSleepID", colnames(word_main_values_df), colnames(extrapsgdata))
                                
                                #Initiate dataframe with results
                                
                                
                                
                                #If first, create final_df
                                if (first) {
                                        final_df <- data.frame(matrix(NA,1,length(all_fields)))
                                        colnames(final_df) <- all_fields
                                        
                                        final_df[1,] <- c(PSGfilename,
                                                          NA,
                                                          word_main_values_df[1,],
                                                          extrapsgdata[1,])
                                        first <- F
                                } else {
                                        
                                        next_final_df <- data.frame(matrix(NA,1,length(all_fields)))
                                        colnames(next_final_df) <- all_fields
                                        
                                        next_final_df[1,] <- c(PSGfilename,
                                                               NA,
                                                               word_main_values_df[1,],
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
        final_df$S02_PSG_Age_at_Study <- interval(mdy(final_df$S02_PSG_PHI_NPBMPatientInfoDOB), mdy(final_df$S02_PSG_NPBMStudyInfoStudyDate)) %/% years(1)
        
        #Create identifiable data frame
        identifiable_df <- final_df
        identifiable_df$ProcessedDate <- Sys.Date()
        
        ProcessedTimeID=gsub("-","",gsub(":","",gsub(" ", "", Sys.time())))
        
        identifiable_df$ProcessedTimeID <- ProcessedTimeID
        
        # Per sample QC (not value filter)
        
        QC_df <- data.frame(S02_PSG_PHI_Filename=identifiable_df$S02_PSG_PHI_Filename,
                            # Missing both MRN fields
                            has_missingMRN=is.na(identifiable_df$S02_PSG_PHI_NPBMSiteInfoHospitalNumber),
                            # Gender not (Male, Female, Unknown)
                            has_missingGender=!(tolower(identifiable_df$S02_PSG_NPBMPatientInfoGender) %in% c("male", "female", "unknown")),
                            # Missing age
                            has_missingAge=is.na(identifiable_df$S02_PSG_Age_at_Study),
                            # Missing BMI
                            has_missingBMI=is.na(identifiable_df$S02_PSG_NPBMPatientInfoBMI)
        )
        
        identifiable_df$PerSample_QC <- NA
        identifiable_df$PerSample_QC[rowSums(QC_df[,2:5])>0] <- "FAIL"
        identifiable_df$PerSample_QC[rowSums(QC_df[,2:5])==0] <- "PASS"
        
        QC_df$PerSample_QC <- identifiable_df$PerSample_QC
        write.csv(QC_df, paste0("QC_dataframe_", ProcessedTimeID, ".csv"))
        
        #Save identifiable Rdata file encrypted
        #key <- key_sodium(sodium::keygen())
        saveRDS(identifiable_df, paste0("PennSleepDatabase_Split02_", ProcessedTimeID, ".Rdata"))
        #cyphr::encrypt_file("PennSleepDatabase.Rdata", key, "PennSleepDatabase.encrypted")
        
        #Save de-identified version in CSV files - implement a way to check in database if sample was processed and not use the same de-identified IDs
        #identifiable_df$PennSleepID <- paste0("PENNSLEEP00000",seq(1:nrow(identifiable_df)))
        deidentified_df <- select(identifiable_df, -starts_with("PSG_PHI_"))
        
        write.csv(deidentified_df, paste0("PennSleepDatabase_Split02_Deidentified",ProcessedTimeID,".csv"), row.names = F)
        
        return(identifiable_df)
        
}

