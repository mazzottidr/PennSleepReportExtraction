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

extract_baseline_2 <- function(paths) {
        
        #paths <- filter(ALL_annotated, predictedTableFormat=="13", predictedStudyType=="Baseline")$linux_path[1:10]
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
                            #dir.create("temp_reports_b02/", showWarnings = FALSE)
                            #file.copy(PSGfilename, to = "temp_reports_b02/")
                            
                            #docx <- paste0("temp_reports_b02/", sapply(strsplit(PSGfilename, "/"), "[[", 6))
                            
                            dir.create("temp_reports_b13/", showWarnings = FALSE)
                            file.copy(PSGfilename, to = "temp_reports_b13/")
                            
                            docx <- paste0("temp_reports_b13/", sapply(strsplit(PSGfilename, "/"), "[[", 6))
                            
                            
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
                                
                                # PSG variables
                                word_main_values_df$VAR_StartTime <- find_field_grep("Start Time :", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$VAR_EndTime <- find_field_grep("End Time :", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$VAR_TotalRecordingTime <- find_field_grep("Total Recording Time (minutes) :", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$VAR_TotalSleepTime <- find_field_grep("Total Sleep Time (minutes) :", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$VAR_SleepOnsetLatency <- find_field_grep("Sleep Onset Latency (minutes) :", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$VAR_NumberREMPeriods <- find_field_grep("Number of REM Periods :", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$VAR_REMLatency <- find_field_grep("REM Latency :", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$VAR_WASO <- find_field_grep("WASO:", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$VAR_StageN1 <- find_field_grep("Stage N1:", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$VAR_StageN2 <- find_field_grep("Stage N2:", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$VAR_StageN3 <- find_field_grep("Stage N3", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$VAR_StageREM <- find_field_grep("REM:", TABLES, type = "col_next", tb.pos=1, fixed = T)
                                word_main_values_df$VAR_StageN1_pct <- trim(gsub("%", "", find_field_grep("Stage N1:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=2)))
                                word_main_values_df$VAR_StageN2_pct <- trim(gsub("%", "", find_field_grep("Stage N2:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=2)))
                                word_main_values_df$VAR_StageN3_pct <- trim(gsub("%", "", find_field_grep("Stage N3", TABLES, type = "col_next", tb.pos=1, fixed = T, n=2)))
                                word_main_values_df$VAR_StageREM_pct <- trim(gsub("%", "", find_field_grep("REM:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=2)))
                                
                                word_main_values_df$VAR_TimeSupine <- period_to_seconds(hms(find_field_grep("Supine:", TABLES, type = "col_next", tb.pos=1, fixed = T)))/60
                                word_main_values_df$VAR_TimeSupine_pct <-  trim(gsub("%", "", find_field_grep("Supine:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=2)))
                                word_main_values_df$VAR_TimeLeft <- period_to_seconds(hms(find_field_grep("Left:", TABLES, type = "col_next", tb.pos=1, fixed = T)))/60
                                word_main_values_df$VAR_TimeLeft_pct <-  trim(gsub("%", "", find_field_grep("Left:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=2)))
                                word_main_values_df$VAR_TimeRight <- period_to_seconds(hms(find_field_grep("Right:", TABLES, type = "col_next", tb.pos=1, fixed = T)))/60
                                word_main_values_df$VAR_TimeRight_pct <-  trim(gsub("%", "", find_field_grep("Right:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=2)))
                                word_main_values_df$VAR_TimeProne <- period_to_seconds(hms(find_field_grep("Prone:", TABLES, type = "col_next", tb.pos=1, fixed = T)))/60
                                word_main_values_df$VAR_TimeProne_pct <-  trim(gsub("%", "", find_field_grep("Prone:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=2)))
                              
                                word_main_values_df$VAR_N_Central_Events <- find_field_grep(pattern = "CENTRAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=1)
                                word_main_values_df$VAR_Idx_Central_Events  <- find_field_grep(pattern = "CENTRAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=2)
                                word_main_values_df$VAR_MeanDur_Central_Events  <- find_field_grep(pattern = "CENTRAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=3)
                                word_main_values_df$VAR_LongestDur_Central_Events  <- find_field_grep(pattern = "CENTRAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=4)
                                word_main_values_df$VAR_N_Central_Events_REM  <- find_field_grep(pattern = "CENTRAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=5)
                                word_main_values_df$VAR_N_Central_Events_NREM  <- find_field_grep(pattern = "CENTRAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=6)
                                word_main_values_df$VAR_Idx_Central_Events_REM  <- find_field_grep(pattern = "CENTRAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=7)
                                word_main_values_df$VAR_Idx_Central_Events_NREM  <- find_field_grep(pattern = "CENTRAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=8)
                                
                                word_main_values_df$VAR_N_Obstructive_Events <- find_field_grep(pattern = "OBSTRUCTIVE", TABLES, type = "row_below", tb.pos=1, fixed = T, n=1)
                                word_main_values_df$VAR_Idx_Obstructive_Events  <- find_field_grep(pattern = "OBSTRUCTIVE", TABLES, type = "row_below", tb.pos=1, fixed = T, n=2)
                                word_main_values_df$VAR_MeanDur_Obstructive_Events  <- find_field_grep(pattern = "OBSTRUCTIVE", TABLES, type = "row_below", tb.pos=1, fixed = T, n=3)
                                word_main_values_df$VAR_LongestDur_Obstructive_Events  <- find_field_grep(pattern = "OBSTRUCTIVE", TABLES, type = "row_below", tb.pos=1, fixed = T, n=4)
                                word_main_values_df$VAR_N_Obstructive_Events_REM  <- find_field_grep(pattern = "OBSTRUCTIVE", TABLES, type = "row_below", tb.pos=1, fixed = T, n=5)
                                word_main_values_df$VAR_N_Obstructive_Events_NREM  <- find_field_grep(pattern = "OBSTRUCTIVE", TABLES, type = "row_below", tb.pos=1, fixed = T, n=6)
                                word_main_values_df$VAR_Idx_Obstructive_Events_REM  <- find_field_grep(pattern = "OBSTRUCTIVE", TABLES, type = "row_below", tb.pos=1, fixed = T, n=7)
                                word_main_values_df$VAR_Idx_Obstructive_Events_NREM  <- find_field_grep(pattern = "OBSTRUCTIVE", TABLES, type = "row_below", tb.pos=1, fixed = T, n=8)
                                
                                word_main_values_df$VAR_N_Mixed_Events <- find_field_grep(pattern = "MIXED", TABLES, type = "row_below", tb.pos=1, fixed = T, n=1)
                                word_main_values_df$VAR_Idx_Mixed_Events  <- find_field_grep(pattern = "MIXED", TABLES, type = "row_below", tb.pos=1, fixed = T, n=2)
                                word_main_values_df$VAR_MeanDur_Mixed_Events  <- find_field_grep(pattern = "MIXED", TABLES, type = "row_below", tb.pos=1, fixed = T, n=3)
                                word_main_values_df$VAR_LongestDur_Mixed_Events  <- find_field_grep(pattern = "MIXED", TABLES, type = "row_below", tb.pos=1, fixed = T, n=4)
                                word_main_values_df$VAR_N_Mixed_Events_REM  <- find_field_grep(pattern = "MIXED", TABLES, type = "row_below", tb.pos=1, fixed = T, n=5)
                                word_main_values_df$VAR_N_Mixed_Events_NREM  <- find_field_grep(pattern = "MIXED", TABLES, type = "row_below", tb.pos=1, fixed = T, n=6)
                                word_main_values_df$VAR_Idx_Mixed_Events_REM  <- find_field_grep(pattern = "MIXED", TABLES, type = "row_below", tb.pos=1, fixed = T, n=7)
                                word_main_values_df$VAR_Idx_Mixed_Events_NREM  <- find_field_grep(pattern = "MIXED", TABLES, type = "row_below", tb.pos=1, fixed = T, n=8)
                                
                                word_main_values_df$VAR_N_Apnea_Events <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=1)
                                word_main_values_df$VAR_Idx_Apnea_Events  <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=2)
                                word_main_values_df$VAR_MeanDur_Apnea_Events  <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=3)
                                word_main_values_df$VAR_LongestDur_Apnea_Events  <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=4)
                                word_main_values_df$VAR_N_Apnea_Events_REM  <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=5)
                                word_main_values_df$VAR_N_Apnea_Events_NREM  <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=6)
                                word_main_values_df$VAR_Idx_Apnea_Events_REM  <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=7)
                                word_main_values_df$VAR_Idx_Apnea_Events_NREM  <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=1, fixed = T, n=8)
                                
                                word_main_values_df$VAR_N_Hypopneas_Events <- find_field_grep(pattern = "HYPOPNEAS", TABLES, type = "row_below", tb.pos=1, fixed = T, n=1)
                                word_main_values_df$VAR_Idx_Hypopneas_Events  <- find_field_grep(pattern = "HYPOPNEAS", TABLES, type = "row_below", tb.pos=1, fixed = T, n=2)
                                word_main_values_df$VAR_MeanDur_Hypopneas_Events  <- find_field_grep(pattern = "HYPOPNEAS", TABLES, type = "row_below", tb.pos=1, fixed = T, n=3)
                                word_main_values_df$VAR_LongestDur_Hypopneas_Events  <- find_field_grep(pattern = "HYPOPNEAS", TABLES, type = "row_below", tb.pos=1, fixed = T, n=4)
                                word_main_values_df$VAR_N_Hypopneas_Events_REM  <- find_field_grep(pattern = "HYPOPNEAS", TABLES, type = "row_below", tb.pos=1, fixed = T, n=5)
                                word_main_values_df$VAR_N_Hypopneas_Events_NREM  <- find_field_grep(pattern = "HYPOPNEAS", TABLES, type = "row_below", tb.pos=1, fixed = T, n=6)
                                word_main_values_df$VAR_Idx_Hypopneas_Events_REM  <- find_field_grep(pattern = "HYPOPNEAS", TABLES, type = "row_below", tb.pos=1, fixed = T, n=7)
                                word_main_values_df$VAR_Idx_Hypopneas_Events_NREM  <- find_field_grep(pattern = "HYPOPNEAS", TABLES, type = "row_below", tb.pos=1, fixed = T, n=8)
                                
                                word_main_values_df$VAR_Total_AHI <- find_field_grep(pattern = "INDEX", TABLES, type = "row_below", tb.pos=1, fixed = T, n=1)
                                word_main_values_df$VAR_Total_AH_count <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=1)
                                word_main_values_df$VAR_NREM_AHI <- find_field_grep(pattern = "INDEX", TABLES, type = "row_below", tb.pos=1, fixed = T, n=2)
                                word_main_values_df$VAR_NREM_AH_count <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=2)
                                word_main_values_df$VAR_REM_AHI <- find_field_grep(pattern = "INDEX", TABLES, type = "row_below", tb.pos=1, fixed = T, n=3)
                                word_main_values_df$VAR_REM_AH_count <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=3)
                                word_main_values_df$VAR_Supine_AHI <- find_field_grep(pattern = "INDEX", TABLES, type = "row_below", tb.pos=1, fixed = T, n=4)
                                word_main_values_df$VAR_Supine_AH_count <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=4)
                                word_main_values_df$VAR_Lateral_AHI <- find_field_grep(pattern = "INDEX", TABLES, type = "row_below", tb.pos=1, fixed = T, n=5)
                                word_main_values_df$VAR_Lateral_AH_count <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=5)
                                word_main_values_df$VAR_Prone_AHI <- find_field_grep(pattern = "INDEX", TABLES, type = "row_below", tb.pos=1, fixed = T, n=6)
                                word_main_values_df$VAR_Prone_AH_count <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=6)
                                word_main_values_df$VAR_Left_AHI <- find_field_grep(pattern = "INDEX", TABLES, type = "row_below", tb.pos=1, fixed = T, n=7)
                                word_main_values_df$VAR_Left_AH_count <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=7)
                                word_main_values_df$VAR_Right_AHI <- find_field_grep(pattern = "INDEX", TABLES, type = "row_below", tb.pos=1, fixed = T, n=8)
                                word_main_values_df$VAR_Right_AH_count <- find_field_grep(pattern = "TOTAL", TABLES, type = "row_below", tb.pos=2, fixed = T, n=8)
                                
                                word_main_values_df$VAR_PLM_Total_Idx <- find_field_grep(pattern = "Total Events:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=1)
                                word_main_values_df$VAR_PLM_Total_Idx_NREM <- find_field_grep(pattern = "Non-REM Events:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=1)
                                word_main_values_df$VAR_PLM_Total_Idx_REM <- strsplit(pull(TABLES[[grep("REM Events:", TABLES)[1]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),2]), " ")[[1]][length(strsplit(pull(TABLES[[grep("REM Events:", TABLES)[1]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),2]), " ")[[1]])]
                                
                                word_main_values_df$VAR_PLM_Total_counts <- find_field_grep(pattern = "Total Events:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=2)
                                word_main_values_df$VAR_PLM_Total_counts_NREM <- find_field_grep(pattern = "Non-REM Events:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=2)
                                word_main_values_df$VAR_PLM_Total_counts_REM <- strsplit(pull(TABLES[[grep("REM Events:", TABLES)[1]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),3]), " ")[[1]][length(strsplit(pull(TABLES[[grep("REM Events:", TABLES)[1]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),3]), " ")[[1]])]
                                
                                word_main_values_df$VAR_PLM_withArousal_Idx <- find_field_grep(pattern = "Total Events:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=3)
                                word_main_values_df$VAR_PLM_withArousal_Idx_NREM <- find_field_grep(pattern = "Non-REM Events:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=3)
                                word_main_values_df$VAR_PLM_withArousal_Idx_REM <-strsplit(pull(TABLES[[grep("REM Events:", TABLES)[1]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),4]), " ")[[1]][length(strsplit(pull(TABLES[[grep("REM Events:", TABLES)[1]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),4]), " ")[[1]])]
                                
                                word_main_values_df$VAR_PLM_withArousal_counts <- find_field_grep(pattern = "Total Events:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=4)
                                word_main_values_df$VAR_PLM_withArousal_counts_NREM <- find_field_grep(pattern = "Non-REM Events:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=4)
                                word_main_values_df$VAR_PLM_withArousal_counts_REM <- strsplit(pull(TABLES[[grep("REM Events:", TABLES)[1]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),5]), " ")[[1]][length(strsplit(pull(TABLES[[grep("REM Events:", TABLES)[1]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),5]), " ")[[1]])]
                                
                                word_main_values_df$VAR_PLM_withoutArousal_Idx <- find_field_grep(pattern = "Total Events:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=5)
                                word_main_values_df$VAR_PLM_withoutArousal_Idx_NREM <- find_field_grep(pattern = "Non-REM Events:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=5)
                                word_main_values_df$VAR_PLM_withoutArousal_Idx_REM <- strsplit(pull(TABLES[[grep("REM Events:", TABLES)[1]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),6]), " ")[[1]][length(strsplit(pull(TABLES[[grep("REM Events:", TABLES)[1]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),6]), " ")[[1]])]
                                
                                word_main_values_df$VAR_PLM_withoutArousal_counts <- find_field_grep(pattern = "Total Events:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=6)
                                word_main_values_df$VAR_PLM_withoutArousal_counts_NREM <- find_field_grep(pattern = "Non-REM Events:", TABLES, type = "col_next", tb.pos=1, fixed = T, n=6)
                                word_main_values_df$VAR_PLM_withoutArousal_counts_REM <- strsplit(pull(TABLES[[grep("REM Events:", TABLES)[1]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),7]), " ")[[1]][length(strsplit(pull(TABLES[[grep("REM Events:", TABLES)[1]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[1]]]$V1),7]), " ")[[1]])]
                                
                                word_main_values_df$VAR_Arousal_Idx <- find_field_grep(pattern = "Total Events:", TABLES, type = "col_next", tb.pos=2, fixed = T, n=1)
                                word_main_values_df$VAR_Arousal_Idx_NREM <- find_field_grep(pattern = "Non-REM Events:", TABLES, type = "col_next", tb.pos=2, fixed = T, n=1)
                                word_main_values_df$VAR_Arousal_Idx_REM <- strsplit(pull(TABLES[[grep("REM Events:", TABLES)[2]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[2]]]$V1),2]), " ")[[1]][length(strsplit(pull(TABLES[[grep("REM Events:", TABLES)[2]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[2]]]$V1),2]), " ")[[1]])]
                                
                                word_main_values_df$VAR_Arousal_counts <- find_field_grep(pattern = "Total Events:", TABLES, type = "col_next", tb.pos=2, fixed = T, n=2)
                                word_main_values_df$VAR_Arousal_counts_NREM <- find_field_grep(pattern = "Non-REM Events:", TABLES, type = "col_next", tb.pos=2, fixed = T, n=2)
                                word_main_values_df$VAR_Arousal_counts_REM <- strsplit(pull(TABLES[[grep("REM Events:", TABLES)[2]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[2]]]$V1),3]), " ")[[1]][length(strsplit(pull(TABLES[[grep("REM Events:", TABLES)[2]]][grep("^REM Events:", TABLES[[grep("REM Events:", TABLES)[2]]]$V1),3]), " ")[[1]])]
                                
                                #Fix as numeric
                                word_main_values_df[,c(9,11,13,21:122)] <- sapply(word_main_values_df[,c(9,11,13,21:122)], as.numeric)
                                
                                
                                #Fix columns names
                                colnames(word_main_values_df)[c(1:3,5:8)] <- paste0("B02_PSG_PHI_",colnames(word_main_values_df)[c(1:3,5:8)])
                                colnames(word_main_values_df)[-c(1:3,5:8)] <- paste0("B02_PSG_",colnames(word_main_values_df)[-c(1:3,5:8)])
                                
                                # Get other relevant values from the word document
                                
                                word_main_values_df$B02_PSG_study_type <- "Baseline Study"
                                word_main_values_df$B02_PSG_PHI_clinical_history <- grep("CLINICAL HISTORY:", textreadr::read_docx(PSGfilename), value = T)
                                word_main_values_df$B02_PSG_PHI_final_diagnosis <- get_final_diagnosis(TABLES)
                                word_main_values_df$B02_PSG_PHI_comm_rec <- get_comm_rec(TABLES)
                                
                                ### Get fields from Excel embedded tables that appear in the report
                                xl_tb <- PSGfilename
                                
                                #Copy file ################## Fix ths to be universal
                                dir.create("zips_b02/", showWarnings = F)
                                file.copy(from = PSGfilename, to = "zips_b02/")
                                file.rename(paste0("zips_b02/", strsplit(PSGfilename, "/")[[1]][6]), "zips_b02/current.zip")
                                unzip("zips_b02/current.zip", exdir = paste0(getwd(),"/zips_b02"))
                                xls_tables_paths <- grep("Worksheet", dir("zips_b02/word/embeddings/", full.names = T), value = T)
                                
                                
                                # Iterate to load all possible tables
                                xtblist <- NULL
                                
                                for (xt in xls_tables_paths) {
                                        xtblist[[xt]]  <- read.xlsx(xt, sheetIndex = 1)
                                }
                                
                                
                                # Try getting position time in same spreadsheet
                                
                                # if (length(which(grepl("Position", xtblist)))==1) {
                                #         
                                #         xlstb1 <- xtblist[[which(grepl("Position", xtblist))]][,c(8,10,11)] #Position time
                                #         
                                # } else {
                                #         
                                #         # Otherwise, in separated
                                #         xlstb1 <- xtblist[[which(grepl("Position", xtblist) &  grepl("Time..min.", xtblist))]] #Position time
                                #         
                                # }
                                
                                
                                xlstb3 <- xtblist[[which(grepl("OXYGEN.SATURATION", xtblist))[1]]] #Oxygen saturation
                                xlstb4 <- xtblist[[which(grepl("PULSE.RATE.RESULTS", xtblist))[1]]] #Pulse rate
                                xlstb5 <- xtblist[[which(grepl("OXYGEN.DESATURATION.EVENTS", xtblist))[1]]] #Oxygen desats
                                
                                
                                if (!is.null(xlstb3)) {
                                        
                                        xls3_vars <- data.frame(
                                                
                                                #Oxygen saturation
                                                B02_PSG_maxSpO2_wake=as.numeric(as.character(xlstb3$Wake[1])),
                                                B02_PSG_meanSpO2_wake=as.numeric(as.character(xlstb3$Wake[2])),
                                                B02_PSG_minSpO2_wake=as.numeric(as.character(xlstb3$Wake[3])),
                                                B02_PSG_timeSpO2_LT89_wake=as.numeric(as.character(xlstb3$Wake[4])),
                                                B02_PSG_perctime_90.100_wake=as.numeric(as.character(xlstb3$Wake[6])),
                                                B02_PSG_perctime_80.89_wake=as.numeric(as.character(xlstb3$Wake[7])),
                                                B02_PSG_perctime_70.79_wake=as.numeric(as.character(xlstb3$Wake[8])),
                                                B02_PSG_perctime_badO2data_wake=as.numeric(as.character(xlstb3$Wake[12])),
                                                
                                                B02_PSG_maxSpO2_nrem=as.numeric(as.character(xlstb3$Non.REM[1])),
                                                B02_PSG_meanSpO2_nrem=as.numeric(as.character(xlstb3$Non.REM[2])),
                                                B02_PSG_minSpO2_nrem=as.numeric(as.character(xlstb3$Non.REM[3])),
                                                B02_PSG_timeSpO2_LT89_nrem=as.numeric(as.character(xlstb3$Non.REM[4])),
                                                B02_PSG_perctime_90.100_nrem=as.numeric(as.character(xlstb3$Non.REM[6])),
                                                B02_PSG_perctime_80.89_nrem=as.numeric(as.character(xlstb3$Non.REM[7])),
                                                B02_PSG_perctime_70.79_nrem=as.numeric(as.character(xlstb3$Non.REM[8])),
                                                B02_PSG_perctime_badO2data_nrem=as.numeric(as.character(xlstb3$Non.REM[12])),
                                                
                                                B02_PSG_maxSpO2_rem=as.numeric(as.character(xlstb3$REM[1])),
                                                B02_PSG_meanSpO2_rem=as.numeric(as.character(xlstb3$REM[2])),
                                                B02_PSG_minSpO2_rem=as.numeric(as.character(xlstb3$REM[3])),
                                                B02_PSG_timeSpO2_LT89_rem=as.numeric(as.character(xlstb3$REM[4])),
                                                B02_PSG_perctime_90.100_rem=as.numeric(as.character(xlstb3$REM[6])),
                                                B02_PSG_perctime_80.89_rem=as.numeric(as.character(xlstb3$REM[7])),
                                                B02_PSG_perctime_70.79_rem=as.numeric(as.character(xlstb3$REM[8])),
                                                B02_PSG_perctime_badO2data_rem=as.numeric(as.character(xlstb3$REM[12])),
                                                
                                                B02_PSG_maxSpO2_tst=as.numeric(as.character(xlstb3$TST[1])),
                                                B02_PSG_meanSpO2_tst=as.numeric(as.character(xlstb3$TST[2])),
                                                B02_PSG_minSpO2_tst=as.numeric(as.character(xlstb3$TST[3])),
                                                B02_PSG_timeSpO2_LT89_tst=as.numeric(as.character(xlstb3$TST[4])),
                                                B02_PSG_perctime_90.100_tst=as.numeric(as.character(xlstb3$TST[6])),
                                                B02_PSG_perctime_80.89_tst=as.numeric(as.character(xlstb3$TST[7])),
                                                B02_PSG_perctime_70.79_tst=as.numeric(as.character(xlstb3$TST[8])),
                                                B02_PSG_perctime_badO2data_tst=as.numeric(as.character(xlstb3$TST[12])),
                                                
                                                B02_PSG_maxSpO2_tib=as.numeric(as.character(xlstb3$TIB[1])),
                                                B02_PSG_meanSpO2_tib=as.numeric(as.character(xlstb3$TIB[2])),
                                                B02_PSG_minSpO2_tib=as.numeric(as.character(xlstb3$TIB[3])),
                                                B02_PSG_timeSpO2_LT89_tib=as.numeric(as.character(xlstb3$TIB[4])),
                                                B02_PSG_perctime_90.100_tib=as.numeric(as.character(xlstb3$TIB[6])),
                                                B02_PSG_perctime_80.89_tib=as.numeric(as.character(xlstb3$TIB[7])),
                                                B02_PSG_perctime_70.79_tib=as.numeric(as.character(xlstb3$TIB[8])),
                                                B02_PSG_perctime_badO2data_tib=as.numeric(as.character(xlstb3$TIB[12])),stringsAsFactors = F)
                                } else {xls3_vars <- NULL}
                                
                               
                                if (!is.null(xlstb4)) {
                                        
                                        xls4_vars <- data.frame(       
                                                #Pulse rate
                                                B02_PSG_maxHR_wake=as.numeric(as.character(xlstb4$Wake[1])),
                                                B02_PSG_meanHR_wake=as.numeric(as.character(xlstb4$Wake[2])),
                                                B02_PSG_minHR_wake=as.numeric(as.character(xlstb4$Wake[3])),
                                                
                                                B02_PSG_maxHR_nrem=as.numeric(as.character(xlstb4$Non.REM[1])),
                                                B02_PSG_meanHR_nrem=as.numeric(as.character(xlstb4$Non.REM[2])),
                                                B02_PSG_minHR_nrem=as.numeric(as.character(xlstb4$Non.REM[3])),
                                                
                                                B02_PSG_maxHR_rem=as.numeric(as.character(xlstb4$REM[1])),
                                                B02_PSG_meanHR_rem=as.numeric(as.character(xlstb4$REM[2])),
                                                B02_PSG_minHR_rem=as.numeric(as.character(xlstb4$REM[3])),
                                                
                                                B02_PSG_maxHR_tst=as.numeric(as.character(xlstb4$TST[1])),
                                                B02_PSG_meanHR_tst=as.numeric(as.character(xlstb4$TST[2])),
                                                B02_PSG_minHR_tst=as.numeric(as.character(xlstb4$TST[3])), stringsAsFactors = F)
                                } else {xls4_vars <- NULL}
                                
                               
                                if (!is.null(xlstb5)) {
                                        
                                        xls5_vars <- data.frame(    
                                                #Oxygen desats
                                                B02_PSG_desats_tst_idx=as.numeric(xlstb5$Index[1]),
                                                B02_PSG_desats_tst_count=as.numeric(xlstb5$Count[1]),
                                                
                                                B02_PSG_desats_nrem_idx=as.numeric(xlstb5$Index[2]),
                                                B02_PSG_desats_nrem_count=as.numeric(xlstb5$Count[2]),
                                                
                                                B02_PSG_desats_rem_idx=as.numeric(xlstb5$Index[3]),
                                                B02_PSG_desats_rem_count=as.numeric(xlstb5$Count[3]),stringsAsFactors = F)
                                } else {xls5_vars <- NULL}
                                
                                
                                
                                
                                
                                
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
                                colnames(extrapsgdata) <- paste0("B02_PSG_EXTRA_",t(xlstb_combined_extraonly$Var_name))
                                
                                
                                ######### Remove all files from zip_folder to avoid confusion when getting new data
                                file.remove(list.files("zips_b02/", include.dirs = F, full.names = T, recursive = T))
                                
                                
                                
                                ######## Combine ALL data
                                
                                all_fields <- c("B02_PSG_PHI_Filename", "PennSleepID", colnames(word_main_values_df), colnames(xls3_vars), colnames(xls4_vars), colnames(xls5_vars), colnames(extrapsgdata))
                                
                                #Initiate dataframe with results
                                
                                
                                
                                #If first, create final_df
                                if (first) {
                                        final_df <- data.frame(matrix(NA,1,length(all_fields)))
                                        colnames(final_df) <- all_fields
                                        
                                        final_df[1,] <- c(PSGfilename,
                                                          NA,
                                                          word_main_values_df[1,],
                                                          xls3_vars[1,],
                                                          xls4_vars[1,],
                                                          xls5_vars[1,],
                                                          extrapsgdata[1,])
                                        first <- F
                                } else {
                                        
                                        next_final_df <- data.frame(matrix(NA,1,length(all_fields)))
                                        colnames(next_final_df) <- all_fields
                                        
                                        next_final_df[1,] <- c(PSGfilename,
                                                               NA,
                                                               word_main_values_df[1,],
                                                               xls3_vars[1,],
                                                               xls4_vars[1,],
                                                               xls5_vars[1,],
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
        final_df$B02_PSG_Age_at_Study <- interval(mdy(final_df$B02_PSG_PHI_NPBMPatientInfoDOB), mdy(final_df$B02_PSG_NPBMStudyInfoStudyDate)) %/% years(1)
        
        #Create identifiable data frame
        identifiable_df <- final_df
        identifiable_df$ProcessedDate <- Sys.Date()
        
        ProcessedTimeID=gsub("-","",gsub(":","",gsub(" ", "", Sys.time())))
        
        identifiable_df$ProcessedTimeID <- ProcessedTimeID
        
        # Per sample QC (not value filter)
        
        QC_df <- data.frame(B02_PSG_PHI_Filename=identifiable_df$B02_PSG_PHI_Filename,
                            # Missing both MRN fields
                            has_missingMRN=is.na(identifiable_df$B02_PSG_PHI_NPBMSiteInfoHospitalNumber),
                            # Gender not (Male, Female, Unknown)
                            has_missingGender=!(tolower(identifiable_df$B02_PSG_NPBMPatientInfoGender) %in% c("male", "female", "unknown")),
                            # Missing age
                            has_missingAge=is.na(identifiable_df$B02_PSG_Age_at_Study),
                            # Missing BMI
                            has_missingBMI=is.na(identifiable_df$B02_PSG_NPBMPatientInfoBMI)
        )
        
        identifiable_df$PerSample_QC <- NA
        identifiable_df$PerSample_QC[rowSums(QC_df[,2:5])>0] <- "FAIL"
        identifiable_df$PerSample_QC[rowSums(QC_df[,2:5])==0] <- "PASS"
        
        QC_df$PerSample_QC <- identifiable_df$PerSample_QC
        write.csv(QC_df, paste0("QC_dataframe_", ProcessedTimeID, ".csv"))
        
        #Save identifiable Rdata file encrypted
        #key <- key_sodium(sodium::keygen())
        saveRDS(identifiable_df, paste0("PennSleepDatabase_Baseline02_", ProcessedTimeID, ".Rdata"))
        #cyphr::encrypt_file("PennSleepDatabase.Rdata", key, "PennSleepDatabase.encrypted")
        
        #Save de-identified version in CSV files - implement a way to check in database if sample was processed and not use the same de-identified IDs
        #identifiable_df$PennSleepID <- paste0("PENNSLEEP00000",seq(1:nrow(identifiable_df)))
        deidentified_df <- select(identifiable_df, -starts_with("PSG_PHI_"))
        
        write.csv(deidentified_df, paste0("PennSleepDatabase_Baseline02_Deidentified",ProcessedTimeID,".csv"), row.names = F)
        
        return(identifiable_df)
        
}

