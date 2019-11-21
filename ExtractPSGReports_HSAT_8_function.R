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

get_values <- function(string, source, strstrip1="CHARFORMAT ", strstrip2=" ", indx=2) {
        
        tryCatch({
                #string <- "ESStatInfo_Resp_MeanDuration_10"
                #source <- textreadr::read_docx(PSGfilename)
                
                element <- grep(paste0("\\b", string, "\\b"), source, value = T)
                
                if (length(element)==0) {
                        return(NA)
                }
                
                stripped <- unlist(strsplit(unlist(strsplit(element, strstrip1)), strstrip2))
                
                final <- stripped[grep(paste0("\\b", string, "\\b"), stripped)+2]
                
                return(final)
                
        }, error=function(e) NA)
      
}

extract_hsat_8 <- function(paths) {
        
        #paths <- h08_paths[1:10]
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
                            dir.create("temp_reports_h08/", showWarnings = FALSE)
                            file.copy(PSGfilename, to = "temp_reports_h08/")
                            
                            #docx <- paste0("temp_reports_h08/", sapply(strsplit(PSGfilename, "/"), "[[", 3))
                            docx <- paste0("temp_reports_h08/", sapply(strsplit(PSGfilename, "/"), "[[", 6))
                            
                            
                            
                            #current_tmp_files <- list.files(tempdir(), full.names = T)
                            
                            message(paste0("Processing: ", sapply(strsplit(PSGfilename, "/"), "[[", 6)))
                            message(paste0(format(100*i/length(PSGfiles), digits = 4),"% done,"))
                            
                            #Get all possible tables
                                #TABLES <-get_tbls(PSGfilename) #This gives a warning
                                TABLES <-sapply(1:docx_tbl_count(docxtractr::read_docx(docx)), docx_extract_tbl, docx=docxtractr::read_docx(docx), header=F)
                                
                                
                                # Include in the df after fixing
                                PatientFullName <- sapply(strsplit(find_field_grep("Patient Name:", TABLES, type = "col_next", tb.pos=1, fixed = T), "CHARFORMAT "), "[[", 2)
                                PatientHeight <- sapply(strsplit(find_field_grep("Height:", TABLES, type = "col_next", tb.pos=1, fixed = T), "CHARFORMAT "), "[[", length(strsplit(find_field_grep("Height:", TABLES, type = "col_next", tb.pos=1, fixed = T), "CHARFORMAT ")[[1]])) # It does not parse correctly
                                PatientWeight <-sapply(strsplit(find_field_grep("Weight:", TABLES, type = "col_next", tb.pos=1, fixed = T), "CHARFORMAT "), "[[", length(strsplit(find_field_grep("Weight:", TABLES, type = "col_next", tb.pos=1, fixed = T), "CHARFORMAT ")[[1]]))
                                
                                main_values_df <- data.frame(
                                        
                                        ### Patient Table
                                        PatientGender=sapply(strsplit(find_field_grep("Sex:", TABLES, type = "col_next", tb.pos=1, fixed = T), "CHARFORMAT "), "[[", 2),
                                        PatientDOB=sapply(strsplit(find_field_grep("D.O.B:", TABLES, type = "col_next", tb.pos=1, fixed = T), "CHARFORMAT "), "[[", length(strsplit(find_field_grep("D.O.B:", TABLES, type = "col_next", tb.pos=1, fixed = T), "CHARFORMAT ")[[1]])), # It does not parse correctly
                                        PatientAge=sapply(strsplit(find_field_grep("Age:", TABLES, type = "col_next", tb.pos=1, fixed = T), "CHARFORMAT "), "[[", 2),
                                        PatientRecorderID=sapply(strsplit(find_field_grep("Recorder ID:", TABLES, type = "col_next", tb.pos=1, fixed = T), "CHARFORMAT "), "[[", 2),
                                        StudyDate=sapply(strsplit(find_field_grep("Study Date:", TABLES, type = "col_next", tb.pos=1, fixed = T), "CHARFORMAT "), "[[", 2),
                                        PatientMRN=sapply(strsplit(find_field_grep("MRN:", TABLES, type = "col_next", tb.pos=1, fixed = T), "CHARFORMAT "), "[[", 2),
                                        ReferringPhysician=sapply(strsplit(find_field_grep("Referring Physician:", TABLES, type = "col_next", tb.pos=1, fixed = T), "CHARFORMAT "),"[[", length(strsplit(find_field_grep("Referring Physician:", TABLES, type = "col_next", tb.pos=1, fixed = T), "CHARFORMAT ")[[1]])),
                                        SleepSpecialist=sapply(strsplit(find_field_grep("Sleep Specialist:", TABLES, type = "col_next", tb.pos=1, fixed = T), "CHARFORMAT "), "[[", length(strsplit(find_field_grep("Sleep Specialist:", TABLES, type = "col_next", tb.pos=1, fixed = T), "CHARFORMAT ")[[1]])),
                                        PatientBMI=sapply(strsplit(find_field_grep("B.M.I:", TABLES, type = "col_next", tb.pos=1, fixed = T), "CHARFORMAT "), "[[", length(strsplit(find_field_grep("B.M.I:", TABLES, type = "col_next", tb.pos=1, fixed = T), "CHARFORMAT ")[[1]])),
                                        ProcedureCode=sapply(strsplit(find_field_grep("Procedure Code:", TABLES, type = "col_next", tb.pos=1, fixed = T), "CHARFORMAT "), "[[", length(strsplit(find_field_grep("Procedure Code:", TABLES, type = "col_next", tb.pos=1, fixed = T), "CHARFORMAT ")[[1]])),
                                        
                                        PatientFirstName=trim(strsplit(PatientFullName, ",")[[1]][2]),
                                        PatientLastName=strsplit(PatientFullName, ",")[[1]][1],
                                        PatientWeight_value=strsplit(PatientWeight, " ")[[1]][1],
                                        PatientWeight_unit=strsplit(PatientWeight, " ")[[1]][2],
                                        
                                        #### Fix DOB
                                        #### Fix Height
                                        PatientHeight_value=PatientHeight,
                                        
                                        
                                        
                                        # Get other relevant values from the word document
                                        study_type="HSAT",
                                        clinical_history=grep("CLINICAL HISTORY:", textreadr::read_docx(PSGfilename), value = T),
                                        final_diagnosis=grep("DIAGNOSIS:", textreadr::read_docx(PSGfilename), value = T),
                                        comm_rec=grep("RECOMMENDATIONS:", textreadr::read_docx(PSGfilename), value = T),
                                        
                                        RecordingTime=gsub("minutes+", "", get_values(string = "ESStatInfo_Resp_IndexTime", source = textreadr::read_docx(PSGfilename))),
                                        Total_AH_count=get_values(string = "ESStatInfo_Resp_Count", source = textreadr::read_docx(PSGfilename)),
                                        TOtal_AH_index=gsub("/+", "", get_values(string = "ESStatInfo_Sleep_TotalSleepTime", source = textreadr::read_docx(PSGfilename))),
                                        Supine_AHI_count=get_values(string = "ESStatInfo_Resp_Count_2", source = textreadr::read_docx(PSGfilename)),
                                        Supine_AHI_index=gsub("/+", "", get_values(string = "ESStatInfo_Sleep_TotalSleepTime_2", source = textreadr::read_docx(PSGfilename))),
                                        NonSupine_AHI_count=get_values(string = "ESStatInfo_Resp_Count_3", source = textreadr::read_docx(PSGfilename)),
                                        NonSupine_AHI_index=gsub("/+", "", get_values(string = "ESStatInfo_Sleep_TotalSleepTime_3", source = textreadr::read_docx(PSGfilename))),
                                        Supine_time=gsub("minutes+", "", get_values(string = "ESStatInfo_Sleep_Duration", source = textreadr::read_docx(PSGfilename))),
                                        Supine_perc=gsub("%+", "", get_values(string = "ESStatInfo_Resp_IndexTime_2", source = textreadr::read_docx(PSGfilename))),
                                        NonSupine_time=gsub("minutes+", "", get_values(string = "ESStatInfo_Sleep_Duration_24", source = textreadr::read_docx(PSGfilename))),
                                        NonSupine_perc=gsub("%+", "", get_values(string = "ESStatInfo_Sleep_TotalSleepTime_5", source = textreadr::read_docx(PSGfilename))),
                                        Upright_time=gsub("minutes+", "", get_values(string = "ESStatInfo_Sleep_Duration_25", source = textreadr::read_docx(PSGfilename))),
                                        Upright_perc=gsub("%+", "", get_values(string = "ESStatInfo_Sleep_TotalSleepTime_6", source = textreadr::read_docx(PSGfilename))),
                                        MeanSaO2=gsub("%+", "", get_values(string = "ESStatInfo_SaO2_Mean_2", source = textreadr::read_docx(PSGfilename))),
                                        OD_count=get_values(string = "ESStatInfo_DesatStatistics_Fall_Count_12", source = textreadr::read_docx(PSGfilename)),
                                        OD_index=gsub("/.+", "", get_values(string = "ESStatInfo_DesatStatistics_Fall_index_2", source = textreadr::read_docx(PSGfilename))),
                                        Snore_time=gsub("minutes+", "", get_values(string = "ESStatInfo_Snoring_trainduration_4", source = textreadr::read_docx(PSGfilename))),
                                        Snore_perc=gsub("%+", "", get_values(string = "ESStatInfo_Snoring_RelativeSnoreTime_2", source = textreadr::read_docx(PSGfilename))),
                                        Snore_episodes=get_values(string = "ESStatInfo_Snoring_trains_3", source = textreadr::read_docx(PSGfilename)),
                                        
                                        # Respiratory events:
                                        
                                        # Counts
                                        Apnea_counts=get_values(string = "ESStatInfo_Resp_Count_15", source = textreadr::read_docx(PSGfilename)),
                                        ObsApnea_counts=get_values(string = "ESStatInfo_Resp_Count_16", source = textreadr::read_docx(PSGfilename)),
                                        CentralApnea_counts=get_values(string = "ESStatInfo_Resp_Count_17", source = textreadr::read_docx(PSGfilename)),
                                        MixedApnea_counts=get_values(string = "ESStatInfo_Resp_Count_18", source = textreadr::read_docx(PSGfilename)),
                                        Hypopnea_counts=get_values(string = "ESStatInfo_Resp_Count_19", source = textreadr::read_docx(PSGfilename)),
                                        TotalApneaHypopnea_counts=get_values(string = "ESStatInfo_Resp_Count_20", source = textreadr::read_docx(PSGfilename)),
                                        
                                        # Percentage
                                        Apnea_perc=get_values(string = "ESStatInfo_Resp_Count_21", source = textreadr::read_docx(PSGfilename)),
                                        ObsApnea_perc=get_values(string = "ESStatInfo_Resp_Count_22", source = textreadr::read_docx(PSGfilename)),
                                        CentralApnea_perc=get_values(string = "ESStatInfo_Resp_Count_23", source = textreadr::read_docx(PSGfilename)),
                                        MixedApnea_perc=get_values(string = "ESStatInfo_Resp_Count_24", source = textreadr::read_docx(PSGfilename)),
                                        Hypopnea_perc=get_values(string = "ESStatInfo_Resp_Count_25", source = textreadr::read_docx(PSGfilename)),
                                        
                                        # Indices
                                        Apnea_index=get_values(string = "ESStatInfo_Resp_index", source = textreadr::read_docx(PSGfilename)),
                                        ObsApnea_index=get_values(string = "ESStatInfo_Resp_index_2", source = textreadr::read_docx(PSGfilename)),
                                        CentralApnea_index=get_values(string = "ESStatInfo_Resp_index_3", source = textreadr::read_docx(PSGfilename)),
                                        MixedApnea_index=get_values(string = "ESStatInfo_Resp_index_4", source = textreadr::read_docx(PSGfilename)),
                                        Hypopnea_index=get_values(string = "ESStatInfo_Resp_index_5", source = textreadr::read_docx(PSGfilename)),
                                        TotalApneaHypopnea_index=get_values(string = "ESStatInfo_Resp_index_6", source = textreadr::read_docx(PSGfilename)),
                                        
                                        # Supine counts
                                        Apnea_Supcounts=get_values(string = "ESStatInfo_Resp_Count_26", source = textreadr::read_docx(PSGfilename)),
                                        ObsApnea_Supcounts=get_values(string = "ESStatInfo_Resp_Count_27", source = textreadr::read_docx(PSGfilename)),
                                        CentralApnea_Supcounts=get_values(string = "ESStatInfo_Resp_Count_28", source = textreadr::read_docx(PSGfilename)),
                                        MixedApnea_Supcounts=get_values(string = "ESStatInfo_Resp_Count_29", source = textreadr::read_docx(PSGfilename)),
                                        Hypopnea_Supcounts=get_values(string = "ESStatInfo_Resp_Count_30", source = textreadr::read_docx(PSGfilename)),
                                        TotalApneaHypopnea_Supcounts=get_values(string = "ESStatInfo_Resp_Count_31", source = textreadr::read_docx(PSGfilename)),
                                        
                                        # NonSupine counts
                                        Apnea_NonSupcounts=get_values(string = "ESStatInfo_Resp_Count_32", source = textreadr::read_docx(PSGfilename)),
                                        ObsApnea_NonSupcounts=get_values(string = "ESStatInfo_Resp_Count_33", source = textreadr::read_docx(PSGfilename)),
                                        CentralApnea_NonSupcounts=get_values(string = "ESStatInfo_Resp_Count_34", source = textreadr::read_docx(PSGfilename)),
                                        MixedApnea_NonSupcounts=get_values(string = "ESStatInfo_Resp_Count_35", source = textreadr::read_docx(PSGfilename)),
                                        Hypopnea_NonSupcounts=get_values(string = "ESStatInfo_Resp_Count_36", source = textreadr::read_docx(PSGfilename)),
                                        TotalApneaHypopnea_NonSupcounts=get_values(string = "ESStatInfo_Resp_Count_37", source = textreadr::read_docx(PSGfilename)),
                                        
                                        # Mean Duration
                                        Apnea_meandur=get_values(string = "ESStatInfo_Resp_MeanDuration_7", source = textreadr::read_docx(PSGfilename)),
                                        ObsApnea_meandur=get_values(string = "ESStatInfo_Resp_MeanDuration_8", source = textreadr::read_docx(PSGfilename)),
                                        CentralApnea_meandur=get_values(string = "ESStatInfo_Resp_MeanDuration_9", source = textreadr::read_docx(PSGfilename)),
                                        MixedApnea_meandur=get_values(string = "ESStatInfo_Resp_MeanDuration_10", source = textreadr::read_docx(PSGfilename)),
                                        Hypopnea_meandur=get_values(string = "ESStatInfo_Resp_MeanDuration_11", source = textreadr::read_docx(PSGfilename)),
                                        TotalApneaHypopnea_meandur=get_values(string = "ESStatInfo_Resp_MeanDuration_12", source = textreadr::read_docx(PSGfilename)),
                                        
                                        # Longest Duration
                                        Apnea_longndur=get_values(string = "ESStatInfo_Resp_LongestDuration_7", source = textreadr::read_docx(PSGfilename)),
                                        ObsApnea_longndur=get_values(string = "ESStatInfo_Resp_LongestDuration_8", source = textreadr::read_docx(PSGfilename)),
                                        CentralApnea_longndur=get_values(string = "ESStatInfo_Resp_LongestDuration_9", source = textreadr::read_docx(PSGfilename)),
                                        MixedApnea_longndur=get_values(string = "ESStatInfo_Resp_LongestDuration_10", source = textreadr::read_docx(PSGfilename)),
                                        Hypopnea_longndur=get_values(string = "ESStatInfo_Resp_LongestDuration_11", source = textreadr::read_docx(PSGfilename)),
                                        TotalApneaHypopnea_longndur=get_values(string = "ESStatInfo_Resp_LongestDuration_12", source = textreadr::read_docx(PSGfilename)),
                                        
                                        
                                        ## SpO2 statistics
                                        MeanSpO2=gsub("%.+", "", get_values(string = "ESStatInfo_SaO2_Mean_3", source = textreadr::read_docx(PSGfilename))),
                                        LowestSpO2=gsub("%.+", "", get_values(string = "ESStatInfo_SaO2_Lowest_4", source = textreadr::read_docx(PSGfilename))),
                                        AvgDesat=gsub("%.+", "", get_values(string = "ESStatInfo_DesatStatistics_AverageFall_3", source = textreadr::read_docx(PSGfilename))),
                                        Time_SpO2lt90=get_values(string = "ESStatInfo_SaO2_DurationRange_59", source = textreadr::read_docx(PSGfilename)),
                                        PercTime_SpO2lt90=gsub("%+.", "", get_values(string = "ESStatInfo_SaO2_DurationRange_62", source = textreadr::read_docx(PSGfilename))),
                                        Time_SpO2lt80=get_values(string = "ESStatInfo_SaO2_DurationRange_60", source = textreadr::read_docx(PSGfilename)),
                                        PercTime_SpO2lt80=gsub("%+", "", get_values(string = "ESStatInfo_SaO2_DurationRange_63", source = textreadr::read_docx(PSGfilename))),
                                        Time_SpO2lt70=get_values(string = "ESStatInfo_SaO2_DurationRange_61", source = textreadr::read_docx(PSGfilename)),
                                        PercTime_SpO2lt70=gsub("%+", "", get_values(string = "ESStatInfo_SaO2_DurationRange_64", source = textreadr::read_docx(PSGfilename))),
                                        
                                        # Pulse rate
                                        
                                        MeanHR_total=get_values(string = "ESStatInfo_PulseStatistics_Average", source = textreadr::read_docx(PSGfilename)),
                                        MeanHR_Sup=get_values(string = "ESStatInfo_PulseStatistics_Average_2", source = textreadr::read_docx(PSGfilename)),
                                        MeanHR_NonSup=get_values(string = "ESStatInfo_PulseStatistics_Average_3", source = textreadr::read_docx(PSGfilename)),
                                        
                                        MinHR_total=get_values(string = "ESStatInfo_PulseStatistics_Min", source = textreadr::read_docx(PSGfilename)),
                                        MinHR_Sup=get_values(string = "ESStatInfo_PulseStatistics_Min_2", source = textreadr::read_docx(PSGfilename)),
                                        MinHR_NonSup=get_values(string = "ESStatInfo_PulseStatistics_Min_3", source = textreadr::read_docx(PSGfilename)),
                                        
                                        MaxHR_total=get_values(string = "ESStatInfo_PulseStatistics_Max", source = textreadr::read_docx(PSGfilename)),
                                        MaxHR_Sup=get_values(string = "ESStatInfo_PulseStatistics_Max_2", source = textreadr::read_docx(PSGfilename)),
                                        MaxHR_NonSup=get_values(string = "ESStatInfo_PulseStatistics_Max_3", source = textreadr::read_docx(PSGfilename)),
                                        
                                        
                                        stringsAsFactors = F)
                                
                                
                                # Fix as numeric
                                main_values_df[,c(3,9,13,20:97)] <- sapply(main_values_df[,c(3,9,13,20:97)], as.numeric)
                                
                                
                                
                                # Get SpO2 position data
                                sp02_pattern <- paste(paste0(rep("ESStatInfo_SaO2_DurationRange_", 8), 67:74), collapse = "|")
                                sp_list <- strsplit(grep(sp02_pattern, textreadr::read_docx(PSGfilename), value = T), " ")
                                
                                sp_names <-  c("SpO2_98_100_",
                                               "SpO2_95_97_",
                                               "SpO2_90_94_",
                                               "SpO2_80_89_",
                                               "SpO2_70_79_",
                                               "SpO2_60_69_",
                                               "SpO2_50_59_",
                                               "SpO2_lt_50_")
                                
                                sp_cols <- c("Sup_minutes",
                                             "CumlSup_minutes",
                                             "NonSup_minutes",
                                             "CumlNonSup_minutes",
                                             "Upright_minutes",
                                             "UprightCuml_minutes")
                                
                                sp_num_results <- list()
                                sp_idx=1
                                for (sp_el in sp_list) {
                                        
                                        #sp_el <- sp_list[[1]]
                                        if(sp_el[1]=="<") {
                                                num_sp <- as.numeric(sp_el)
                                                num_sp <- num_sp[!is.na(num_sp)][-1]
                                        } else {
                                                num_sp <- as.numeric(sp_el)
                                                num_sp <- num_sp[!is.na(num_sp)]
                                        }
                                        
                                        sp_df <- as.data.frame(t(num_sp))
                                        colnames(sp_df) <- paste0(sp_names[sp_idx],sp_cols)
                                        
                                        sp_num_results[[sp_idx]] <- sp_df
                                        sp_idx=sp_idx+1
                                        
                                }
                                
                                sp_position_df <- bind_cols(sp_num_results)
                                
                                ### Fix names
                                colnames(main_values_df)[c(2,4,5,6,7,11,12,17,18,19)] <- paste0("H08_PSG_PHI_", colnames(main_values_df)[c(2,4,5,6,7,11,12,17,18,19)])
                                colnames(main_values_df)[-c(2,4,5,6,7,11,12,17,18,19)] <- paste0("H08_PSG_", colnames(main_values_df)[-c(2,4,5,6,7,11,12,17,18,19)])
                                colnames(sp_position_df) <- paste0("H08_PSG_", colnames(sp_position_df))
                             
                                ######## Combine ALL data
                                all_fields <- c("H08_PSG_PHI_Filename", "PennSleepID", colnames(main_values_df), colnames(sp_position_df))
                                
                                #Initiate dataframe with results
                                
                                #If first, create final_df
                                if (first) {
                                        final_df <- data.frame(matrix(NA,1,length(all_fields)))
                                        colnames(final_df) <- all_fields
                                        
                                        final_df[1,] <- c(PSGfilename,
                                                          NA,
                                                          main_values_df[1,],
                                                          sp_position_df[1,])
                                        first <- F
                                } else {
                                        
                                        next_final_df <- data.frame(matrix(NA,1,length(all_fields)))
                                        colnames(next_final_df) <- all_fields
                                        
                                        next_final_df[1,] <- c(PSGfilename,
                                                               NA,
                                                               main_values_df[1,],
                                                               sp_position_df[1,])
                                        
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
        final_df$H08_PSG_Age_at_Study <- interval(mdy(final_df$H08_PSG_PHI_PatientDOB), mdy(final_df$H08_PSG_PHI_StudyDate)) %/% years(1)
        
        #Create identifiable data frame
        identifiable_df <- final_df
        identifiable_df$ProcessedDate <- Sys.Date()
        
        ProcessedTimeID=gsub("-","",gsub(":","",gsub(" ", "", Sys.time())))
        
        identifiable_df$ProcessedTimeID <- ProcessedTimeID
        
        # Per sample QC (not value filter)
        
        QC_df <- data.frame(H08_PSG_PHI_Filename=identifiable_df$H08_PSG_PHI_Filename,
                            # Missing both MRN fields
                            has_missingMRN=is.na(identifiable_df$H08_PSG_PHI_PatientMRN),
                            # Gender not (Male, Female, Unknown)
                            has_missingGender=!(tolower(identifiable_df$H08_PSG_PatientGender) %in% c("male", "female", "unknown")),
                            # Missing age
                            has_missingAge=(is.na(identifiable_df$H08_PSG_PatientAge) & is.na(identifiable_df$H08_PSG_Age_at_Study)),
                            # Missing BMI
                            has_missingBMI=is.na(identifiable_df$H08_PSG_PatientBMI)
        )
        
        identifiable_df$PerSample_QC <- NA
        identifiable_df$PerSample_QC[rowSums(QC_df[,2:5])>0] <- "FAIL"
        identifiable_df$PerSample_QC[rowSums(QC_df[,2:5])==0] <- "PASS"
        
        QC_df$PerSample_QC <- identifiable_df$PerSample_QC
        write.csv(QC_df, paste0("QC_dataframe_", ProcessedTimeID, ".csv"))
        
        #Save identifiable Rdata file encrypted
        #key <- key_sodium(sodium::keygen())
        saveRDS(identifiable_df, paste0("PennSleepDatabase_HSAT08_", ProcessedTimeID, ".Rdata"))
        #cyphr::encrypt_file("PennSleepDatabase.Rdata", key, "PennSleepDatabase.encrypted")
        
        #Save de-identified version in CSV files - implement a way to check in database if sample was processed and not use the same de-identified IDs
        #identifiable_df$PennSleepID <- paste0("PENNSLEEP00000",seq(1:nrow(identifiable_df)))
        deidentified_df <- select(identifiable_df, -starts_with("PSG_PHI_"))
        
        write.csv(deidentified_df, paste0("PennSleepDatabase_HSAT08_Deidentified",ProcessedTimeID,".csv"), row.names = F)
        
        return(identifiable_df)
        
}

