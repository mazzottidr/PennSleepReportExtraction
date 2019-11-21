# Function to extract feaatures from docx documents
# Diego Mazzotti
# January 2019
# University of Pennsylvania


library(officer)
library(dplyr)
library(docxtractr)


extract_docx_features <- function(my.paths) {
        
        # Calculate some basic features based on the summary
        ncols=72
        docx_features <- as.data.frame(matrix(NA, nrow = length(my.paths), ncol=ncols))
        
        #Define colnames
        col_names <- c("styles_nrow","styles_N_unique_id","styles_N_is_custom",
                       "header_l","footer_l",
                       "sect_dim_width","sect_dim_height","sect_dim_margins_top","sect_dim_margins_bottom","sect_dim_margins_left","sect_dim_margins_right","sect_dim_margins_header","sect_dim_margins_footer",
                       "docx_summ_l",
                       "docx_summ_max_index","docx_summ_mean_index","docx_summ_sd_index",
                       "docx_summ_content_type_paragraph_count","docx_summ_content_type_tablecell_count",
                       "docx_summ_max_row_id","docx_summ_mean_row_id","docx_summ_sd_row_id",
                       "docx_summ_max_cell_id","docx_summ_mean_cell_id","docx_summ_sd_cell_id",
                       "docx_summ_max_col_span","docx_summ_mean_col_span","docx_summ_sd_col_span",
                       "docx_summ_max_row_span","docx_summ_mean_row_span","docx_summ_sd_row_span",
                       "docx_summ_shifts","docx_summ_min_run_length","docx_summ_max_run_length","docx_summ_median_run_length",
                       "docx_tbl_N",
                       "num_baseline_mentions","first_baseline_mention_id","last_baseline_mention_id","median_baseline_mention_id",
                       "num_treatment_mentions","first_treatment_mention_id","last_treatment_mention_id","median_treatment_mention_id",
                       "num_split_mentions","first_split_mention_id","last_split_mention_id","median_split_mention_id",
                       "num_multiple_mentions","first_multiple_mention_id","last_multiple_mention_id","median_multiple_mention_id",
                       "num_mslt_mentions","first_mslt_mention_id","last_mslt_mention_id","median_mslt_mention_id",
                       "num_cpap_mentions", "first_cpap_mention_id","last_cpap_mention_id", "median_cpap_mention_id",
                       "num_mask_mentions", "first_mask_mention_id","last_mask_mention_id", "median_mask_mention_id",
                       "num_bipap_mentions", "first_bipap_mention_id","last_bipap_mention_id", "median_bipap_mention_id",
                       "num_DOCPROPERTY_mentions", "first_DOCPROPERTY_mention_id","last_DOCPROPERTY_mention_id", "median_DOCPROPERTY_mention_id")
        
        colnames(docx_features) <- col_names
        
        i=1
        for (docx_path in my.paths) {
                
                # Copy and rename docx before processsing to avoid error with file names
                file.copy(docx_path, to = ".")
                
                docx <- sapply(strsplit(docx_path, "/"), "[[", 2)
                
                current_tmp_files <- list.files(tempdir(), full.names = T)
                
                message(paste0("Processing: ", docx))
                message(paste0(format(100*i/length(my.paths), digits = 4),"% done,")) 
                
                possibleError <- tryCatch( {
                        current_docx <- officer::read_docx(docx)
                        current_docx_2 <- docxtractr::read_docx(docx)
                }, error=function(e) e)
                
                if (inherits(possibleError, "error")) {
                        
                        message(paste0("Sample ", docx_path, " did not work!"))
                        message(possibleError)
                        write(docx_path,file="failed.txt",append=TRUE)
                        
                        # Delete corresponding folder
                        new_tmp_files <- list.files(tempdir(), full.names = T, recursive = T)
                        tmp_files_toRemove <- new_tmp_files[!(new_tmp_files %in% current_tmp_files)]
                        tmp_folder_toRemove <- unique(paste(sapply(strsplit(tmp_files_toRemove, "/"), "[[", 1), sapply(strsplit(tmp_files_toRemove, "/"), "[[", 2), sep = "/"))
                        unlink(tmp_folder_toRemove, recursive = T)
                        
                        i=i+1
                        next()
                }
                
                
                
                
                #### Features from officer package
                
                # Load file
                #current_docx <- officer::read_docx(docx)
                
                # File name
                docx_name <- docx
                rownames(docx_features)[i] <- docx_name
                
                # Get styles_df
                styles_df <- current_docx$styles
                
                # Get styles_df features
                styles_nrow <- nrow(styles_df)
                styles_N_unique_id <- length(unique(styles_df$style_id))
                styles_N_is_custom <- sum(styles_df$is_custom)
                
                # Get length of header and footer
                header_l <- length(current_docx$headers)
                footer_l <- length(current_docx$footers)
                
                # Get section dimensions
                sect_dim_width <- current_docx$sect_dim$page[1]
                sect_dim_height  <- current_docx$sect_dim$page[2]
                sect_dim_margins_top  <- current_docx$sect_dim$margins[1]
                sect_dim_margins_bottom  <- current_docx$sect_dim$margins[2]
                sect_dim_margins_left <- current_docx$sect_dim$margins[3]
                sect_dim_margins_right <- current_docx$sect_dim$margins[4]
                sect_dim_margins_header <- current_docx$sect_dim$margins[5]
                sect_dim_margins_footer <- current_docx$sect_dim$margins[6]
                
                # Explore docx_summary
                current_docx_summary <- docx_summary(current_docx)
                
                docx_summ_l <- nrow(current_docx_summary)
                
                docx_summ_max_index <- max(current_docx_summary$doc_index, na.rm = T)
                docx_summ_mean_index <- mean(current_docx_summary$doc_index, na.rm = T)
                docx_summ_sd_index <- sd(current_docx_summary$doc_index, na.rm = T)
                
                docx_summ_content_type_paragraph_count <- summary(as.factor(current_docx_summary$content_type))[1]
                docx_summ_content_type_tablecell_count <- summary(as.factor(current_docx_summary$content_type))[2] # Make sure there aren't other classes here
                
                docx_summ_max_row_id <- max(current_docx_summary$row_id, na.rm = T)
                docx_summ_mean_row_id <- mean(current_docx_summary$row_id, na.rm = T)
                docx_summ_sd_row_id <- sd(current_docx_summary$row_id, na.rm = T)
                
                docx_summ_max_cell_id <- max(current_docx_summary$cell_id, na.rm = T)
                docx_summ_mean_cell_id <- mean(current_docx_summary$cell_id, na.rm = T)
                docx_summ_sd_cell_id <- sd(current_docx_summary$cell_id, na.rm = T)
                
                docx_summ_max_col_span <- max(current_docx_summary$col_span, na.rm = T)
                docx_summ_mean_col_span <- mean(current_docx_summary$col_span, na.rm = T)
                docx_summ_sd_col_span <- sd(current_docx_summary$col_span, na.rm = T)
                
                docx_summ_max_row_span <- max(current_docx_summary$row_span, na.rm = T)
                docx_summ_mean_row_span <- mean(current_docx_summary$row_span, na.rm = T)
                docx_summ_sd_row_span <- sd(current_docx_summary$row_span, na.rm = T)
                
                # Lengths of runs of tables and paragraphs
                docx_summ_shifts <- length(rle(current_docx_summary$content_type)$lengths)
                docx_summ_min_run_length <- min(rle(current_docx_summary$content_type)$lengths, na.rm = T)
                docx_summ_max_run_length <- max(rle(current_docx_summary$content_type)$lengths, na.rm = T)
                docx_summ_median_run_length <- median(rle(current_docx_summary$content_type)$lengths, na.rm = T)
                
                
                #### Features from docxtractr package
                #current_docx_2 <- docxtractr::read_docx(docx)
                docx_tbl_N <- docx_tbl_count(current_docx_2)
                
                
                #### Other fetaures relataed to regular expressions
                lowered_text <- tolower(current_docx_summary$text)
                
                num_baseline_mentions <- ifelse(length(grep("baseline",lowered_text))==0, 0, length(grep("baseline",lowered_text)))
                first_baseline_mention_id <- ifelse(length(grep("baseline",lowered_text))==0, NA, grep("baseline",lowered_text)[1])
                last_baseline_mention_id <- ifelse(length(grep("baseline",lowered_text))==0, NA, tail(grep("baseline",lowered_text),n=1))
                median_baseline_mention_id <- ifelse(length(grep("baseline",lowered_text))==0, NA, median(grep("baseline",lowered_text)))
                
                num_treatment_mentions <- ifelse(length(grep("treatment",lowered_text))==0, 0, length(grep("treatment",lowered_text)))
                first_treatment_mention_id <- ifelse(length(grep("treatment",lowered_text))==0, NA, grep("treatment",lowered_text)[1])
                last_treatment_mention_id <- ifelse(length(grep("treatment",lowered_text))==0, NA, tail(grep("treatment",lowered_text),n=1))
                median_treatment_mention_id <- ifelse(length(grep("treatment",lowered_text))==0, NA, median(grep("treatment",lowered_text)))
                
                num_split_mentions <- ifelse(length(grep("split",lowered_text))==0, 0, length(grep("split",lowered_text)))
                first_split_mention_id <- ifelse(length(grep("split",lowered_text))==0, NA, grep("split",lowered_text)[1])
                last_split_mention_id <- ifelse(length(grep("split",lowered_text))==0, NA, tail(grep("split",lowered_text),n=1))
                median_split_mention_id <- ifelse(length(grep("split",lowered_text))==0, NA, median(grep("split",lowered_text)))
                
                num_multiple_mentions <- ifelse(length(grep("multiple",lowered_text))==0, 0, length(grep("multiple",lowered_text)))
                first_multiple_mention_id <- ifelse(length(grep("multiple",lowered_text))==0, NA, grep("multiple",lowered_text)[1])
                last_multiple_mention_id <- ifelse(length(grep("multiple",lowered_text))==0, NA, tail(grep("multiple",lowered_text),n=1))
                median_multiple_mention_id <- ifelse(length(grep("multiple",lowered_text))==0, NA, median(grep("multiple",lowered_text)))
                
                num_mslt_mentions <- ifelse(length(grep("mslt",lowered_text))==0, 0, length(grep("mslt",lowered_text)))
                first_mslt_mention_id <- ifelse(length(grep("mslt",lowered_text))==0, NA, grep("mslt",lowered_text)[1])
                last_mslt_mention_id <- ifelse(length(grep("mslt",lowered_text))==0, NA, tail(grep("mslt",lowered_text),n=1))
                median_mslt_mention_id <- ifelse(length(grep("mslt",lowered_text))==0, NA, median(grep("mslt",lowered_text)))
                
                #Other ideas: cpap, mask, bipap
                num_cpap_mentions <- ifelse(length(grep("cpap",lowered_text))==0, 0, length(grep("cpap",lowered_text)))
                first_cpap_mention_id <- ifelse(length(grep("cpap",lowered_text))==0, NA, grep("cpap",lowered_text)[1])
                last_cpap_mention_id <- ifelse(length(grep("cpap",lowered_text))==0, NA, tail(grep("cpap",lowered_text),n=1))
                median_cpap_mention_id <- ifelse(length(grep("cpap",lowered_text))==0, NA, median(grep("cpap",lowered_text)))
                
                num_mask_mentions <- ifelse(length(grep("mask",lowered_text))==0, 0, length(grep("mask",lowered_text)))
                first_mask_mention_id <- ifelse(length(grep("mask",lowered_text))==0, NA, grep("mask",lowered_text)[1])
                last_mask_mention_id <- ifelse(length(grep("mask",lowered_text))==0, NA, tail(grep("mask",lowered_text),n=1))
                median_mask_mention_id <- ifelse(length(grep("mask",lowered_text))==0, NA, median(grep("mask",lowered_text)))
                
                num_bipap_mentions <- ifelse(length(grep("bipap",lowered_text))==0, 0, length(grep("bipap",lowered_text)))
                first_bipap_mention_id <- ifelse(length(grep("bipap",lowered_text))==0, NA, grep("bipap",lowered_text)[1])
                last_bipap_mention_id <- ifelse(length(grep("bipap",lowered_text))==0, NA, tail(grep("bipap",lowered_text),n=1))
                median_bipap_mention_id <- ifelse(length(grep("bipap",lowered_text))==0, NA, median(grep("bipap",lowered_text)))
                
                num_DOCPROPERTY_mentions <- ifelse(length(grep("docproperty",lowered_text))==0, 0, length(grep("docproperty",lowered_text)))
                first_DOCPROPERTY_mention_id <- ifelse(length(grep("docproperty",lowered_text))==0, NA, grep("docproperty",lowered_text)[1])
                last_DOCPROPERTY_mention_id <- ifelse(length(grep("docproperty",lowered_text))==0, NA, tail(grep("docproperty",lowered_text),n=1))
                median_DOCPROPERTY_mention_id <- ifelse(length(grep("docproperty",lowered_text))==0, NA, median(grep("docproperty",lowered_text)))
                
                
                # Create row with feautures and assign to docx_features data.frame
                docx_features[i,] <- c(styles_nrow, styles_N_unique_id, styles_N_is_custom,
                                       header_l, footer_l,
                                       sect_dim_width, sect_dim_height, sect_dim_margins_top, sect_dim_margins_bottom, sect_dim_margins_left, sect_dim_margins_right, sect_dim_margins_header, sect_dim_margins_footer,
                                       docx_summ_l,
                                       docx_summ_max_index, docx_summ_mean_index, docx_summ_sd_index,
                                       docx_summ_content_type_paragraph_count, docx_summ_content_type_tablecell_count,
                                       docx_summ_max_row_id, docx_summ_mean_row_id, docx_summ_sd_row_id,
                                       docx_summ_max_cell_id, docx_summ_mean_cell_id, docx_summ_sd_cell_id,
                                       docx_summ_max_col_span, docx_summ_mean_col_span, docx_summ_sd_col_span,
                                       docx_summ_max_row_span, docx_summ_mean_row_span, docx_summ_sd_row_span,
                                       docx_summ_shifts,docx_summ_min_run_length, docx_summ_max_run_length, docx_summ_median_run_length,
                                       docx_tbl_N,
                                       num_baseline_mentions, first_baseline_mention_id, last_baseline_mention_id, median_baseline_mention_id,
                                       num_treatment_mentions, first_treatment_mention_id, last_treatment_mention_id, median_treatment_mention_id,
                                       num_split_mentions, first_split_mention_id, last_split_mention_id, median_split_mention_id,
                                       num_multiple_mentions, first_multiple_mention_id, last_multiple_mention_id, median_multiple_mention_id,
                                       num_mslt_mentions, first_mslt_mention_id,last_mslt_mention_id, median_mslt_mention_id,
                                       num_cpap_mentions, first_cpap_mention_id,last_cpap_mention_id, median_cpap_mention_id,
                                       num_mask_mentions, first_mask_mention_id,last_mask_mention_id, median_mask_mention_id,
                                       num_bipap_mentions, first_bipap_mention_id,last_bipap_mention_id, median_bipap_mention_id,
                                       num_DOCPROPERTY_mentions, first_DOCPROPERTY_mention_id,last_DOCPROPERTY_mention_id, median_DOCPROPERTY_mention_id)
                
               
                write.csv(docx_features, "docx_features_bkp.csv", row.names = T)
                i=i+1
                
                
                # Delete corresponding folder
                new_tmp_files <- list.files(tempdir(), full.names = T, recursive = T)
                tmp_files_toRemove <- new_tmp_files[!(new_tmp_files %in% current_tmp_files)]
                tmp_folder_toRemove <- unique(paste(sapply(strsplit(tmp_files_toRemove, "/"), "[[", 1), sapply(strsplit(tmp_files_toRemove, "/"), "[[", 2), sep = "/"))
                unlink(tmp_folder_toRemove, recursive = T)
                
                # Delete file from working directory
                file.remove(docx)
                
        }
        
        
        
        # Replace infinite and NA values with 0
        docx_features_no_NA <- docx_features
        docx_features_no_NA[is.na(docx_features)] <- 0
        docx_features_no_NA[sapply(docx_features, is.infinite)] <- 0
        
        saveRDS(docx_features_no_NA, paste0("docx_features_", gsub("-", "_", gsub(":", "_", gsub(" ", "_",Sys.time()))), ".Rdata"))
        return(docx_features_no_NA)
        
}