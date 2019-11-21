library(dplyr)
library(lubridate)
library(stringr)
library(xlsx)
library(limma)

##### Next steps

# Load all extracted data together OK

# Re-run quality control for each dataset

# Compile how many were extracted appropriately OK


#### Combining, checking quality and cleaning PennSleepDatabase
setwd("/data1/home/diegomaz/ALL_07232019/combined_data")

### Load metadata
PennSleepMetadata <- readRDS("PennSleepDatabase_Paths_ALL_07242019.Rdata")

# Load extracted sleep data
extracted_sleep_data <- list()
for (f in list.files(path = "extracted_Rdata/", full.names = T)) {
        
        name_f <- sapply(strsplit(f, "//"), "[[", 2)
        
        extracted_sleep_data[[name_f]] <- readRDS(f)
        
}
saveRDS(extracted_sleep_data, "ExtractedSleepData_lists.Rdata")



# Save column names for harmonization
colnames_list <- sapply(extracted_sleep_data, colnames)
for (c_id in 1:length(colnames_list)) {
        
        write.csv(colnames_list[c_id], paste0("ColumnNames_",names(colnames_list)[c_id], ".csv"), row.names = F)
}


# Mark whether file was extracted or not
for (d in extracted_sleep_data) {
        
        PennSleepMetadata$IsExtracted[PennSleepMetadata$file_name %in% sapply(strsplit(d[,1], "/"), "[[", length(strsplit(d[,1], "/")[[1]]))] <- T
        
}
saveRDS(PennSleepMetadata, "PennSleepMetadata_AfterExtraction_08162019.Rdata")
# > sum(PennSleepMetadata$IsExtracted)
# [1] 39467

#### Add GAP and PMBB annotations
GAP_metadata <- readRDS("all_plus_GAP_wData.Rdata")
GAP_metadata <- GAP_metadata[,1:19]
# Create Linux Path
GAP_metadata$linux_path <- paste0("/data1/home/diegomaz/sleepstudyreports_Dec2018/", GAP_metadata$file_name)

PMBB_metadata <- read.xlsx("sleep_study_pmbb_genotype_consent_status.xlsx", sheetIndex = 1, stringsAsFactors=F)
PMBB_metadata_consented <- filter(PMBB_metadata, consent=="Y") %>% distinct()

# Merge PennSleep with GAP Metadata
PennSleepMetadata_wGAP <- merge(PennSleepMetadata, GAP_metadata, by.x="linux_path", by.y="linux_path", all.x=T)
PennSleepMetadata_wGAP$pt_MRN.x <- as.character(PennSleepMetadata_wGAP$pt_MRN.x)

# Merge with PMBB_consented
PennSleepMetadata_wGAP_wPMBB <- left_join(PennSleepMetadata_wGAP, PMBB_metadata_consented, c("pt_MRN.x"="mrn")) # Be aware that there are some MRNs with more than one Genotyping ID


# Add number of studies per MRN
PennSleepMetadata_wGAP_wPMBB <- PennSleepMetadata_wGAP_wPMBB %>%
        group_by(pt_MRN.x) %>%
        add_tally()
colnames(PennSleepMetadata_wGAP_wPMBB)[colnames(PennSleepMetadata_wGAP_wPMBB)=="n"] <- "N_Studies"

#### Save
saveRDS(PennSleepMetadata_wGAP_wPMBB, "PennSleepMetadata_wGAP_wPMBB_081919.Rdata")

### Select relevant metadata
PennSleepMetadata_wGAP_wPMBB <- PennSleepMetadata_wGAP_wPMBB %>%
        select(pt_MRN.x, file_name.x, linux_path,
               predictedStudyType,
               predictedTableFormat,
               IsExtracted,
               N_Studies,
               GAP.ID., SAGIC.ID., TYPE.OF.QUESTIONNAIRE, DATE.OF.QUESTIONNAIRE, PAPER.ELECTRONIC, DX.PSG.DATE, TYPE.OF.PSG, REPORT.in.Chart, DNA.SAMPLE, TYPE.OF.SAMPLE, LOCATION.OF.SAMPLE, PICTURES, COMMENTS,
               consent, HUP_MRN, PMC_MRN, PAH_MRN, GENO_ID)
colnames(PennSleepMetadata_wGAP_wPMBB) <- c("SleepStudy_MRN", "file_name", "linux_path", "predictedStudyType", "predictedTableFormat", "IsExtracted", "N_Studies",
                                            "GAP_ID", "SAGIC_ID", "TypeQuestionnaire", "DateQuestionnaire", "PaperElectronic", "GAP_PSG_date", "GAP_PSG_type", "ReportInChart", "GAP_DNA_Sample", "GAP_TypeOfSample", "GAP_LocationOfSample", "GAP_Pictures", "GAP_Comments", "PMBB_consented", "HUP_MRN", "PMC_MRN", "PAH_MRN", "GENO_ID")


PennSleepMetadata_wGAP_wPMBB$GAP_DNA_Sample_clean <- NA
PennSleepMetadata_wGAP_wPMBB$GAP_DNA_Sample_clean[PennSleepMetadata_wGAP_wPMBB$GAP_DNA_Sample=="NO"] <- "No"
PennSleepMetadata_wGAP_wPMBB$GAP_DNA_Sample_clean[PennSleepMetadata_wGAP_wPMBB$GAP_DNA_Sample=="yes"] <- "Yes"
PennSleepMetadata_wGAP_wPMBB$GAP_DNA_Sample_clean[PennSleepMetadata_wGAP_wPMBB$GAP_DNA_Sample=="YES"] <- "Yes"
PennSleepMetadata_wGAP_wPMBB$GAP_DNA_Sample_clean[PennSleepMetadata_wGAP_wPMBB$GAP_DNA_Sample=="YES "] <- "Yes"
PennSleepMetadata_wGAP_wPMBB$GAP_DNA_Sample <- PennSleepMetadata_wGAP_wPMBB$GAP_DNA_Sample_clean
PennSleepMetadata_wGAP_wPMBB$GAP_DNA_Sample_clean <- NULL

saveRDS(PennSleepMetadata_wGAP_wPMBB, "PennSleepMetadata_wGAP_wPMBB_081919.Rdata")

# Create data frame with Unique MRN counts, if it is available on GAP, PMBB and counts of studies per type
SleepStudy_GAP_PMBB_Summary <- data.frame(SleepStudy_MRN=unique(PennSleepMetadata_wGAP_wPMBB$SleepStudy_MRN), stringsAsFactors = F)
SleepStudy_GAP_PMBB_Summary$PMBB_consented <- SleepStudy_GAP_PMBB_Summary$SleepStudy_MRN %in% PennSleepMetadata_wGAP_wPMBB$SleepStudy_MRN[PennSleepMetadata_wGAP_wPMBB$PMBB_consented=="Y"]
SleepStudy_GAP_PMBB_Summary$PMBB_genotyped <- SleepStudy_GAP_PMBB_Summary$SleepStudy_MRN %in% PennSleepMetadata_wGAP_wPMBB$SleepStudy_MRN[!is.na(PennSleepMetadata_wGAP_wPMBB$GENO_ID)]
SleepStudy_GAP_PMBB_Summary$GAP_withID <- SleepStudy_GAP_PMBB_Summary$SleepStudy_MRN %in% PennSleepMetadata_wGAP_wPMBB$SleepStudy_MRN[!is.na(PennSleepMetadata_wGAP_wPMBB$GAP_ID)]
SleepStudy_GAP_PMBB_Summary$GAP_withDNA <- SleepStudy_GAP_PMBB_Summary$SleepStudy_MRN %in% PennSleepMetadata_wGAP_wPMBB$SleepStudy_MRN[PennSleepMetadata_wGAP_wPMBB$GAP_DNA_Sample=="Yes"]


N_TotalStudies <- PennSleepMetadata_wGAP_wPMBB %>%
        group_by(SleepStudy_MRN) %>%
        summarise(N_Studies=n())

SleepStudy_GAP_PMBB_Summary <- left_join(SleepStudy_GAP_PMBB_Summary, N_TotalStudies)

saveRDS(SleepStudy_GAP_PMBB_Summary, "SleepStudy_GAP_PMBB_Summary.Rdata")

#### Create Counts (Venn Diagram)
pdf("VennDiagram_GAP_PMBB.pdf", height = 5)
# GAP ID available / PMBB consented
vennDiagram(vennCounts(SleepStudy_GAP_PMBB_Summary[,c(2,4)]), main="Consented PMBB and GAP ID available")
# GAP DNA available / PMBB genotyped
vennDiagram(vennCounts(SleepStudy_GAP_PMBB_Summary[,c(3,5)]), main="Geneotyped PMBB and GAP DNA sample available")
dev.off()


##### PMBB genotyped: 2069
##### GAP ID DNA available: 1343
##### Intersection: 3300 (unique individuals with potential of having genotyped data and PSG data extracted)

# Get list of GAP already genotyped
GAP_wDNA <- SleepStudy_GAP_PMBB_Summary %>%
        filter(GAP_withDNA)

GAP_wDNA <- left_join(GAP_wDNA, PennSleepMetadata_wGAP_wPMBB[, c("SleepStudy_MRN", "GAP_ID")]) %>% distinct()

GAP_wDNA_deid <- select(GAP_wDNA, GAP_ID, PMBB_consented, PMBB_genotyped)
write.csv(GAP_wDNA_deid, "PennSleepDatabase_GAP_extracted_PMBB.csv", row.names = F)








