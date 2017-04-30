#

# clear the stored variables
rm(list=ls())
# clear any previously checked files
unlink(paste(getwd(),"checked",sep="/"), recursive = TRUE)
# initialize the text body to be outputted in the end
txt <- c()

 
# import the XLSX libary to read in Excel fiels (both XLSX or XLS are compatible)
library(xlsx)

# import the functions we use
source(file="./lib/R/getSheetNames.R")
source(file="./lib/R/findMissingFiles.R")
source(file="./lib/R/keepNumericRows.R")
source(file="./lib/R/delimitOCR.R")
source(file="./lib/R/matchRowNames.R")
source(file="./lib/R/matchColumnValues.R")
source(file="./lib/R/findErrors.R")


# confirm that the folder names are correct
#hand_entered <- "hand_entered"
#compare_with <- "ocr_ch"
hand_entered <- readline("What is the name of the folder containing hand-entered tables?")
compare_with <- readline("What is the name of the folder containing OCR/correct tables?")
# get a list of sheet names and prepare the output txt file with checking results
dw1filename <- getSheetNames(hand_entered)$filename
dw2filename <- getSheetNames(compare_with)$filename


# store the filenames -- assuming the list of OCR'ed files is complete
txt.name <- paste("checked_stat")
txt <-cbind(txt, filenames = dw2filename)

# compare filenames from 2 to 1
missing_file_flag <- findMissingFiles(hand_entered_list = dw1filename, ocr_list = dw2filename)
txt <-cbind(txt, missing_file_flag)

# let's forget about the files with different names for now!
# the assumption here is that the hand_entered files constitute a subset of the ocr_files
same_dw1files <- getSheetNames(hand_entered)$fullpath
same_dw2files <- getSheetNames(compare_with)$fullpath[missing_file_flag == 0]

# we also prepare a data frame to store the error stat for each file that we will check
errors <- data.frame(filenames = getSheetNames(hand_entered)$filename,
                     num_errors = NA, 
                     num_checked = NA)

# loop over files in the directory
# if there were issues reading in the tables, then we could not check them
 for (ii in 1:length(same_dw1files)){

   # try reading in the OCR'ed table
   fail_read_ocr <- tryCatch(
     {
   dw2 <- read.xlsx(same_dw2files[ii],1,as.data.frame = T, header = F, encoding = "UTF-8")
     }
   ,
   error = function(e) e
   )
  if (is.null(dw2) == F) {
   
   if( !inherits(fail_read_ocr, "error") && (ncol(dw2) > 2 && nrow(dw2) > 2) ) {
     
     
    dw2 <- data.frame(lapply(dw2, as.character),stringsAsFactors = FALSE)
   
    dw1 <- read.xlsx(same_dw1files[ii],1,as.data.frame = T, header = F, encoding = "UTF-8")
    dw1 <- data.frame(lapply(dw1, as.character),stringsAsFactors = FALSE)

    
    dw1_num <- keepNumericRows(dw1)$numeric
    string <- keepNumericRows(dw1)$string
    dw2_num <- keepNumericRows(dw2)$numeric
    
    # delimit and clean the OCR files
    try(
{
  dw2_num <- delimitOCR(dw2_num)
  cat("...cleaned OCR'ed","table", getSheetNames(hand_entered)$filename[ii],"\n")
}
    )

similar_nrow <- (0.9 < nrow(dw1_num)/nrow(dw2_num) && nrow(dw1_num)/nrow(dw2_num) < 1.1)
similar_ncol <- (0.9 < ncol(dw1_num)/ncol(dw2_num) && ncol(dw1_num)/ncol(dw2_num) < 1.1)
similar <- (similar_nrow + similar_ncol) == 2
if ( similar == TRUE ) {
  
  # check if dimensions are the same
  tryCatch(
{
  test <- matchRowNames(hand_entered_data = dw1_num, ocr_data = dw2_num)
  dw1_num <- test$hand_entered_data
  dw2_num <- test$ocr_data
  
  
  # match on the column names -- more precisely the values in the first row, not column names
  
  test <- matchColumnValues(hand_entered_data = dw1_num, ocr_data = dw2_num)
  dw1_num <- test$hand_entered_data
  dw2_num <- test$ocr_data
}
  )

}
# at this point, if the dimensions agree, we can compare cell by cell, if not, skip to the next method
if ( nrow(dw1_num) == nrow(dw2_num) && ncol(dw1_num) == ncol(dw2_num) ) {
  
  cat("...dimensions are the same, comparing OCR'ed table and hand-entered table cell by cell", 
      "table", getSheetNames(hand_entered)$filename[ii],"\n")
  # now we are ready to find the errors in the hand_entered table and mark them for hand curating
  
  test <- findErrors(hand_entered_data = dw1_num, ocr_data = dw2_num)
  names(test)
  test$hand_entered_data_checked
  test$tot_checked
  test$tot_errors
  
  # append the rows of string back, but we might need to adjust the length of the string rows
  if (ncol(dw1_num) > ncol(string)) {
    string <- cbind(keepNumericRows(dw1)$string, xxx=rep(NA, ncol(dw1_num)-ncol(string)))
  }else{
    string <- keepNumericRows(dw1)$string[1:ncol(dw1_num)]
  }
  
  numeric <- test$hand_entered_data_checked
  colnames( numeric ) <- colnames( string ) 
  
  checked_dw1 <- rbind(string,numeric)
  
  unlink(paste(getwd(),paste(dw2filename[ii],"xlsx", sep = "."),sep="/"))

  write.xlsx(checked_dw1, paste(getwd(),paste(dw1filename[ii],"xlsx", sep = "."),sep="/"),
             col.names=F, row.names=F)
  

  
  # now we output the error stat to an output file with ID as filenames
  # note that we only looped the non-missing files, so we need to merge back to the complete file list
  
  errors$num_errors[ii] <- test$tot_errors
  errors$num_checked[ii] <- test$tot_checked
  
}else{
  # end for failing dimension checks
  cat("...dimensions are not the same, check if hand_entered values are in OCR'ed tables", 
      "table", getSheetNames(hand_entered)$filename[ii],"\n")
  
  dw2_unlist <- unlist(dw2_num)
  numeric <- dw1_num
  errors_counter <- 0
  for (rr in 1:nrow(dw1_num)) {
    diff <- !(dw1_num[rr,] %in% dw2_unlist)
    # sum over a row how many matched values
    errors_counter <- errors_counter+ sum(diff)
    
    # for non-matched valeus, mark the values with "***|"
    numeric[rr,][diff  == TRUE] <- paste("***",numeric[rr,][diff  == TRUE] ,sep="|")
    
  }
  checked_dw1 <- rbind(string,numeric)
  
    unlink(paste(getwd(),paste(dw2filename[ii],"xlsx", sep = "."),sep="/"))

  write.xlsx(checked_dw1, paste(getwd(),paste(dw1filename[ii],"xlsx", sep = "."),sep="/"),
             col.names=F, row.names=F)
  
  errors$num_errors[ii] <- errors_counter
  
  errors$num_checked[ii] <- nrow(dw1_num)*ncol(dw1_num)
  
  
  
}
# end for check cell by cell or any value
cat(">>>finished checking","table", getSheetNames(hand_entered)$filename[ii],"\n")

   }
}else{
  errors$num_errors[ii] <- NA
  errors$num_checked[ii] <- NA
  cat(">>>not able to read in the OCR'ed table. So I did not check" 
      ,"table", getSheetNames(hand_entered)$filename[ii],"\n")
  
}
# end for skipping a corrupted OCR table
}
# end for (ii in 1:length(same_dw1files))

# now merge in the stat on errors back to txt data frame with file names
txt <- merge(txt, errors, by="filenames", all = T )

outputCheck<-paste(getwd(),paste(txt.name,".txt",sep=""),sep="/")

write.table(txt, outputCheck, sep="\t", col.names = TRUE, row.names = FALSE)



