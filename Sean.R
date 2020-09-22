# ===== export R PATH before running the script; Skip if it has been done====
# open cmd mode and enter set PATH=%PATH%;C:\Program Files\R\R-3.6.3\bin (change the path if needed)
# === Packages ====
library(readxl)
library(xlsx)
library(rapportools)
# === setup Working directory ====
dirPath<-NULL
cat("Please enter the path? or press 'Enter' for current working director\n")
dirPath <- readLines(file("stdin"), n = 1L)
if (is.empty(dirPath)){
  dirPath <- getwd()
} else{
  dirPath <-gsub ('\\\\','/',dirPath)
}
setwd(dirPath)
# === Getting data from files ====
xlsx.files<-list.files(dirPath,pattern = "xlsx$")

for (f in 1:length(xlsx.files)){
  # read data
  xlsx.file.table<-data.frame(read_excel(xlsx.files[f], col_types="text", range=cell_cols(c("H","W")), col_names = TRUE ,trim_ws = TRUE, sheet = 1))
  
  # remove unwanted data
  xlsx.file.table<-xlsx.file.table[,c(1,2,5:16)]
  # create a blank data
  sheet_names<-c('Spot_Area','Region_Intensity','Objects_Numbers', 'Nucleus_Area', 'Nucleus_Roundness', 'Nucleus_Ratio', 'YFP_Range', 'Spots_Area',
                 'Intensity_Nucleus_YFP','YFP_Intensity_cutoff','pos_cells', 'spot_area_cutoff')
  
  for (sheet_name in sheet_names){
    assign(sheet_name, data.frame())
  }
  
  # add data into blank data
  for (n in 1:length(sheet_names)){
    for (m in 1:nrow(xlsx.file.table)){
      x<-xlsx.file.table[m,1]
      y<-xlsx.file.table[m,2]
      write(paste0(sheet_names[n],'[x,y]<-as.numeric(xlsx.file.table[m,(n+2)])'), 'temp.R')
      source('temp.R', local = TRUE)
      file.remove('temp.R')
    }
    # change the colname and rowname
    write(paste0('colnames(',sheet_names[n],')<-3:22'), 'temp.R')
    write(paste0('row.names(',sheet_names[n],')<-LETTERS[2:15]'), 'temp.R', append = TRUE)
    source('temp.R', local = TRUE)
    file.remove('temp.R')
  }
  
  # export data to the original xlsx file
  for (n in 1:length(sheet_names)){
    write(paste0('write.xlsx(',sheet_names[n], ',file=xlsx.files[',f,'],', 'sheetName=sheet_names[',n,'], row.names=TRUE, col.names = TRUE, append = TRUE)'),'tmp.R')
    write(paste0('print(paste(xlsx.files[',f,'],sheet_names[',n,']))'),'tmp.R', append = TRUE)
    source('tmp.R', local = TRUE)
    file.remove('tmp.R')
  }
}



