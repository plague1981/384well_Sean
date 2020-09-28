# ===== export R PATH before running the script; Skip if it has been done====
# open cmd mode and enter set PATH=%PATH%;C:\Program Files\R\R-3.6.3\bin (change the path if needed)
# === Packages ====
if("Xmisc" %in% rownames(installed.packages()) == TRUE) {
  library(Xmisc)
} else install.packages(Xmisc)
if (check.packages(readxl)){
  library(readxl)
} else install.packages(readxl)
if (check.packages(xlsx)){
  library(xlsx)
} else install.packages(xlsx)
require(methods)
# === setting environment ===
parser <- ArgumentParser$new()
parser$add_usage('Sean.R [options]')
parser$add_description('An executable R script parsing arguments from Unix-like command line.')
parser$add_argument('--h',type='logical', action='store_true', help='Print the help page')
parser$add_argument('--help',type='logical',action='store_true',help='Print the help page')
parser$add_argument('--dir', type = 'character', default = getwd(), help = 'Enter the files existing directory')
parser$helpme()
# === Working directory ====
dirPath <- dir
if (dir.exists(dirPath)){
  setwd(dirPath)
}


# Trying creating a 'results' folder for upcoming excel files.
tryCatch(
  expr = {
    dir.create(file.path(getwd(), 'Results'))
    message("Successfully created 'results' folder." )
  },
  error = function(e){
    message('Caught an error!')
    print(e)
  },
  warning = function(w){
    message('Caught an warning!')
    print(w)
  },
  finally = {
    message('All set')
  }
)

# Set export directory
export_dir<-paste0(dirPath,'/Results')
# check the files created
open_folder <-function(dir){
  if (.Platform['OS.type'] == "windows"){
    shell.exec(dir)  
  } else {
    system(paste(Sys.getenv("R_BROWSER"), dir))
  }
}
# Call the function to open the folder
open_folder(export_dir)

# === Getting data from files ====
xlsx.files<-list.files(dirPath,pattern = "xlsx$")

for (f in 1:length(xlsx.files)){
  # read data
  xlsx.file.table<-data.frame(read_excel(xlsx.files[f], col_types="text", range=cell_cols(c("H","AF")), col_names = TRUE ,trim_ws = TRUE, sheet = 1))

  # remove unwanted data
  xlsx.file.table<-xlsx.file.table[,c(1,2,5:25)]
  # create a blank data
  sheet_names<-paste0('sheet_',(seq(1:21)))

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
    write(paste0('colnames(',sheet_names[n],')<-1:24'), 'temp.R')
    write(paste0('row.names(',sheet_names[n],')<-LETTERS[1:16]'), 'temp.R', append = TRUE)
    source('temp.R', local = TRUE)
    file.remove('temp.R')
  }

  # export data (xlsx files) into the "Results" fold
  setwd(export_dir)
  for (n in 1:length(sheet_names)){
    
    write(paste0('write.xlsx(',sheet_names[n], ',file=xlsx.files[',f,'],', 'sheetName=sheet_names[',n,'], row.names=TRUE, col.names = TRUE, append = TRUE)'),'tmp.R')
    source('tmp.R', local = TRUE)
    file.remove('tmp.R')
  }
  setwd(dirPath)
}
