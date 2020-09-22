# 384well_Sean
This script is only for converting a 384-well data sheet into several individual data sheets with R 
# essential package
1. Xmisc
2. xlsx
3. readxl

# Before running the script, please make sure 'Rscript' can be run in cmd mode.
If you do know how to creat the path in your cmd environemnt, please look up the link below. 
https://www.architectryan.com/2018/08/31/how-to-change-environment-variables-on-windows-10/ (For windows 10)

# How to use
1. open the 'cmd'
2. go to the directory containing the script.
3. run the script by typing 'Rscript Sean.R'
The script will loop through all xlsx files in the current directory.

For example:
c:\Sean\Rscript Sean.R

For selecting xlsx files in another directory, 
c:\Sean\Rscript Sean.R --dir=PATH  
PATH= the path of the directory

For help,
c:\Sean\Rscript Sean.R -h
