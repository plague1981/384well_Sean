# 384well_Sean
This script is only for converting a 384-well data sheet into several individual data sheets with R 
# Essential packages
1. Xmisc
2. xlsx
3. readxl

# Before running the script, please make sure 'Rscript' can be run in cmd mode.
If you do not know how to creat the path in your cmd environemnt, please look up the link below. \
https://www.architectryan.com/2018/08/31/how-to-change-environment-variables-on-windows-10/ (For windows 10) \
or type below in the cmd \
PATH=%PATH%;C:\Program Files\R\R-3.6.3\bin (change the path if needed)
# How to use
1. download the script
2. open the 'cmd'
3. go to the directory containing the script.
4. run the script by typing 'Rscript 384wells.R'
The script will loop through all xlsx files in the current or seleted directory.

For example, \
c:\username\Rscript 384wells.R

For selecting xlsx files in another directory, \
c:\username\Rscript 384wells.R --dir=PATH  
PATH= the path of the directory

For help, \
c:\username\Rscript 384wells.R -h
