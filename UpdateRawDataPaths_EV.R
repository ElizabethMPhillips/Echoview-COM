
## This script will update directory paths in EV files to network or local drive, to make analyses more efficient
#### THINGS TO CHANGE ####

year<-2021
loc<-"US" #Should be US or CAN for different datasets
surveyName <- paste(year,loc,sep="_") #for input into subset call below

networkdir <- "N:/Survey.Acoustics/2024 HakeRes Sum SH/Data_SH/Acoustics/EK80_raw/Leg 2"
localdir <- "D:/Survey.Acoustics/2024 HakeRes Sum SH/Data_SH/Acoustics/EK80_raw/Leg 2"
dir.create(localdir, recursive = TRUE)
judgedir <- "N:/Survey.Acoustics/2024 HakeRes Sum SH/Judging/Leg 2/Judging-EV-Files"
localjudgedir <- "D:/Survey.Acoustics/2024 HakeRes Sum SH/Judging/Leg 2/Judging-EV-Files"
dir.create(localjudgedir, recursive = TRUE)
EVfile.list <- list.files(judgedir, pattern=".EV$", ignore.case = TRUE)

# Copy files to local drive
file.copy(from = paste0(judgedir,"/", EVfile.list),
          to = paste0(localjudgedir, "/", EVfile.list))

###############################################
#              Run in Echoview                #
###############################################

require(RDCOMClient)
# create COM connection between R and Echoview
EVApp <- COMCreate("EchoviewCom.EvApplication")

#rm(i)
#i = "02august2024-MG1-Pass1-1.1.EV"          

for (i in EVfile.list){
  EVfileName <- file.path(localjudgedir, i)
  print(EVfileName)
  # open EV file
  EVfile <- EVApp$OpenFile(EVfileName)
  
  #### Remove directory raw files, add local raw files

  filesetObj <- EVfile[["Filesets"]]$Item(0)  #RAW
  num <- filesetObj[["DataFiles"]]$Count()
  
  raws <- NULL
  for (l in 0:(num-1)){
    dataObj <- filesetObj[["DataFiles"]]$Item(l)
    dataPath <- dataObj$FileName()
    dataName <- sub(".*\\\\|.*/","",dataPath)
    raws <- c(raws,dataName) 
  }
  
  filesetObj[["DataFiles"]]$RemoveAll() #need to remove first, then add other ones back in
  
  # Add correct raw files
  for (r in raws){
    filesetObj[["DataFiles"]]$Add(file.path(localdir,r))
  }
  
   # save EV file
  EVfile$Save()
  #close EV file
  EVApp$CloseFile(EVfile)
  }
#quit echoview
EVApp$Quit()

