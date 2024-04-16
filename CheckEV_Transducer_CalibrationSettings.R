### Uses COM objects to run Echoview and get transducer and calibration settings

# required packages 
require(RDCOMClient)
require(DescTools)
require(data.table)

## Location of EV files to check transducer settings ####
EVdir <- as.character("N:/Survey.Acoustics/2023 Hake Sum SH_JF/Judging/Euphausiids/copy of FINAL for biomass")

expected_ducerdepth <- 9.15 #expected transducer vertical (z-) offset for NOAA Ship Shimada

# If interested in only the 38 kHz and 120 kHz transducers, should be T2 and T4 for most Shimada datasets
vars <- c("Sv pings T2", "Sv pings T4") #T2 should be 38 kHz, and T4 should be 120 kHz

# list the EV files to check
EVfile.list <- list.files(paste0(EVdir), pattern=".EV$", ignore.case = TRUE)

#####################################################################
#   Open each EV file to get transducer and calibration settings  #
####################################################################

# Setup empty lists to store the settings
ducer_location.list = list()
cal_settings.list = list() 
EV_settings.list = list()

#i="x1.EV"
#rm(i)

for (i in EVfile.list){
  # EV filename
  name <- tools::file_path_sans_ext(i)[[1]]
  EVfileName <- file.path(paste0(EVdir), i)
  print(EVfileName)
  EvName <- strsplit(i, split = '*.EV')[[1]]

  # create COM connection between R and Echoview
  EVApp <- COMCreate("EchoviewCom.EvApplication")
  EVApp$Minimize()  #Minimize EV file to run in background
  
  # open EV file
  EVfile <- EVApp$OpenFile(EVfileName)
  
### access transducer properties
  transducerobj <- EVfile[["Transducers"]]
  no_trans <- transducerobj[["Count"]] #get number of transducers
  no_trans_list <-vector(mode="list", length=no_trans)
  
  #get transducer depth (vertical offset), and x- y- offset settings
  
  for (k in 1:length(no_trans_list)){
    # get offset settings for each transducer
    transducerobj_props <- EVfile[["Transducers"]]$Item(k-1)
    Freq <- transducerobj_props[["Name"]]
    x_offset <- transducerobj_props[["AlongshipOffset"]] 
    y_offset <- transducerobj_props[["AthwartshipOffset"]] 
    z_offset <- transducerobj_props[["VerticalOffset"]] 
    flag <- fifelse(z_offset==expected_ducerdepth,"N","Y")
    ducer_location.list[[k]] <- c(name,Freq,x_offset,y_offset,z_offset,flag)
  }    
    #combine settings for each transducer
    ducer_location = do.call(rbind, ducer_location.list)
    #change column names
    colnames(ducer_location) <- c("tran","freq","x","y","z","flag") 
    
    #subset just "38 kHz" and "120 kHz" settings
    ducer_location <- as.data.frame(ducer_location) #needs to be a dataframe for subset call
    ducer_location <- subset(ducer_location, ducer_location$freq == "38 kHz" | ducer_location$freq == "120 kHz")
    as.matrix(ducer_location) #back to matrix just to be safe, to merge with other list (don't think this matters...)
   
### access T2 and T4 variable properties
    Obj <- EVfile[["Variables"]]

    #get calibration values of interest for T2 and T4
    
    for(j in vars){
    varac <- Obj$FindByName(j)$AsVariableAcoustic()
    varname <- j
    excl_below<-varac[["Properties"]][["Analysis"]][["Excludebelow"]]
    baddata<-varac[["Properties"]][["Analysis"]][["BadDataHasVolume"]]
    TDmode<-varac[["Properties"]][["Grid"]][["TimeDistanceMode"]]
      if(TDmode==2){TDmodech<-"GPSNMi"}
      if(TDmode==3){TDmodech<-"VesselLogNMi"}
    varaccal<-varac[["Properties"]][["Calibration"]]
    SoundSpeed<-varaccal$Get("SoundSpeed",1)
    TSGain<-varaccal$Get("Ek5TsGain",1)
    SaCorr<-varaccal$Get("EK60SaCorrection",1)
    cal_settings.list[[j]] <- c(varname,excl_below,baddata,TDmodech,SoundSpeed,TSGain,SaCorr)
  }
    #rbind settings for each transducer
    cal_settings = do.call(rbind, cal_settings.list)
    #rename rows and columns
    colnames(cal_settings) <- c("Variable","Exclude_below_line","Include_volume_nodata_samples","TimeDistanceMode","SoundSpeed", "TsGain", "SaCorrection") 
    rownames(cal_settings) <- c(2,4)
    cal_settings <- as.data.frame(cal_settings)

  #combine transducer and calibration settings into one data frame
  EV_settings = cbind(ducer_location, cal_settings)
  
  #Change column names of final data frame
  names(EV_settings)[1] <- "Transect"
  names(EV_settings)[2] <- "Transducer name"
  names(EV_settings)[3] <- "x_offset"
  names(EV_settings)[4] <- "y_offset"
  names(EV_settings)[5] <- "z_offset"
  names(EV_settings)[6] <- "Flag_verticaloffset"
  names(EV_settings)[7] <- "Variable name"
  names(EV_settings)[8] <-  "Exclude_below_line"
  names(EV_settings)[9] <- "Include_volume_nodata_samples"
  names(EV_settings)[10] <- "TimeDistanceMode"
  names(EV_settings)[11] <- "SoundSpeed"
  names(EV_settings)[12] <- "TsGain"
  names(EV_settings)[13] <- "SaCorrection"
  EV_settings
  
  #add settings for transect to rest of the transects
  EV_settings.list[[i]] <- EV_settings 
  
  # Bind rows from all transects into a data frame
  EV_Trans_Cal_Settings = as.data.frame(do.call(rbind, EV_settings.list))
  
# Close EV file
EVApp$CloseFile(EVfile)

# Quit echoview
EVApp$Quit()

} 

# Save final dataframe as .csv
write.csv(EV_Trans_Cal_Settings, paste0(EVdir,"/TransducerCalibrationCheck.csv"),row.names = FALSE)
