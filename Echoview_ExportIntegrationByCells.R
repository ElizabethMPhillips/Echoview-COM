### Created by Chelsea Stanley, DFO
### Modifed by Beth Phillips, NOAA-NWFSC

### Uses COM objects to open Echoview and export echogram -integration by cells

###############################################
#----------------  INPUT  --------------------#
###############################################

# acoustic variables(s) to integrate and their frequency
variables <- c("EUPH export 18kHz","EUPH export 38kHz","EUPH export 70kHz","EUPH export 120kHz","EUPH export 200kHz")
frequency <- c("18","38","70","120","200")

# required packages
require(RDCOMClient)
require(dplyr)
require(stringr)

# set the working directory
#setwd('..'); setwd('..')
BaseYearPath<-"E:/"
BaseJudgePath<-"E:/"
BaseProjPath<-"E:/"
setwd(BaseJudgePath)

gdepth=10 # depth for grid cells, in meters (e.g. 1 m)

# Set location of EV files

EUPH_EV <- "" # Directory in BaseJudgePath
EVdir<- ""

# Where to put exports

EUPH_exportbase <- "" #Directory to save .csv exports 
dir.create(file.path(BaseYearPath, EUPH_exportbase))
orig_exportdir<-"1_OriginalExports_surfaceto300m" #Example of how to track exports and range settings
EUPH_export<-file.path(EUPH_exportbase, orig_exportdir)
dir.create(file.path(BaseYearPath, EUPH_export))

# bind variable and frequency together to create separate folders for each variable export (for multi-frequency exports)
vars <- data.frame(variables,frequency, stringsAsFactors = FALSE)
        for(v in 1:nrow(vars)){
          var <- vars$variables[v]
          dir.create(file.path(BaseYearPath, EUPH_export,var))
        }

#list the EV files to integrate
EVfile.list <- list.files(file.path(BaseJudgePath, EUPH_EV), pattern = ".EV")

# create folder in EUPH_export folder for each variable
for(f in variables){
  dir.create(file.path(BaseYearPath, EUPH_export, f))
}

# Loop through EV files 

for (i in EVfile.list){
  # create COM connection between R and Echoview
  EVApp <- COMCreate("EchoviewCom.EvApplication")
  EVApp$Minimize()  #keep window open in background
  
  # EV filenames to open
  EVfileNames <- file.path(getwd(), EUPH_EV, i)
  EvName <- strsplit(i, split = '*.EV')[[1]]
  
  # open EV file
  EVfile <- EVApp$OpenFile(EVfileNames)
  EVfileName <- file.path(getwd(),EUPH_EV, i)
  print(EVfileName)
  print(i)
  # Variables object
  Obj <- EVfile[["Variables"]]
  
  # loop through variables for integration
  for(v in 1:nrow(vars)){
    var <- vars$variables[v]
    freq <- vars$frequency[v]
    varac <- Obj$FindByName(var)$AsVariableAcoustic()
    
    # Set analysis lines
    Obj_propA<-varac[['Properties']][['Analysis']]
    Obj_propA[['ExcludeAboveLine']]<-"14 m from surface"
    Obj_propA[['ExcludeBelowLine']]<-"Final bottom" 
    
    # Set analysis grid and exclude lines on Sv data
    Obj_propGrid <- varac[['Properties']][['Grid']]
    Obj_propGrid$SetDepthRangeGrid(1, gdepth)
    Obj_propGrid$SetTimeDistanceGrid(3, 0.5)
 
    # export by cells, with Echoview file, frequency, and cell resolution (0.5 nmi x 10 m) in filename
    exportcells <- file.path(BaseYearPath, EUPH_export, var, paste(EvName, freq, "0.5nmi_10m_cells.csv", sep="_"))
    varac$ExportIntegrationByCellsAll(exportcells)

    # Set analysis grid and exclude lines on Sv data back to original values
    Obj_propGrid<-varac[['Properties']][['Grid']]
    Obj_propGrid$SetDepthRangeGrid(1, 50)
    Obj_propGrid$SetTimeDistanceGrid(3, 0.5)
    }

  # save EV file
  EVfile$Save()

  #close EV file
  EVApp$CloseFile(EVfile)

  #quit echoview
  EVApp$Quit()


## ------------- end loop

}



