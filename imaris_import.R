library(readxl)
library(stringr)

## Input to be changed when the file changes ##

# Read in the excel file
rootpath = 'C:/Data folder/akvile'
filename = "Imaris_output.xls"
excel_file = file.path(rootpath, filename)

## functions for later use ##

decorate_surface = function(number) {
  return(paste("Surface", as.character(number)))
}

## end of function ##


sheetnames = excel_sheets(excel_file)

#find the sheetnames associated with Shortest distance
sheetnames_distance = sheetnames[str_detect(sheetnames, regex("shortest", ignore_case = T))]

#find the numbers for the all the different surfaces
temp_data = read_excel(excel_file, sheet = sheetnames[str_detect(sheetnames, regex("shortest", ignore_case = T))][1], skip = 1)
surfaces_all = as.numeric(str_extract(temp_data$`Surpass Object`, "\\d+"))
surface_current = as.numeric(str_extract(temp_data$Surfaces[1],"\\d+"))
surfaces_all = sort(append(surfaces_all,surface_current))
surfaces_number = length(surfaces_all)

#prepare the dataframe 
ShortestDistance = as.data.frame(setNames(replicate(surfaces_number,numeric(surfaces_number), simplify = F), lapply(surfaces_all,decorate_surface)))
rownames(ShortestDistance) = surfaces_all

#create the loop which will load in the data
for (currentsheet in sheetnames[str_detect(sheetnames, regex("shortest", ignore_case = T))]){
  temp_data = read_excel(excel_file, currentsheet, skip = 1)
  #get the current surface and the list of surfaces without
  #these are used to index and identify where the insert of the diagonal
  #is supposed to go
  surface_current = as.numeric(str_extract(temp_data$Surfaces[1],"\\d+"))
  surfaces_other = as.numeric(str_extract(temp_data$`Surpass Object`, "\\d+"))
  distances = temp_data$`Shortest Distance to Surfaces`
  distances = c(distances[surfaces_other < surface_current], NaN, distances[surfaces_other > surface_current])
  ShortestDistance[surfaces_all == surface_current] = distances
  

}
