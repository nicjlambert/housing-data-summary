#'@author nicholas.lambert
#'@description This code downloads an Excel file from a URL, reads the file, and 
#'creates a list of data frames for each sheet in the file. It then filters the 
#'list of data frames to include only sheets whose names contain the word "Table". 
#'The code assigns names to each data frame in the list and slices each data frame 
#'to remove unnecessary rows. It then uses pivot_longer() to unpivot the data frames 
#'and store the result in a new list. The new list of data frames are then bound
#'together into one data frame and written to CSV.

# Set up workspace --------------------------------------------------------

# Check if packages are installed and install them if necessary
# Set the CRAN mirror to use
CRAN <- "https://cran.csiro.au"

# Define the packages to check
packages <- c("readxl", "tidyr", "httr", "dplyr", "xlsx")

# Check if packages are installed and install them if necessary
lapply(packages, function(x) {
  if (!require(x, character.only = TRUE)) {
    install.packages(x, dependencies = TRUE, repos = CRAN)
    library(x, character.only = TRUE)
  }
})

# # Set proxy configuration if necessary
config_proxy <- use_proxy(
  url = curl::ie_get_proxy_for_url(),
  auth = "ntlm",
  username = ""
)

# Headers
header <- list(
  "Table 1. Dwelling Type by State and Territory" = c("Dwelling Type","New South Wales","Victoria","Queensland","South Australia","Western Australia","Tasmania","Northern Territory","Australian Capital Territory", "Total")
  , "Table 2. Dwelling Location by Dwelling Structure" = c("Dwelling Location","Separate house","Semi-detached, row or terrace house, townhouse etc.","Flat or apartment","Caravan","Cabin, houseboat","Improvised home, tent, sleepers out", "House or flat attached to a shop, office, etc.", "Not stated", "Total")
  , "Table 3. Dwelling Structure by Tenure and Landlord Type" = c("Dwelling Structure","Owned outright","Owned with a mortgage","Rented: Real estate agent","Rented: State or territory housing authority","Rented: Community housing provider", "Rented: Person not in same household",	"Rented: Other landlord type","Rented: Landlord type not stated", "Other tenure type", "Tenure type not stated", "Total")
  , "Table 4. Dwelling Structure by Number of Bedrooms" = c("Dwelling Structure","None (includes studio apartments or bedsitters)","One bedroom","Two bedrooms","Three bedrooms","Four bedrooms or more","Not stated","Total")
  , "Table 5. Dwelling Structure by Household Composition" = c("Dwelling Structure","One family households", "Multiple family households","Total family households","Group houeholds","Lone person households","Total person in occupied private dwellings")
  , "Table 6. Tenure Type by Total Household Income (Weekly)" = c("Tenure Type","Negative/Nil income","$1-$149","$150-$299","$300-$399","$400-$499","$500-$649","$650-$799","$800-$999","$1,000-$1,249","$1,250-$1,499","$1,500-$1,749","$1,750-$1,999","$2,000-$2,499","$2,000-$2,999","$3,000-$3,499","$3,500-$3,999","4000	or more","Not stated","Total")
  , "Table 7. Mortgage Repayment (Monthly) by Dwelling Structure" = c("Mortgage Repayment (Weekly)","Separate house","Semi-detached, row or terrace house, townhouse etc.","Flat or apartment","Caravan", "Cabin, houseboat", "Improvised home, tent, sleepers out","House or flat attached to a shop, office, etc.","Not stated","Total")
  , "Table 8. Rent (Weekly) by Dwelling Structure" = c("Rent (Weekly)","Separate house","Semi-detached, row or terrace house, townhouse etc.","Flat or apartment","Caravan","Cabin, houseboat", "Improvised home, tent, sleepers out","House or flat attached to a shop, office, etc.","Not stated","Total")
)

# Define the URL of the .xlsx file and the path to where you want to save it on your local machine
url <- "https://www.abs.gov.au/statistics/people/housing/housing-census/2021/Housing%20data%20summary.xlsx"

# Create a temporary file
tf <- tempfile()

# Download the .xlsx file from the URL
httr::GET(url, write_disk(tf, overwrite = TRUE), config = config_proxy)

# Read the .xlsx file and create a list of data frames for each sheet in the file
sheets <- excel_sheets(tf)

# Create a list of data frames containing only the sheets whose names contain the word "Table"
filtered_data_frames <- lapply(sheets, read_excel, path = tf)[grepl("Table", sheets)]

# Name each object in the list_of_data_frame so we know which dataframe is which
names(filtered_data_frames) <- sheets[grepl("Table", sheets)]

# Initialize the second list with the appropriate number of elements
lists <- c("indices","head_element", "tail_element")

# Assign a vector of lists with the appropriate number of elements to the variable with the specified name
for(list in lists){
  assign(list, vector("list", length = length(filtered_data_frames)))
}

# Loop over the data frames in the list, slice each one, and add the sliced data frames to the new list
for (i in seq_along(filtered_data_frames)) {
  
  for (j in 1:ncol(filtered_data_frames[[i]])) {
    
    # Find the edges of each Table
    # Use grep() to find the indices of the words "Total"
    edges <- grep("Total", filtered_data_frames[[i]][[j]])[1]
    
    indices[[i]][[j]] <- edges
    
  }
  
  indices[[i]] <- Filter(function(x) {
    !is.na(x) == TRUE
  }, indices[[i]])
  
  
  # Select the first element of the list
  head_element[[i]] <- head(indices[[i]], 1)
  
  # Select the last element of the list
  tail_element[[i]] <- tail(indices[[i]], 1)
  
}

# Initialize an empty list to store the trimmed data frames
trim_data_frames <- list()

# Loop over the data frames in the list, slice each one, and add the sliced data frames to the new list
for (i in seq_along(filtered_data_frames)) {
  
  trim_data_frames[[i]] <- slice(
    filtered_data_frames[[i]], tail_element[[i]][[1]]:head_element[[i]][[1]]
  )
  
}

# Select the first 8 columns from the 4th and 5th element of the trim_data_frames list
trim_data_frames[[4]] <- trim_data_frames[[4]][, 1:8]
trim_data_frames[[5]] <- trim_data_frames[[5]][-1, 1:7]

# Use lapply() to create a list of logical vectors indicating which rows in the first column of each data frame contain missing values
rows_to_exclude <- lapply(trim_data_frames, function(x) is.na(x[, 1]))

# Initialize an empty list to store the cleaned data frames
cleaned_data_frames <- list()

# Loop over the data frames in the list
for (i in seq_along(trim_data_frames)) {
  
  # Use the is.na() function to create a logical vector indicating which rows in the first column contain missing values
  rows_to_exclude <- is.na(trim_data_frames[[i]][, 1])
  
  # Use the [ operator to exclude the rows with missing values from the current data frame
  cleaned_data_frames[[i]] <- trim_data_frames[[i]][!rows_to_exclude, ]
  
}

# Use the [ operator to exclude the first row from the 6th data frame
cleaned_data_frames[[6]] <- cleaned_data_frames[[6]][2:nrow(cleaned_data_frames[[6]]), ]

for (i in seq_along(cleaned_data_frames)) {
  
  names(cleaned_data_frames[[i]]) <- header[[i]]
  
}

unpivoted_dfs <-  list()

for (i in seq_along(cleaned_data_frames) ) {
  
  colnames(cleaned_data_frames[[i]])[1] <- "First_Attribute"
  
  # Unpivot the data frame using pivot_longer()
  unpivoted_df <- 
    pivot_longer(cleaned_data_frames[[i]],
                 cols = -c(names(cleaned_data_frames[[i]][,1])),
                 names_to = "Second_Attribute",
                 values_to = "Value") %>%
    mutate(Table_Name = names(header[i])) %>% 
    select(Table_Name, First_Attribute, Second_Attribute, Value)
  
  unpivoted_dfs[[i]] <- unpivoted_df
}

# Bind all list of dfs together into one dataframe
df_data <- do.call(rbind, unpivoted_dfs)

# Use the gsub() function to remove the string '(a)', '(b)', or '(c)'
df_data$First_Attribute <- gsub("\\([a-z]\\)", "", df_data$First_Attribute)

# Write the df_data to a file
write.xlsx(
   df_data,
   file = paste0("Housing data summary.xlsx")
)

# Delete the temporary file
unlink(tf)
