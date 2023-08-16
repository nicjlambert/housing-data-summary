#'@description This code downloads an Excel file from a URL, reads the file, and 
#'creates a list of data frames for each sheet in the file. It then filters the 
#'list of data frames to include only sheets whose names contain the word "Table". 
#'The code assigns names to each data frame in the list and slices each data frame 
#'to remove unnecessary rows. It then uses pivot_longer() to unpivot the data frames 
#'and store the result in a new list. The new list of data frames are then bound
#'together into one data frame and written to CSV.

CRAN <- "https://cran.csiro.au"

packages <- c("readxl", "tidyr", "httr", "dplyr", "xlsx")

# Check if packages are installed and install them if necessary
lapply(packages, function(x) {
  if (!require(x, character.only = TRUE)) {
    install.packages(x, dependencies = TRUE, repos = CRAN)
    library(x, character.only = TRUE)
  }
})

download_file <- function(url) {
  tf <- tempfile()
  httr::GET(url, write_disk(tf, overwrite = TRUE))
  return(tf)
}

read_and_filter_sheets <- function(file_path) {
  sheets <- excel_sheets(file_path)
  lapply(sheets, read_excel, path = file_path)[grepl("Table", sheets)]
}

trim_data <- function(data_frames) {
  lapply(data_frames, function(df) {
    edges <- grep("Total", df)
    slice(df, tail(edges, 1):head(edges, 1))
  })
}

clean_and_unpivot <- function(trimmed_dfs, headers) {
  unpivoted_dfs <- lapply(seq_along(trimmed_dfs), function(i) {
    df <- trimmed_dfs[[i]][!is.na(trimmed_dfs[[i]][, 1]), ]
    colnames(df) <- headers[[i]]
    pivot_longer(df, cols = -1, names_to = "Second_Attribute", values_to = "Value") %>%
      mutate(Table_Name = names(headers[i])) %>%
      select(Table_Name, everything())
  })
  do.call(rbind, unpivoted_dfs)
}

url <- "https://www.abs.gov.au/statistics/people/housing/housing-census/2021/Housing%20data%20summary.xlsx"

headers <- list(
  "Table 1. Dwelling Type by State and Territory" = c("Dwelling Type","New South Wales","Victoria","Queensland","South Australia","Western Australia","Tasmania","Northern Territory","Australian Capital Territory", "Total")
  , "Table 2. Dwelling Location by Dwelling Structure" = c("Dwelling Location","Separate house","Semi-detached, row or terrace house, townhouse etc.","Flat or apartment","Caravan","Cabin, houseboat","Improvised home, tent, sleepers out", "House or flat attached to a shop, office, etc.", "Not stated", "Total")
  , "Table 3. Dwelling Structure by Tenure and Landlord Type" = c("Dwelling Structure","Owned outright","Owned with a mortgage","Rented: Real estate agent","Rented: State or territory housing authority","Rented: Community housing provider", "Rented: Person not in same household",	"Rented: Other landlord type","Rented: Landlord type not stated", "Other tenure type", "Tenure type not stated", "Total")
  , "Table 4. Dwelling Structure by Number of Bedrooms" = c("Dwelling Structure","None (includes studio apartments or bedsitters)","One bedroom","Two bedrooms","Three bedrooms","Four bedrooms or more","Not stated","Total")
  , "Table 5. Dwelling Structure by Household Composition" = c("Dwelling Structure","One family households", "Multiple family households","Total family households","Group houeholds","Lone person households","Total person in occupied private dwellings")
  , "Table 6. Tenure Type by Total Household Income (Weekly)" = c("Tenure Type","Negative/Nil income","$1-$149","$150-$299","$300-$399","$400-$499","$500-$649","$650-$799","$800-$999","$1,000-$1,249","$1,250-$1,499","$1,500-$1,749","$1,750-$1,999","$2,000-$2,499","$2,000-$2,999","$3,000-$3,499","$3,500-$3,999","4000	or more","Not stated","Total")
  , "Table 7. Mortgage Repayment (Monthly) by Dwelling Structure" = c("Mortgage Repayment (Weekly)","Separate house","Semi-detached, row or terrace house, townhouse etc.","Flat or apartment","Caravan", "Cabin, houseboat", "Improvised home, tent, sleepers out","House or flat attached to a shop, office, etc.","Not stated","Total")
  , "Table 8. Rent (Weekly) by Dwelling Structure" = c("Rent (Weekly)","Separate house","Semi-detached, row or terrace house, townhouse etc.","Flat or apartment","Caravan","Cabin, houseboat", "Improvised home, tent, sleepers out","House or flat attached to a shop, office, etc.","Not stated","Total")
)

file_path <- download_file(url)
filtered_data_frames <- read_and_filter_sheets(file_path)
trimmed_data_frames <- trim_data(filtered_data_frames)
final_data <- clean_and_unpivot(trimmed_data_frames, headers)
final_data$First_Attribute <- gsub("\\([a-z]\\)", "", final_data$First_Attribute)
write.xlsx(final_data, file = "Housing data summary.xlsx")
unlink(file_path)
