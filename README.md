# M2L_FullForecastModel

# Within the process_data.R file, the mappings for:
# -KeyTier1Other
# -Solicitor Type 
# -Product Name Clean
# Should be checked every year to ensure that the filters are valid and effective

# Last checked date:
# 25-09-2023

# There are two initial components of this R Project
# load_package.R -
# This installs or loads all the necessary packages for the code to be run
# process_data.R -
# This formats the daily extract data frame into the necessary format for any 
further computations

# Seperating out the functions for loading packages and formatting data means, 
# that updating the intialisation scripts will allow us to perform the updates 
# on all other scripts that call them in the project.