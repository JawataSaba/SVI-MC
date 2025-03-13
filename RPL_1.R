## Step 1: Package Installation and Library Loading
# Install required packages (if not already installed) and load libraries.
install.packages("quantmod")
install.packages("ggplot2")
install.packages("openxlsx")
install.packages("readxl")
install.packages("writexl")
install.packages("reshape2")

library(quantmod)
library(ggplot2)
library(openxlsx)
library(readxl)
library(writexl)
library(reshape2)

## Step 2: Set the Working Directory and Read Data
# Adjust the working directory to where your Excel file is located.
setwd("C:\\Users\\jsaba\\Desktop\\Paper 1")

# Read data from the specified Excel file and sheet. MoE has been calculated at 99% CI and saved 
# at the working directory. 
data <- read_excel("FEMA_4_99MoE.xlsx", sheet = "Sheet 1")

## Step 3: Set the Seed for Reproducibility
# This ensures that the random number generation is reproducible.
set.seed(1231)  # You can change 1231 to any other number if desired.

## Step 4: Define the Monte Carlo Simulation Function
monte_carlo_iteration <- function(MoE_df, X1, X2, num_simulations) {
  # Calculate upper and lower bounds for the simulation.
  upper_bound <- X1 + X2
  lower_bound <- X1 - X2
  
  # Initialize a matrix to store simulated data.
  simulated_data <- matrix(NA, nrow = nrow(MoE_df), ncol = num_simulations)
  
  # For each row in the data, generate simulated values.
  for (i in 1:nrow(MoE_df)) {
    # Ensure that the bounds are valid.
    if (!is.na(lower_bound[i]) && !is.na(upper_bound[i]) && upper_bound[i] >= lower_bound[i]) {
      # Generate random numbers from a normal distribution with mean at the midpoint and SD set so that ~99.7% values lie within the bounds.
      simulated_data[i, ] <- rnorm(num_simulations, 
                                   mean = (upper_bound[i] + lower_bound[i]) / 2, 
                                   sd = (upper_bound[i] - lower_bound[i]) / 6)
    } else {
      # If bounds are not valid, fill the row with NA.
      simulated_data[i, ] <- rep(NA, num_simulations)
    }
  }
  
  return(simulated_data)
}

## Step 5: Run Monte Carlo Simulations for Variable Pairs
# Define simulation parameters.
num_iterations <- 100
num_simulations_per_variable <- 100

# Create a list to store simulation results.
simulation_results <- vector("list", length = num_iterations)

# Define the list of paired variables from the data.
all_variables <- list(data$EP_POV150, data$MP_POV150, 
                      data$EP_UNEMP, data$MP_UNEMP, 
                      data$EP_HBURD, data$MP_HBURD, 
                      data$EP_NOHSDP, data$MP_NOHSDP, 
                      data$EP_UNINSUR, data$MP_UNINSUR)

# Loop over the variable list in pairs (odd-indexed as X1 and even-indexed as X2).
for (j in 1:(length(all_variables)-1)) {
  if (j %% 2 != 0) {
    for (iteration in 1:num_iterations) {
      x1 <- as.numeric(unlist(all_variables[[j]]))
      x2 <- as.numeric(unlist(all_variables[[j+1]]))
      simulated_data <- monte_carlo_iteration(data, x1, x2, num_simulations_per_variable)
      simulation_results[[iteration]] <- simulated_data
    }
    simulation_results_df <- as.data.frame(simulation_results)
    # Bind additional identifier columns to the simulation results.
    simulation_results_df_final <- cbind(data$ST, data$STATE, data$ST_ABBR, 
                                         data$STCNTY, data$COUNTY, data$FIPS, simulation_results_df)
    final_df <- paste('variable', as.character(j), sep = '_')
    assign(final_df, simulation_results_df)
  }
}

## Step 6: Calculate Percentile Ranks for the First Simulation (EPL)
# Define a function to calculate percentile ranks (scaling ranks between 0 and 1).
calculate_percentile_rank <- function(variable_1, significance = 4) {
  percentile_rank_data <- apply(variable_1, 2, function(column) {
    rank_values <- rank(column, na.last = "keep")
    percentile_rank <- round((rank_values - 1) / (length(rank_values) - 1), significance)
    return(percentile_rank)
  })
  return(percentile_rank_data)
}

# Calculate the percentile rank for 'variable_1'.
percent_rank_data <- calculate_percentile_rank(variable_1)
percent_rank_df <- as.data.frame(percent_rank_data)

## Step 7: Calculate Percentile Ranks for All Specified Simulation Variables
# List the variable names to process.
variable_names <- c("variable_1", "variable_3", "variable_5", "variable_7", "variable_9")

# Loop through each simulation result, calculate percentile ranks, and assign to a new object.
for (variable_name in variable_names) {
  variable_df <- get(variable_name)
  percent_rank_data <- calculate_percentile_rank(variable_df)
  percent_rank_df <- as.data.frame(percent_rank_data)
  new_variable_name <- paste("percentrank", variable_name, sep = "_")
  assign(new_variable_name, percent_rank_df)
}

## Step 8: Calculate SPL (Summed Percentile Level)
# Create a list of the percentile rank dataframes.
data_frames_list <- list(percentrank_variable_1, percentrank_variable_3, 
                         percentrank_variable_5, percentrank_variable_7, 
                         percentrank_variable_9)

# Initialize a result matrix filled with zeros.
result_matrix <- matrix(0, nrow = nrow(variable_1), ncol = ncol(variable_1))

# Sum the percentile values across the specified dataframes.
for (col_index in 1:10000) {
  for (df in data_frames_list) {
    result_matrix[, col_index] <- result_matrix[, col_index] + df[, col_index]
  }
}

# Convert the summed result to a dataframe and save as CSV.
result_data <- data.frame(result_matrix)
write.csv(result_data, "result_data_final_df_1.csv", row.names = FALSE)

## Step 9: Calculate RPL (Relative Percentile Level) and Additional Statistics
# Redefine the percentile rank function for RPL calculation (same logic as before).
calculate_percentile_rank <- function(result_data, significance = 4) {
  RPL_rank_data_1 <- apply(result_data, 2, function(column) {
    rank_values <- rank(column, na.last = "keep")
    percentile_rank <- round((rank_values - 1) / (length(rank_values) - 1), significance)
    return(percentile_rank)
  })
  return(RPL_rank_data_1)
}

# Calculate RPL based on the SPL results.
RPL_rank_data_1 <- calculate_percentile_rank(result_data)
RPL_rank_df_1 <- as.data.frame(RPL_rank_data_1)

# Calculate row-wise mean and standard deviation.
mean_svi <- rowMeans(RPL_rank_df_1, na.rm = TRUE)
stdv_svi <- apply(RPL_rank_df_1, 1, function(x) sd(x, na.rm = TRUE))
z_score <- 1.645

# Define and calculate the Margin of Error (MoE).
calculate_MoE <- function(z_score, std_dev) {
  MoE <- z_score * (std_dev / sqrt(100))
  return(MoE)
}
MoE_values <- calculate_MoE(z_score, stdv_svi)

# Define and calculate the Coefficient of Variation (CV).
calculate_CV <- function(MoE, mean) {
  CV <- (MoE / 1.645) / mean * 100
  return(CV)
}
CV_values <- calculate_CV(MoE_values, mean_svi)

# Append calculated statistics to the RPL dataframe.
RPL_rank_df_1$Mean <- mean_svi
RPL_rank_df_1$SD <- stdv_svi
RPL_rank_df_1$Z_Score <- z_score
RPL_rank_df_1$MoE <- MoE_values
RPL_rank_df_1$CV <- CV_values

# Combine with county and location information.
RPL_rank_df_1 <- cbind(data$ST, data$STATE, data$ST_ABBR, data$STCNTY, 
                       data$COUNTY, data$FIPS, data$LOCATION, RPL_rank_df_1)

# Save the final RPL results to Excel and CSV files.
write.xlsx(RPL_rank_df_1, "RPL_rank_1_df.xlsx", row.names = FALSE)
write.csv(RPL_rank_df_1, "RPL_rank_1_df.csv", row.names = FALSE)
