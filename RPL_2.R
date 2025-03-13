################################################################################
## Step 1: Package Installation and Library Loading
################################################################################
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

################################################################################
## Step 2: Set the Working Directory and Read Data
################################################################################
# Adjust the working directory to where your Excel file is located. 
setwd("C:\\Users\\jsaba\\Desktop\\Paper 1")

# Read data from the specified Excel file and sheet.MoE has been calculated at 99% CI and saved 
# at the working directory.
data <- read_excel("FEMA_4_99MoE.xlsx", sheet = "Sheet 1")

################################################################################
## Step 3: Set the Seed for Reproducibility
################################################################################
# This ensures that random number generation is reproducible.
set.seed(1232)  # Change the seed value if needed.

################################################################################
## Step 4: Define the Monte Carlo Simulation Function
################################################################################
monte_carlo_iteration <- function(MoE_df, X1, X2, num_simulations) {
  # Calculate Upper and Lower Bounds
  upper_bound <- X1 + X2
  lower_bound <- X1 - X2
  
  # Initialize simulated_data matrix
  simulated_data <- matrix(NA, nrow = nrow(MoE_df), ncol = num_simulations)
  
  # For each row in the dataset, generate simulated values if bounds are valid.
  for (i in 1:nrow(MoE_df)) {
    if (!is.na(lower_bound[i]) && !is.na(upper_bound[i]) && upper_bound[i] >= lower_bound[i]) {
      simulated_data[i, ] <- rnorm(num_simulations, 
                                   mean = (upper_bound[i] + lower_bound[i]) / 2, 
                                   sd = (upper_bound[i] - lower_bound[i]) / 6)
    } else {
      simulated_data[i, ] <- rep(NA, num_simulations)
    }
  }
  return(simulated_data)
}

################################################################################
## Step 5: Run Monte Carlo Simulations for Variable Pairs
################################################################################
# Define simulation parameters.
num_iterations <- 100
num_simulations_per_variable <- 100

# Create a list to store simulation results.
simulation_results <- vector("list", length = num_iterations)

# Define the list of paired variables (note: MP_DISABL is counted twice as per comment).
all_variables_2 <- list(data$EP_AGE65, data$MP_AGE65, 
                        data$EP_AGE17, data$MP_AGE17, 
                        data$EP_DISABL, data$MP_DISABL,
                        data$EP_SNGPNT, data$MP_SNGPNT, 
                        data$EP_LIMENG, data$MP_LIMENG)

# Loop through the variable pairs (odd-indexed as X1 and the next even-indexed as X2).
for (j in 1:(length(all_variables_2)-1)) {
  if (j %% 2 != 0) {
    for (iteration in 1:num_iterations) {
      x1 <- as.numeric(unlist(all_variables_2[[j]]))
      x2 <- as.numeric(unlist(all_variables_2[[j+1]]))
      simulated_data <- monte_carlo_iteration(data, x1, x2, num_simulations_per_variable)
      simulation_results[[iteration]] <- simulated_data
    }
    simulation_results_df_2 <- as.data.frame(simulation_results)
    # Combine simulation results with identifying county/state information.
    simulation_results_df_final_2 <- cbind(data$ST, data$STATE, data$ST_ABBR, 
                                           data$STCNTY, data$COUNTY, data$FIPS, simulation_results_df_2)
    final_df_2 <- paste('variable', as.character(j), sep = '_')
    assign(final_df_2, simulation_results_df_2)
  }
}

################################################################################
## Step 6: Calculate Percentile Ranks (EPL Calculation)
################################################################################
# Define a function to calculate percentile ranks (scaling between 0 and 1).
calculate_percentile_rank <- function(variable_1, significance = 4) {
  percentile_rank_data <- apply(variable_1, 2, function(column) {
    rank_values <- rank(column, na.last = "keep")
    percentile_rank <- round((rank_values - 1) / (length(rank_values) - 1), significance)
    return(percentile_rank)
  })
  return(percentile_rank_data)
}

# Calculate percentile rank for the first simulation output ("variable_1").
percent_rank_data <- calculate_percentile_rank(variable_1)
percent_rank_df <- as.data.frame(percent_rank_data)

################################################################################
## Step 7: Calculate Percentile Ranks for All Specified Simulation Variables
################################################################################
# List of simulation variable names to process.
variable_names <- c("variable_1", "variable_3", "variable_5", "variable_7", "variable_9")

# Loop through each simulation result, compute percentile ranks, and assign to a new object.
for (variable_name in variable_names) {
  variable_df <- get(variable_name)
  percent_rank_data <- calculate_percentile_rank(variable_df)
  percent_rank_df <- as.data.frame(percent_rank_data)
  new_variable_name <- paste("percentrank", variable_name, sep = "_")
  assign(new_variable_name, percent_rank_df)
}

################################################################################
## Step 8: Calculate SPL
################################################################################
# Combine the percentile rank dataframes into a list.
data_frames_list <- list(percentrank_variable_1, percentrank_variable_3, 
                         percentrank_variable_5, percentrank_variable_7, 
                         percentrank_variable_9)

# Initialize a result matrix with zeros.
result_matrix_2 <- matrix(0, nrow = nrow(variable_1), ncol = ncol(variable_1))

# Sum the percentile values across all dataframes for each column.
for (col_index in 1:10000) {
  for (df in data_frames_list) {
    result_matrix_2[, col_index] <- result_matrix_2[, col_index] + df[, col_index]
  }
}

# Convert the summed result into a dataframe and save it.
result_data_final_df_2 <- as.data.frame(result_matrix_2)
write.csv(result_data_final_df_2, "result_data_final_df_2.csv", row.names = FALSE)

################################################################################
## Step 9: Calculate RPL 
################################################################################
# Redefine the percentile rank function (used for RPL calculation).
calculate_percentile_rank <- function(result_data, significance = 4) {
  RPL_rank_2_data <- apply(result_data_2, 2, function(column) {
    rank_values <- rank(column, na.last = "keep")
    percentile_rank <- round((rank_values - 1) / (length(rank_values) - 1), significance)
    return(percentile_rank)
  })
  return(RPL_rank_2_data)
}

# Calculate RPL based on the SPL results.
RPL_rank_2_data <- calculate_percentile_rank(result_data)
RPL_rank_2_df <- as.data.frame(RPL_rank_2_data)

################################################################################
## Step 10: Additional Statistics
################################################################################
# Calculate row-wise Mean and Standard Deviation.
mean_svi <- rowMeans(RPL_rank_2_df, na.rm = TRUE)
stdv_svi <- apply(RPL_rank_2_df, 1, function(x) sd(x, na.rm = TRUE))
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

# Append the calculated statistics to the RPL dataframe.
RPL_rank_2_df$Mean <- mean_svi
RPL_rank_2_df$SD <- stdv_svi
RPL_rank_2_df$Z_Score <- z_score
RPL_rank_2_df$MoE <- MoE_values
RPL_rank_2_df$CV <- CV_values

# Combine the RPL results with county and location information.
RPL_rank_2_df <- cbind(data$ST, data$STATE, data$ST_ABBR, data$STCNTY, 
                       data$COUNTY, data$FIPS, data$LOCATION, RPL_rank_2_df)

# Save the final RPL results to Excel and CSV files.
write.xlsx(RPL_rank_2_df, "RPL_rank_2_df.xlsx", row.names = FALSE)
write.csv(RPL_rank_2_df, "RPL_rank_2_df.csv", row.names = FALSE)
