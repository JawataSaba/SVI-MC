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
# Adjust the working directory to the folder where your Excel file is located.
setwd("C:\\Users\\jsaba\\Desktop\\Paper 1")

# Read the data from the specified Excel file and sheet. MoE has been calculated at 99% CI and saved 
# at the working directory. 
data <- read_excel("FEMA_4_99MoE.xlsx", sheet = "Sheet 1")

################################################################################
## Step 3: Set the Seed for Reproducibility
################################################################################
# Set the seed so that the random simulations are reproducible.
set.seed(1234)  # You may change this seed value if desired.

################################################################################
## Step 4: Define the Monte Carlo Simulation Function
################################################################################
monte_carlo_iteration <- function(MoE_df, X1, X2, num_simulations) {
  # Calculate Upper and Lower Bounds for each observation.
  upper_bound <- X1 + X2
  lower_bound <- X1 - X2
  
  # Initialize a matrix to store simulated data.
  simulated_data <- matrix(NA, nrow = nrow(MoE_df), ncol = num_simulations)
  
  # For each row, generate simulated values if bounds are valid.
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
## Step 5: Run Monte Carlo Simulations for Theme 4 Variables
################################################################################
# Define simulation parameters.
num_iterations <- 100
num_simulations_per_variable <- 100

# Create a list to store simulation results.
simulation_results <- vector("list", length = num_iterations)

# Define the list of paired variables for Theme 4.
all_variables_4 <- list(data$EP_MUNIT, data$MP_MUNIT, 
                        data$EP_MOBILE, data$MP_MOBILE, 
                        data$EP_CROWD, data$MP_CROWD,
                        data$EP_NOVEH, data$MP_NOVEH, 
                        data$EP_GROUPQ, data$MP_GROUPQ)

# Loop through each variable pair (odd-indexed variable as X1 and next even-indexed as X2).
for (j in 1:(length(all_variables_4)-1)) {
  if (j %% 2 != 0) {
    for (iteration in 1:num_iterations) {
      x1 <- as.numeric(unlist(all_variables_4[[j]]))
      x2 <- as.numeric(unlist(all_variables_4[[j+1]]))
      simulated_data <- monte_carlo_iteration(data, x1, x2, num_simulations_per_variable)
      simulation_results[[iteration]] <- simulated_data
    }
    simulation_results_df_4 <- as.data.frame(simulation_results)
    final_df_4 <- paste("variable", as.character(j), sep = "_")
    assign(final_df_4, simulation_results_df_4)
  }
}

################################################################################
## Step 6: Calculate EPL (Percentile Rank Calculation)
################################################################################
# Define a function to calculate percentile ranks, scaling from 0 (lowest) to 1 (highest).
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
## Step 7: Calculate Percentile Ranks for All Simulation Variables
################################################################################
# Process each simulation result to compute its percentile rank.
variable_names <- c("variable_1", "variable_3", "variable_5", "variable_7", "variable_9")
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
# Combine the percentile rank data frames into a list.
data_frames_list <- list(percentrank_variable_1, percentrank_variable_3, 
                         percentrank_variable_5, percentrank_variable_7, 
                         percentrank_variable_9)

# Initialize a result matrix with zeros.
result_matrix_4 <- matrix(0, nrow = nrow(variable_1), ncol = ncol(variable_1))

# Sum the percentile values across all specified data frames column-wise.
for (col_index in 1:10000) {
  for (df in data_frames_list) {
    result_matrix_4[, col_index] <- result_matrix_4[, col_index] + df[, col_index]
  }
}

# Convert the summed matrix to a data frame and save as CSV.
result_data_4 <- data.frame(result_matrix_4)
result_data_final_df_4 <- as.data.frame(result_data_4)
write.csv(result_data_final_df_4, "result_data_final_df_4.csv", row.names = FALSE)

################################################################################
## Step 9: Calculate RPL 
################################################################################
# Redefine the percentile rank function to compute the RPL from the SPL data.
calculate_percentile_rank <- function(result_data, significance = 4) {
  RPL_rank_4_data <- apply(result_data_4, 2, function(column) {
    rank_values <- rank(column, na.last = "keep")
    percentile_rank <- round((rank_values - 1) / (length(rank_values) - 1), significance)
    return(percentile_rank)
  })
  return(RPL_rank_4_data)
}

# Calculate RPL based on the aggregated SPL results.
RPL_rank_4_data <- calculate_percentile_rank(result_data_4)
RPL_rank_4_df <- as.data.frame(RPL_rank_4_data)

################################################################################
## Step 10: Additional Statistics
################################################################################
# Calculate row-wise Mean and Standard Deviation.
mean_svi <- rowMeans(RPL_rank_4_df, na.rm = TRUE)
stdv_svi <- apply(RPL_rank_4_df, 1, function(x) sd(x, na.rm = TRUE))
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
RPL_rank_4_df$Mean <- mean_svi
RPL_rank_4_df$SD <- stdv_svi
RPL_rank_4_df$Z_Score <- z_score
RPL_rank_4_df$MoE <- MoE_values
RPL_rank_4_df$CV <- CV_values

# Combine county and location information with the RPL results.
RPL_rank_4_df <- cbind(data$ST, data$STATE, data$ST_ABBR, data$STCNTY, 
                       data$COUNTY, data$FIPS, data$LOCATION, RPL_rank_4_df)

# Save the final RPL results to Excel and CSV files.
write.xlsx(RPL_rank_4_df, "RPL_rank_4_df.xlsx", row.names = FALSE)
write.csv(RPL_rank_4_df, "RPL_rank_4_df.csv", row.names = FALSE)
