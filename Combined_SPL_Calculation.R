################################################################################
# Step 1: Package Installation and Library Loading
################################################################################
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
# Step 2: Set Working Directory and Read Data
################################################################################
# Adjust the working directory as needed.
setwd("C:\\Users\\jsaba\\Desktop\\Paper 1")

# Read the main Excel file (used here for county and location information)
data <- read_excel("FEMA_4_99MoE.xlsx", sheet = "Sheet 1")

################################################################################
# Step 3: Read SPL Files
################################################################################
# Read the four SPL CSV files into separate data frames.
SPL_file_1 <- read.csv("result_data_final_df_1.csv", header = TRUE)
SPL_file_2 <- read.csv("result_data_final_df_2.csv", header = TRUE)
SPL_file_3 <- read.csv("result_data_final_df_3.csv", header = TRUE)
SPL_file_4 <- read.csv("result_data_final_df_4.csv", header = TRUE)

################################################################################
# Step 4: Combine SPL Files
################################################################################
# Element-wise addition of the four SPL data frames.
total_SPL <- SPL_file_1 + SPL_file_2 + SPL_file_3 + SPL_file_4

# Create a data frame with the combined SPL values.
total_SPL_df <- data.frame(Total_SPL = total_SPL)

# Save the combined SPL data to a CSV file.
write.csv(total_SPL_df, "Total_SPL.csv", row.names = FALSE)

################################################################################
# Step 5: Calculate RPL 
################################################################################
# Define a function to compute percentile ranks (scaling values between 0 and 1).
calculate_percentile_rank <- function(df, significance = 4) {
  apply(df, 2, function(column) {
    rank_values <- rank(column, na.last = "keep")
    round((rank_values - 1) / (length(rank_values) - 1), significance)
  })
}

# Apply the function to the combined SPL data.
RPL_rank_data <- calculate_percentile_rank(total_SPL)
RPL_rank_df <- as.data.frame(RPL_rank_data)

################################################################################
# Step 6: Calculate Summary Statistics (Mean, SD, Margin of Error, CV)
################################################################################
# Calculate row-wise mean and standard deviation.
mean_svi <- rowMeans(RPL_rank_df, na.rm = TRUE)
stdv_svi <- apply(RPL_rank_df, 1, function(x) sd(x, na.rm = TRUE))

# Define a z-score for a 90% confidence interval.
z_score <- 1.645

# Function to calculate the Margin of Error (MoE)
calculate_MoE <- function(z, sd_val) {
  z * (sd_val / sqrt(100))
}
MoE_values <- calculate_MoE(z_score, stdv_svi)

# Function to calculate the Coefficient of Variation (CV)
calculate_CV <- function(MoE, mean_val) {
  (MoE / 1.645) / mean_val * 100
}
CV_values <- calculate_CV(MoE_values, mean_svi)

# Append summary statistics to the RPL data frame.
RPL_rank_df$Mean    <- mean_svi
RPL_rank_df$SD      <- stdv_svi
RPL_rank_df$Z_Score <- z_score
RPL_rank_df$MoE     <- MoE_values
RPL_rank_df$CV      <- CV_values

################################################################################
# Step 7: Combine County/Location Information and Save Final Results
################################################################################
# Combine county and location info from the original data.
RPL_rank_df <- cbind(data$ST, data$STATE, data$ST_ABBR, data$STCNTY, 
                     data$COUNTY, data$FIPS, data$LOCATION, RPL_rank_df)

# Save the final RPL results to Excel and CSV files.
write.xlsx(RPL_rank_df, "RPL_rank_df.xlsx", row.names = FALSE)
write.csv(RPL_rank_df, "RPL_rank_df.csv", row.names = FALSE)
