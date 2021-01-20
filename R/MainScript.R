# Load packages
library(openxlsx)
library(dplyr)
library(tidyr)
library(ggplot2)
library(reshape2)

# Set names of xlsx files
file <- data.frame( "Z" = "2007_USA_Z_Matrix_pxp.xlsx",
                    "Make" = "IOMake_After_Redefinitions_DET.xlsx",
                    "Use" = "IOUse_After_Redefinitions_PRO_DET.xlsx",
                    "Filter" = "Filter_settings.xlsx",
                    stringsAsFactors = FALSE )

# Set paths to folder
path <- data.frame( "input" = paste0(getwd(),"/input/") ,
                    "output" = paste0(getwd(),"/output/") ,
                    "R" = paste0(getwd(),"/R/") ,
                    stringsAsFactors = FALSE ) 

# Create time stamp for identifying output files
time <- Sys.time()
time <- gsub(":","-",time)

# Read tables
raw <- list()
raw[["Z"]] <- read.xlsx(xlsxFile = paste0(path$input, file$Z) , sheet = 1)  # Transaction matrix
raw[["U"]] <- read.xlsx(xlsxFile = paste0(path$input, file$Use) , sheet = 3)  # Use table
raw[["Y"]] <- raw$U[ 5:(5+400), 429]  # Domestic final demand

# Clean Z tables and convert to numeric
Z <- raw$Z[ 3:403, 3:403]
Z <- apply(Z,c(1,2),as.numeric)

# Set NA to zero and convert Y to numeric 
Y <- raw$Y
Y[is.na(Y)] <- 0
Y <- as.numeric(Y)

Y[Y < 0]  # Negatives in final demand

# Check sums
sum(Y)
sum(Z)

# Estimate gross production
x <- rowSums(Z) + Y

# Calculate technology matrix
A <- t( t(Z)/x )

# Remove col and rownames
rownames(A) <- colnames(A) <- NULL

# Check sum and if all positive numeric
sum(A)
min(A)

# Load settings for present run
tmp <- read.xlsx(xlsxFile = paste0( path$input, file$Filter), sheet = 1 )

# Read name of material
material <- colnames(tmp)[2]

# Clean filter and add column names
filter <- tmp[2:402, 3:4 ]
colnames(filter) <- c("material","product")
filter <- as.data.frame( apply(filter, c(1,2), as.numeric) )

# Clean end-use category aggregator
EndUse <- tmp[2:402, 5:ncol(tmp) ]
colnames(EndUse) <- tmp[1, 5:ncol(tmp) ]
EndUse <- apply(EndUse, c(1,2), as.numeric)

# Read sector names
IO.code <- tmp[2:402, 2]

# Apply filter to technology matrix
A_new <- A * rowSums(filter)
A_new <- t( t(A_new) * filter$product )

# Add labels and write filtered technology matrix to file
colnames(A_new) <- rownames(A_new) <- IO.code
write.xlsx( x = A_new, file = paste0( path$output, time," ",material, " A_new.xlsx"), row.names = TRUE )

# Estimate new Leontief inverse and new gross production vector
I <- diag( rep(1,401) )
L_new <- solve( I - A_new )

# Estimate flow matrix
X_new <- L_new %*% diag(Y)

# Explanation: Matrix X_new quantifies how much (gross) output/production of row-sector i
# is directly and indirectly required to satisfy demand for final product column-sector j.
# Row sums must add up to the modified gross production vector x_new 

# Check gross production of material sector(s)
sum(filter$material * X_new)  

# Select material sector in rows
Material2EndUse <- colSums(filter$material * X_new)

# Aggregate final products (columns) to End-Use Categories
Material2EndUse <- Material2EndUse %*% EndUse 

# Estimate shares and round to 2 digits
Shares <- round( Material2EndUse / sum(Material2EndUse), digits = 2 )
rownames(Shares) <- material

# Write result to files
write.xlsx( x = Shares, file = paste0( path$output, time," ",material, " End-Use-Shares",".xlsx"), row.names = TRUE )

# Make copy of present settings file and rename accordingly
file.copy(from = paste0(path$input, "Filter_settings.xlsx"), 
          to = path$output )

file.rename(from = file.path(path$output, "Filter_settings.xlsx"), 
            to = file.path(path$output, paste0(time, " ", material, " Filter_settings.xlsx") ))


# FIN



