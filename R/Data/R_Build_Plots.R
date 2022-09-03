
##############################
### PATHS
##############################
project_home <- file.path(
  "C:", "Users", "parfa", "OneDrive", "Documents",
  "Databases", "R_Automation", fsep="\\"
)
csvData <- file.path(
    project_home, "Data", "Precious_Metals_Prices.csv", fsep="\\"
)
boxplotImg <- file.path(
    project_home, "Data", "Precious_Metals_BoxPlot.png", fsep="\\"
)
yearplotImg <- file.path(
    project_home, "Data", "Precious_Metals_YearPlot.png", fsep="\\"
)

##############################
### DATA
##############################

metals_df <- read.csv(csvData)

agg_df <- do.call(
    data.frame, 
    aggregate(
        avg_price ~ metal, metals_df, 
        function(x) c(
            min=min(x), 
            median=median(x), 
            mean=mean(x), 
            max=max(x),
            sd=sd(x)
        )
    )
)
        
agg_df[,2:6] <- round(agg_df[,2:6],4)

seaborn_palette <- c(
    "#4C72B0", "#DD8452", "#55A868", "#C44E52", "#8172B3", "#937860", 
    "#DA8BC3", "#8C8C8C", "#CCB974", "#64B5CD", "#4C72B0", "#DD8452"
)
plt_colors <- as.list(setNames(seaborn_palette[1:5], unique(metals_df$metal)))

png(boxplotImg, width = 1200, height = 800, units = "px")
  par(mar=c(5, 5, 3, 1), oma=c(0, 0, 0, 0)) 
  boxplot(avg_price ~ metal, metals_df, col=seaborn_palette[1:5],
          main="Distribution of Precious Metals Spot Prices, 1950-2020",
          ylab="Spot Price", xlab="Metal",
          cex.main=2, cex.axis=1.5, cex.lab=2)
invisible(dev.off())

png(yearplotImg, width = 1200, height = 800, units = "px")
  par(mfrow=c(2,2), mar=c(4, 5, 3, 1), oma=c(0, 0, 2, 0)) 
  out <- by(metals_df, metals_df$metal, function(sub)
    with(sub, {
      plot(year, avg_price, main=sub$metal[[1]], pch=16, col=plt_colors[[sub$metal[[1]]]], 
           ylab="Spot Price", xlab="Year",
           cex.main=2, cex.axis=1.5, cex.lab=2)
      lines(year, avg_price, col=plt_colors[[sub$metal[[1]]]])
    })
  )
  mtext("Average Spot Price of Precious Metals, 1950-2020", outer = TRUE, cex = 2.0)
invisible(dev.off())