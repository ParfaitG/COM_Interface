options(scipen = 999)

set.seed(1222025)

### DATA BUILD
metals_wide_df <- data.frame(
  year = 1950:2019,
  gold = runif(70, 1, 2000),
  silver = runif(70, 1, 100),
  platinum = runif(70, 1, 1000),
  palladium = runif(70, 1, 2000),
  rhodium = runif(70, 1, 1000)
)

comma_output <- function(df, row=10)
  cat(paste0(paste(names(df), collapse=", "), "\n"),
      apply(tail(df, row), 1, function(x) paste0(paste(x, collapse=", "), "\n")))



metals_wide_df[,2:6] <- round(metals_wide_df[,2:5], 2)

comma_output(metals_wide_df)



metals_long_df <- reshape(metals_wide_df, varying = names(metals_wide_df)[-1], times = names(metals_wide_df)[-1], 
                          idvar = "year", ids = NULL, v.names = "avg_spot_price", timevar = "metal",
                          new.row.names = 1:1E3, direction = "long")

metals_long_df <- with(metals_long_df, metals_long_df[order(year, metal),])


comma_output(metals_long_df)




agg_df <- do.call(data.frame, 
                  aggregate(avg_spot_price ~ metal, metals_long_df, function(x) 
                            c(min=min(x), 
                              median=median(x), 
                              mean=mean(x), 
                              max=max(x),
                              sd=sd(x)))
)
agg_df[,2:6] <- round(agg_df[,2:6],4)


names(agg_df) <- gsub("avg_spot_price.", "", names(agg_df))

comma_output(agg_df)



properties_df <- data.frame(
  metal = c("gold", "silver", "platinum", "palladium", "rhodium"),
  atomic_number = c(79, 47, 78, 46, 45),
  symbol = c("Au", "Ag", "Pt" , "Pd", "Rh"),
  atomic_weight = c(196.967, 107.868, 195.09, 106.40, 102.905),
  melting_point = c(1063, 960.8, 1769, 1554.9, 1966),
  boiling_point = c(2966, 2212, 3827, 2963, 3727),
  stringsAsFactors = FALSE
)

metals_long_df <- merge(metals_long_df, properties_df, by="metal")


write.csv(metals_long_df, "E:\\Sandbox\\Precious_Metals.csv", row.names = FALSE)

metals_long_df <- with(metals_long_df, metals_long_df[order(date, metal),])

comma_output(metals_long_df)


results <- summary(lm(avg_spot_price ~ date + atomic_weight + melting_point + boiling_point, metals_long_df))

comma_output(data.frame(cbind(row.names(results$coefficients), round(results$coefficients, 4))))

seaborn_palette <- c("#4C72B0", "#DD8452", "#55A868", "#C44E52", "#8172B3", "#937860", 
                     "#DA8BC3", "#8C8C8C", "#CCB974", "#64B5CD", "#4C72B0", "#DD8452")


png('E:\\Sandbox\\Precious_Metals.png')
  boxplot(avg_spot_price ~ metal, metals_long_df, col=seaborn_palette[1:5])
dev.off()

with(agg_long_df[], barplot(tapply(avg_spot_price, list(metal, year), mean), col=seaborn_palette[1:5]))


plt_colors <- as.list(setNames(seaborn_palette[1:5], unique(metals_long_df$metal)))

png('E:\\Sandbox\\Precious_Metals_Year.png')
  par(mfrow=c(3,2), mar=c(3, 5, 3, 1), oma=c(0, 0, 2, 0)) 
  by(metals_long_df, metals_long_df$metal, function(sub)
    with(sub, {
      plot(year, avg_spot_price, main=sub$metal[[1]], pch=16, col=plt_colors[[sub$metal[[1]]]])
      lines(year, avg_spot_price, col=plt_colors[[sub$metal[[1]]]])
    })
  )
  mtext("Average Spot Price of Precious Metals, 1950-2020", outer = TRUE, cex = 1.5)
dev.off()



