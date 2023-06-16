###------------------------------------###
#### R Syntax for Nelson 1994 T-Chart ####
###------------------------------------###

# V0.1 - Produce T-Chart in line with Nelson (1994).
#        Adjustment so T-Chart has Centreline, LCL, and UCL annotated with respective values.

# Install Packages
install.packages("tidyverse")
install.packages("janitor")
install.packages("rio")
install.packages("lubridate")
install.packages("scales")
install.packages("openxlsx")

# Libraries
library(tidyverse)
library(janitor)
library(rio)
library(lubridate)
library(scales)
library(openxlsx)

###-------------###
### Sample Data ###
###-------------###

Data <- import("[INSERT FILE PATH HERE]/2_Nelson_1994_T_Chart_Sample_Data.xlsx", sheet = 1)
Data <- import("[INSERT FILE PATH HERE]/2_Nelson_1994_T_Chart_Sample_Data.xlsx", sheet = 2)

###----------------------------------###
#### Making a Decision Around Zeros ####
###----------------------------------###

Data <- Data %>%
  select(Date_of_Rare_Event = [INSERT COLUMN NAME HERE])

# # Option 1 for Zero Values
# # Remove zeros (because these are not accepted when qic() is set to 't' as the chart type.
# Remove_Zeros <- Data %>%
#   filter(Time_Between == 0)
# Data <- Data %>%
#   anti_join(Remove_Zeros)
# 
# # Option 2 for Zero Values
# # Converting the zeros to 0.5 instead.
# Data <- Data %>%
#   mutate(Time_Between = ifelse(Time_Between == 0, 0.5, Time_Between))

###-------------------------------------------------###
#### Carrying Out the Calculations for the T-Chart ####
###-------------------------------------------------###

# Calculate the Time Between value for each.
Data <- Data %>%
  arrange(Date_of_Rare_Event) %>%
  mutate(Date_of_Rare_Event = as.Date(Date_of_Rare_Event)) %>%
  mutate(Time_Between = Date_of_Rare_Event - lag(Date_of_Rare_Event, n = 1)) %>% # Go back from current Date row by n=1 and return that value. 
  mutate(Time_Between = as.numeric(Time_Between))

# Calculate the y value for each (i.e. transform the time between variable).
Data <- Data %>%
  mutate(y_Value = Time_Between ^ 0.2777) %>%
  mutate(y_Value = round(y_Value, 1))

# Calculate Y-Bar (i.e. mean of the y values).
Data <- Data %>%
  mutate(y_Bar = mean(y_Value, na.rm = T)) %>%
  mutate(y_Bar = round(y_Bar, 1))

# Calculate the Moving Range value for each.
Data <- Data %>%
  mutate(Moving_Range = y_Value - lag(y_Value, n = 1)) %>% # Go back from current Date row by n=1 and return that value. 
  mutate(Moving_Range = abs(Moving_Range))

# Calculate Moving Range-Bar (i.e. the mean of the Moving Range values).
Data <- Data %>%
  mutate(MR_Bar = mean(Moving_Range, na.rm = T)) %>%
  mutate(MR_Bar = round(MR_Bar, 1))

# Calculate Moving Range_Bar * 3.27 (i.e. multiply by 3.27).
Data <- Data %>%
  mutate(MR_Bar_3.27 = MR_Bar * 3.27) %>%
  mutate(MR_Bar_3.27 = round(MR_Bar_3.27, 1))

# Identify those Moving Range values which are outliers (i.e. MR > MR-Bar * 3.27).
Data <- Data %>%
  mutate(Outlier_Warning = ifelse(Moving_Range >= MR_Bar_3.27,
                                  "Outlier", "Not Outlier"))

# Carry over those values that are not outliers. 
Data <- Data %>%
  mutate(Keep = ifelse(Outlier_Warning == "Outlier", 0, Moving_Range))

# Calculate the Moving Range-Bar again without any outliers which might skew it. 
# Call this MR-Bar' (MR_Bar_Adjusted in the syntax below). Where there are no outliers,
# the "Keep" column will be the same as the "Moving_Range" column, and the MR_Bar_Adjusted 
# column will be the same as the Mr_Bar column. 
Data <- Data %>%
  mutate(MR_Bar_Adjusted = mean(Keep, na.rm = T)) %>%
  mutate(MR_Bar_Adjusted = round(MR_Bar_Adjusted, 1))

# Most of the intermediate calculations have now been carried out (except for Upper Limit
# and Lower Limit). This next step calculates the Centreline, Upper Control Limit, and 
# Lower Control Limit (as well as the final two intermediate calculations needed).
Data <- Data %>%
  mutate(Centreline = y_Bar^3.6,
         Centreline = round(Centreline, 1),
         Upper_Limit = y_Bar + 2.66 * MR_Bar_Adjusted,
         Upper_Limit = round(Upper_Limit, 1),
         Lower_Limit = y_Bar - 2.66 * MR_Bar_Adjusted,
         Lower_Limit = round(Lower_Limit, 1),
         Upper_Control_Limit = Upper_Limit^3.6,
         Upper_Control_Limit = round(Upper_Control_Limit , 1),
         Lower_Control_Limit = Lower_Limit^3.6,
         Lower_Control_Limit = round(Lower_Control_Limit, 1))

Data <- Data %>%
  mutate(Lower_Control_Limit = ifelse(Lower_Control_Limit == "NaN", 0, Lower_Control_Limit),
         Lower_Control_Limit = ifelse(Lower_Control_Limit < 0.5, 0, Lower_Control_Limit))


###-----------------------------------------------###
#### Plotting the T-Chart Using Package ggplot 2 ####
###-----------------------------------------------###

Graph_Data <- Data %>%
  select(Date_of_Rare_Event, Time_Between, Centreline, Upper_Control_Limit, 
         Lower_Control_Limit) %>%
  mutate(Date_of_Rare_Event = as.character(Date_of_Rare_Event),
         Above = ifelse(Time_Between > Centreline, 
                        Time_Between, NA),
         Below = ifelse(Time_Between < Centreline & Time_Between > Lower_Control_Limit,
                        Time_Between, NA),
         Outlier_LCL = ifelse(Time_Between <= Lower_Control_Limit, Time_Between, NA)) %>%
  mutate(Outlier_LCL = as.numeric(Outlier_LCL))

# Linetype has to be outside of the aes mapping for it to work, not within those brackets.
# Colour outside the aes mapping brackets will be that colour but not included in the legend.
# Colour inside aes mapping brackets, then added to scale_colour_manual, will be included in legend.

# Title for the graph.
Graph_Title <- paste0("T-Chart for Rare Events")
Graph_Title

# Annotate the Centreline.
Centreline_Value <- Graph_Data$Centreline %>% head(1)
Centreline_Annotation <- paste0("CL = ", Centreline_Value, " Days")
#Position_CL_Annotation <- Graph_Data$Date_of_Rare_Event %>% head(2) %>% tail(1)
Position_CL_Annotation <- Graph_Data$Date_of_Rare_Event %>% tail(2) %>% head(1)

# Annotate the Lower Control Limit.
LCL_Value <- Graph_Data$Lower_Control_Limit %>% head(1)
LCL_Annotation <- paste0("LCL = ", LCL_Value, " Days")
#Position_LCL_Annotation <- Graph_Data$Date_of_Rare_Event %>% head(2) %>% tail(1)
Position_LCL_Annotation <- Graph_Data$Date_of_Rare_Event %>% tail(2) %>% head(1)

# Annotate the Upper Control Limit.
UCL_Value <- Graph_Data$Upper_Control_Limit %>% head(1)
UCL_Annotation <- paste0("UCL = ", UCL_Value, " Days")
#Position_UCL_Annotation <- Graph_Data$Date_of_Rare_Event %>% head(2) %>% tail(1)
Position_UCL_Annotation <- Graph_Data$Date_of_Rare_Event %>% tail(2) %>% head(1)

Graph <- ggplot(Graph_Data) +
  geom_line(aes(x = Date_of_Rare_Event, 
                y = Centreline, group = 1, colour = "Centreline"), linetype = "solid", linewidth = 1.0) +
  geom_line(aes(x = Date_of_Rare_Event, 
                y = Upper_Control_Limit, group = 1, colour = "UCL"), linetype = "dashed", linewidth = 1.0) +
  geom_line(aes(x = Date_of_Rare_Event, 
                y = Lower_Control_Limit, group = 1, colour = "LCL"), linetype = "dashed", linewidth = 1.0) +
  geom_point(aes(x = Date_of_Rare_Event, 
                 y = Above), colour = "black", show.legend = FALSE) + # Change colour if needed.
  geom_point(aes(x = Date_of_Rare_Event, 
                 y = Below), colour = "black", show.legend = FALSE) + # Change colour if needed.
  geom_point(aes(x = Date_of_Rare_Event, 
                 y = Outlier_LCL), colour = "black") + # Change colour if needed.
  geom_line(aes(x = Date_of_Rare_Event, 
                y = Time_Between, group = 1), colour = "black", linetype = "solid", linewidth = 0.5, 
            alpha = 0.5) +
  labs(title = Graph_Title,
       x = "Date of Rare Event",
       y = "Time Between (Days)",
       colour = "") +
  theme(axis.text.x = element_text(color = "black", angle = 90, vjust = 0.5, hjust=1),
        axis.text.y = element_text(color = "black", size = 8),
        plot.background = element_blank(),
        panel.background = element_blank(),
        axis.line = element_line(colour = "black"),
        axis.ticks = element_line(color = "black"),
        axis.title.x = element_text(margin = margin(t = 10), color = "black", size = 10),
        axis.title.y = element_text(margin = margin(r = 10), color = "black", size = 10),
        legend.position = "bottom") +
  scale_colour_manual(values = c("Centreline" = "black",
                                 "UCL" = "darkgreen",
                                 "LCL" = "darkred")) +
  annotate("text", x = Position_CL_Annotation, y = Centreline_Value + 2, label = Centreline_Annotation, size = 2, colour = "black") +
  annotate("text", x = Position_LCL_Annotation, y = LCL_Value + 1.5, label = LCL_Annotation, size = 2, colour = "darkred") +
  annotate("text", x = Position_UCL_Annotation, y = UCL_Value + 1.5, label = UCL_Annotation, size = 2, colour = "darkgreen")

# For legend.position c(0,0) corresponds to bottom left and c(1,1) corresponds to top right. It is also
# possible to use "top", "bottom", "right", and "left". 
print(Graph)

# Create workbook and write data. 
wb1 <- createWorkbook()
addWorksheet(wb1, "T-Chart Data")
writeData(wb1, "T-Chart Data", Data, startCol = 1, startRow = 1)

addWorksheet(wb1, "T-Chart") 

writeData(wb1, "T-Chart",
          paste0("T-Chart for Time Between Rare Events"),
          startRow = 1, startCol = 1)

insertPlot(wb1, "T-Chart",
           startRow = 3, startCol = 2,
           height = 4.5, width = 8.5, units = "in",
           fileType = "png",  dpi = 500)

saveWorkbook(wb1,
             file = paste0("[INSERT FILE PATH HERE]/", today(),"_T_Chart_And_Data.xlsx"),
             overwrite = TRUE)

rm(wb1)
