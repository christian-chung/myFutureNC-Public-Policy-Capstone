library(ggstance)
library(huxtable)
library(jtools)
library(stargazer)
library(sjmisc)
library(foreign)
library(dplyr)
library(tidyr)
library(ggplot2)
library(ggrepel)
library(ggeffects)
library(haven)
library(sjlabelled)
library(Hmisc)
library(readxl)
library(stringr)
library(skimr)
library(dplyr)
library(haven)
library(margins)
library(gtsummary)
library(tables)
library(openxlsx)
library(scales)
library(extrafont)
font_import()  
loadfonts(device = "pdf")  # Use "win" for Windows, "pdf" for PDF, "png" for PNG, etc.

# NCDPI sourced excel sheets
perform <- read_excel("/Users/cakchung/Downloads/2022-23 School Performance Grades (modified).xlsx")
enroll <- read_excel("/Users/cakchung/Downloads/PLCY 698 Capstone - myFutureNC/rcd_college(2022).xlsx")
CCP <- read_excel("/Users/cakchung/Downloads/PLCY 698 Capstone - myFutureNC/CCP_CIHS.xlsx")
CIE <- read_excel("/Users/cakchung/Downloads/PLCY 698 Capstone - myFutureNC/rcd_cie (modified).xlsx")
IB <- read_excel("/Users/cakchung/Downloads/PLCY 698 Capstone - myFutureNC/rcd_ib (modified).xlsx")
AP <- read_excel("/Users/cakchung/Downloads/PLCY 698 Capstone - myFutureNC/rcd_ap (modified).xlsx")
CTE <- read_excel("/Users/cakchung/Downloads/PLCY 698 Capstone - myFutureNC/rcd_cte_enrollment (modified).xlsx")

# Data organization and merging
high.school.performance <- c("UN-GR", "PK-12", "12-13", "11-13", "11-12", "10-12", "0K-12", "0K-11", "09-13", "09-12", "09-11", "08-12", "07-13", "07-12", "06-13", "06-12", "05-12", "04-12", "03-12", "01-12")
low_scores <- c("D", "F", "NA", "I")
All <- c("ALL")
high_school <- perform[perform$grade_span %in% high.school.performance & perform$subgroup %in% c(All, high.school.performance) & perform$spg_grade %in% low_scores, ]
high_school_data <- high_school[c("reporting_year", "lea_code", "lea_name", "school_code", "school_name", "sbe_region", "grade_span", "title_1", "spg_grade", "spg_score")]
# View(high_school_data)
# View(enroll)
# View(CIE)
# View(IB)
# View(AP)
# View(CTE)
# View(CCP)

merged_data <- high_school_data %>%
  left_join(CIE, by = c("school_code" = "agency_code")) %>%
  left_join(IB, by = c("school_code" = "agency_code")) %>%
  left_join(AP, by = c("school_code" = "agency_code")) %>%
  left_join(CCP, by = c("school_name" = "School Name")) %>%
  left_join(CTE, by = c("school_code" = "agency_code")) %>%
  left_join(enroll, by = c("school_code" = "agency_code"))
# View(merged_data)

# IB cleaned data
ib_regression_data <- merged_data[c("school_name", "spg_score", "pct_ib_participation", "pct_enrolled", "lea_name", "sbe_region")]
ib_cleaned_data <- na.omit(ib_regression_data)
# View(ib_cleaned_data)

# AP cleaned data
ap_regression_data <- merged_data[c("school_name", "spg_score", "pct_ap_participation", "pct_enrolled", "lea_name", "sbe_region")]
ap_cleaned_data <- na.omit(ap_regression_data)
# View(ap_cleaned_data)

# CIE cleaned data
cie_regression_data <- merged_data[c("school_name", "spg_score", "pct_cie_participation", "pct_enrolled", "lea_name", "sbe_region")]
cie_cleaned_data <- na.omit(cie_regression_data)
cie_cleaned_data$pct_cie_participation <- as.numeric(cie_cleaned_data$pct_cie_participation)
cie_cleaned_data$pct_enrolled <- as.numeric(cie_cleaned_data$pct_enrolled)
# View(cie_cleaned_data)

# CTE cleaned data
cte_regression_data <- merged_data[c("school_name", "spg_score", "pct", "pct_enrolled", "lea_name", "sbe_region")]
cte_cleaned_data <- na.omit(cte_regression_data)
cte_cleaned_data$pct <- cte_cleaned_data$pct*.01
# View(cte_cleaned_data)

# CCP (dual enrollment/college transfer and early & middle colleges/cooperative innovative high schools) cleaned data
ccp_regression_data <- merged_data[c("school_name", "spg_score", "cohortgradpct", "pct_enrolled", "lea_name", "sbe_region")]
ccp_cleaned_data <- na.omit(ccp_regression_data)
ccp_cleaned_data$cohortgradpct <- as.numeric(ccp_cleaned_data$cohortgradpct)
# View(ccp_cleaned_data)

# Renames program participation to 'pro_enroll' and creates 'program' variable
names(ap_cleaned_data)[names(ap_cleaned_data) == "pct_ap_participation"] <- "pro_enroll"
names(ib_cleaned_data)[names(ib_cleaned_data) == "pct_ib_participation"] <- "pro_enroll"
names(cie_cleaned_data)[names(cie_cleaned_data) == "pct_cie_participation"] <- "pro_enroll"
names(cte_cleaned_data)[names(cte_cleaned_data) == "pct"] <- "pro_enroll"
names(ccp_cleaned_data)[names(ccp_cleaned_data) == "cohortgradpct"] <- "pro_enroll"

ap_cleaned_data$program <- "AP"
ib_cleaned_data$program <- "IB"
cie_cleaned_data$program <- "CIE"
cte_cleaned_data$program <- "CTE"
ccp_cleaned_data$program <- "CCP"

# Combines cleaned program-specific data
combined_data <- rbind(
  ap_cleaned_data,
  ib_cleaned_data,
  cie_cleaned_data,
  cte_cleaned_data,
  ccp_cleaned_data)
View(combined_data)

# NCStar "causative" case comparison
combined_data$NCStar <- as.vector(0) # NCStar-selected school districts
combined_data$NCStar[combined_data$lea_name == "Jackson County Schools" | combined_data$lea_name == "Union County Public Schools" | combined_data$lea_name == "Wilkes County Schools" | combined_data$lea_name == "Richmond County Schools" | combined_data$lea_name == "Brunswick County Schools" | combined_data$lea_name == "Randolph County Schools" | combined_data$lea_name == "Wake County Schools" | combined_data$lea_name == "Beaufort County Schools"] <- 1
combined_data$case <- as.vector(NA)
combined_data$case[combined_data$lea_name == "Charlotte-Mecklenburg Schools"] <- 0
combined_data$case[combined_data$lea_name == "Union County Public Schools"] <- 1

# Removes duplicate school entries
duplicates <- any(duplicated(combined_data[c("school_name", "lea_name")]))
if (duplicates) {
  combined_data <- combined_data[!duplicated(combined_data[c("school_name", "lea_name")]), ]
}
View(combined_data)

model_1 <- lm(pct_enrolled ~ NCStar, data = combined_data) # Linear regression analysis for postsecondary enrollment rates in low-performing schools receiving NCStar improvement training. 
summary(model_1)
model_2 <- lm(pct_enrolled ~ case, data=combined_data) # Linear regression comparing postsecondary enrollment rates between low-performing Union County Public Schools (with NCStar training) and Charlotte-Mecklenburg Schools (without NCStar training).
summary(model_2)

summary_model_1 <- summary(model_1)
summary_model_2 <- summary(model_2)

stargazer( # Regression table
  model_1, model_2,
  title = "Regression Analysis for Postsecondary Enrollment from Low Performing Secondary Schools",
  align = TRUE,
  dep.var.labels = "Postsecondary Enrollment (%)",
  covariate.labels = c("NCStar Training", "Union vs. CMS", "Intercept"),
  omit = "f",
  type = "html",
  out = "regression_table.html",
  out.header = TRUE,
  column.labels = c("<b>Model 1</b>", "<b>Model 2</b>"),  
  model.names = TRUE,
  star.cutoffs = c(0.05, 0.01, 0.001),
  add.lines = list(c("<b>Model Fitness</b>", "<b>Model 1</b>", "<b>Model 2</b>")),
  single.row = TRUE,  
  digits = 3
)
html_text <- '<table style="text-align:center"><caption><strong>Regression Analysis for Postsecondary Enrollment from Low Performing Secondary Schools</strong></caption>
<tr><td colspan="3" style="border-bottom: 1px solid black"></td></tr><tr><td style="text-align:left"></td><td colspan="2"><em>Dependent variable:</em></td></tr>
<tr><td></td><td colspan="2" style="border-bottom: 1px solid black"></td></tr>
<tr><td style="text-align:left"></td><td colspan="2">Postsecondary Enrollment (%)</td></tr>
<tr><td style="text-align:left"></td><td colspan="2"><em>OLS</em></td></tr>
<tr><td style="text-align:left"></td><td><b>Model 1</b></td><td><b>Model 2</b></td></tr>
<tr><td style="text-align:left"></td><td>(1)</td><td>(2)</td></tr>
<tr><td colspan="3" style="border-bottom: 1px solid black"></td></tr><tr><td style="text-align:left">NCStar Training</td><td></td><td>0.031 (0.120)</td></tr>
<tr><td colspan="3" style="border-bottom: 1px solid black"></td></tr><tr><td style="text-align:left"><b>Model Fitness</b></td><td><b>Model 1</b></td><td><b>Model 2</b></td></tr>
<tr><td style="text-align:left">Observations</td><td>93</td><td>8</td></tr>
<tr><td style="text-align:left">R<sup>2</sup></td><td>0.009</td><td>0.011</td></tr>
<tr><td style="text-align:left">Adjusted R<sup>2</sup></td><td>-0.002</td><td>-0.154</td></tr>
<tr><td style="text-align:left">Residual Std. Error</td><td>0.132 (df = 91)</td><td>0.112 (df = 6)</td></tr>
<tr><td style="text-align:left">F Statistic</td><td>0.799 (df = 1; 91)</td><td>0.067 (df = 1; 6)</td></tr>
<tr><td colspan="3" style="border-bottom: 1px solid black"></td></tr><tr><td style="text-align:left"><em>Note:</em></td><td colspan="2" style="text-align:right"><sup>*</sup>p<0.05; <sup>**</sup>p<0.01; <sup>***</sup>p<0.001</td></tr>
</table>'
htmltools::HTML(html_text)
browseURL("regression_table.html")

# Average program enrollment at low performing (D or F) NC secondary schools bar graph
program_means <- combined_data %>%
  group_by(program) %>%
  summarise_all(list(mean = ~mean(.)*100))
View(program_means)

font_theme <- theme_minimal(base_size = 12) +
  theme(
    text = element_text(family = "Tahoma"),
    axis.text.x = element_text(color = "black", angle = 0, hjust = 0.5, face = "bold", size = 18),
    axis.title.x = element_text(color = "black", face = "bold", size = 18),
    axis.text.y = element_text(color = "black", face = "bold", size = 18),
    axis.title.y = element_text(color = "black", face = "bold", size = 18),
    plot.title = element_text(face = "bold", hjust = 0.5, size = 20),
    legend.text = element_text(face = "bold", size = 12),
    legend.title = element_text(face = "bold", size = 16),
    panel.grid.major = element_blank(),
    panel.grid.minor = element_blank(),
    panel.background = element_rect(fill = "#D7D7D7"),
    legend.position = "top",
    legend.background = element_rect(fill = "#D7D7D7"),
    panel.border = element_rect(color = "black", fill = NA, size = 1),
    plot.margin = unit(c(1, 1, 1, 1), "cm")
  )
ggplot(program_means, aes(x = program, y = pro_enroll_mean, fill = program)) +
  geom_bar(stat = "identity", position = "stack", width = 0.7, color = "grey", linetype = "solid") +
  geom_text(aes(label = sprintf("%.1f%%", pro_enroll_mean), family = "Tahoma"), position = position_stack(vjust = 0.5), size = 8, color = "black") +
  labs(x = "Program", y = "Mean Enrollment (%)", fill = "Program") +
  ggtitle("Average Student Enrollment in Early Postsecondary Programs at Low-Performing Secondary Schools (2022-2023)") +
  font_theme +
  scale_fill_manual(values = c("AP" = "#43D27B", "IB" = "#3F76D0", "CIE" = "#43D2C5", "CCP" = "#43D2A7", "CTE" = "#439CD2")) +
  labs(fill = "Program") +
  scale_y_continuous(labels = scales::percent_format(scale = 1), limits = c(0, 100))

