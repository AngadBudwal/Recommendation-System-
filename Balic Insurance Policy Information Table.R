install.packages("shiny", dependencies = TRUE)
install.packages("fastmap")
install.packages("RSQLite")
install.packages("dplyr")


library(dplyr)

# Initialize the data frame with a Serial_Number column
insurance_plans <- data.frame(
  Serial_Number = integer(),
  Plan_Name = character(),
  Type = character(),
  Variants = character(),
  Key_Benefits = character(),
  Min_Age_At_Entry = numeric(),
  Max_Age_At_Entry = numeric(),
  Min_Policy_Term = numeric(),
  Max_Policy_Term = numeric(),
  Min_Premium_Payment_Term = numeric(),
  Max_Premium_Payment_Term = numeric(),
  Sum_Assured = numeric(),
  Exclusions = character(),
  Plan_Type = character(),
  stringsAsFactors = FALSE
)

# Example of adding a policy with manual entries
insurance_plans <- rbind(insurance_plans, data.frame(
  Serial_Number = 1,
  Plan_Name = "Bajaj Allianz Life eTouch",
  Type = "Non-Linked, Non-Participating",
  Variants = "Life Shield, Life Shield Plus, Life Shield ROP",
  Key_Benefits = "Death Benefit, Terminal Illness Benefit, Accidental Death Benefit, Return of Premiums, Waiver of Premiums",
  Min_Age_At_Entry = 18,
  Max_Age_At_Entry = 65,
  Min_Policy_Term = 85,
  Max_Policy_Term = 99,
  Min_Premium_Payment_Term = 5,  
  Max_Premium_Payment_Term = 25, 
  Sum_Assured = NA,  
  Exclusions = "Suicide Clause, certain accidental death and disability exclusions",
  Plan_Type = "Term Insurance",
  stringsAsFactors = FALSE
))

insurance_plans <- rbind(insurance_plans, data.frame(
  Serial_Number = 2,
  Plan_Name = "Bajaj Allianz Life Diabetic Term Plan II Sub 8 HbA1c",
  Type = "Non-Linked, Non-Participating",
  Variants = "N/A",
  Key_Benefits = "Death Benefit, Health Management Services, Premium Payment Flexibility, Tax Benefits",
  Min_Age_At_Entry = 30,
  Max_Age_At_Entry = 60,
  Min_Policy_Term = 5,
  Max_Policy_Term = 25,
  Min_Premium_Payment_Term = 5,  
  Max_Premium_Payment_Term = 25, 
  Sum_Assured = 2500000,  
  Exclusions = "Suicide Clause, no maturity benefit, no surrender value",
  Plan_Type = "Term Insurance",
  stringsAsFactors = FALSE
))

insurance_plans <- rbind(insurance_plans, data.frame(
  Serial_Number = 3,
  Plan_Name = "Bajaj Allianz Life iSecure",
  Type = "Non-Linked, Non-Participating",
  Variants = "N/A",
  Key_Benefits = "Death Benefit, Joint Life Cover, Premium Rebates, Rider Options",
  Min_Age_At_Entry = 18,
  Max_Age_At_Entry = 65,
  Min_Policy_Term = 10,
  Max_Policy_Term = 85, 
  Min_Premium_Payment_Term = 5,  
  Max_Premium_Payment_Term = 25, 
  Sum_Assured = 2500000, 
  Exclusions = "Suicide Clause, no maturity benefit, surrender value only for limited premium policies",
  Plan_Type = "Term Insurance",
  stringsAsFactors = FALSE
))

insurance_plans <- rbind(insurance_plans, data.frame(
  Serial_Number = 4,
  Plan_Name = "Bajaj Allianz Life Smart Protection Goal",
  Type = "Non-Linked, Non-Participating",
  Variants = "N/A",
  Key_Benefits = "Death Benefit, Maturity Benefit (with ROP), Critical Illness Benefit, Whole Life Cover",
  Min_Age_At_Entry = 30,
  Max_Age_At_Entry = 60,
  Min_Policy_Term = 5,
  Max_Policy_Term = 25,
  Min_Premium_Payment_Term = 5,  
  Max_Premium_Payment_Term = 25, 
  Sum_Assured = 2500000, 
  Exclusions = "Suicide Clause, no maturity benefit without ROP, no surrender value for regular premium policies",
  Plan_Type = "Term Insurance",
  stringsAsFactors = FALSE
))


insurance_plans <- rbind(insurance_plans, data.frame(
  Serial_Number = 5,
  Plan_Name = "Bajaj Allianz Life Saral Jeevan Bima",
  Type = "Non-Linked, Non-Participating",
  Variants = "N/A",
  Key_Benefits = "Death Benefit, Accidental Death Benefit during the initial 45-day waiting period",
  Min_Age_At_Entry = 30,
  Max_Age_At_Entry = 60,
  Min_Policy_Term = 5,
  Max_Policy_Term = 25,
  Min_Premium_Payment_Term = 5, 
  Max_Premium_Payment_Term = 25, 
  Sum_Assured = 2500000,  
  Exclusions = "Suicide Clause, 45-day waiting period",
  Plan_Type = "Term Insurance",
  stringsAsFactors = FALSE
))

insurance_plans <- rbind(insurance_plans, data.frame(
  Serial_Number = 6,
  Plan_Name = "Bajaj Allianz Life Goal Assure II",
  Type = "Unit-Linked, Non-Participating",
  Variants = "N/A",
  Key_Benefits = "Death Benefit, Maturity Benefit, Fund Booster, Return of Mortality Charges, Loyalty Additions, Multiple Investment Strategies",
  Min_Age_At_Entry = 30,
  Max_Age_At_Entry = 60,
  Min_Policy_Term = 5,
  Max_Policy_Term = 20,  
  Min_Premium_Payment_Term = 5,  
  Max_Premium_Payment_Term = 20,
  Sum_Assured = 7,  # Minimum 7 times the annualized premium; Adjust based on age and policy term
  Exclusions = "Suicide Clause, Market Risks",
  Plan_Type = "ULIP",
  stringsAsFactors = FALSE
))

insurance_plans <- rbind(insurance_plans, data.frame(
  Serial_Number = 7,
  Plan_Name = "Bajaj Allianz Life Smart Wealth Goal III",
  Type = "Unit-Linked, Non-Participating",
  Variants = "N/A",
  Key_Benefits = "Death Benefit, Maturity Benefit, Fund Booster, Return of Mortality Charges, Multiple Investment Strategies, Loyalty Benefits, Flexibility Options",
  Min_Age_At_Entry = 0,   
  Max_Age_At_Entry = 60,  
  Min_Policy_Term = 10,   
  Max_Policy_Term = 60,
  Min_Premium_Payment_Term = 5,  
  Max_Premium_Payment_Term = 60, 
  Sum_Assured = 7,  # Minimum 7 times the annualized premium for regular/limited premium; Minimum 1.25 times the single premium for single premium
  Exclusions = "Suicide Clause, Market Risks",
  Plan_Type = "ULIP",
  stringsAsFactors = FALSE
))

insurance_plans <- rbind(insurance_plans, data.frame(
  Serial_Number = 8,
  Plan_Name = "Bajaj Allianz Life Magnum Fortune Plus II",
  Type = "Unit-Linked, Non-Participating",
  Variants = "N/A",
  Key_Benefits = "Death Benefit, Maturity Benefit, Loyalty Additions, Maturity Booster, Return of Mortality Charges, Family Benefit, Multiple Investment Strategies, Flexibility Options",
  Min_Age_At_Entry = 0,   
  Max_Age_At_Entry = 65,
  Min_Policy_Term = 10,    
  Max_Policy_Term = 30,
  Min_Premium_Payment_Term = 5,  
  Max_Premium_Payment_Term = 25,
  Sum_Assured = 7,  # Minimum 7 times the annualized premium; Adjust based on the policyholder’s age and underwriting policy
  Exclusions = "Suicide Clause, Market Risks",
  Plan_Type = "ULIP",
  stringsAsFactors = FALSE
))

insurance_plans <- rbind(insurance_plans, data.frame(
  Serial_Number = 9,
  Plan_Name = "Bajaj Allianz Life Goal Based Saving II",
  Type = "Unit-Linked, Non-Participating",
  Variants = "N/A",
  Key_Benefits = "Death Benefit, Maturity Benefit, Additional Allocation, Return of Mortality Charges, Multiple Investment Options, Liquidity Options",
  Min_Age_At_Entry = 30,
  Max_Age_At_Entry = 60,
  Min_Policy_Term = 5,   
  Max_Policy_Term = 25,
  Min_Premium_Payment_Term = 5,  
  Max_Premium_Payment_Term = 25,
  Sum_Assured = 2500000,  # Minimum ₹25,00,000; No upper limit, subject to the company's underwriting policy
  Exclusions = "Suicide Clause, Market Risks",
  Plan_Type = "ULIP",
  stringsAsFactors = FALSE
))

insurance_plans <- rbind(insurance_plans, data.frame(
  Serial_Number = 10,
  Plan_Name = "Bajaj Allianz Life LongLife Goal",
  Type = "Unit-Linked, Non-Participating",
  Variants = "N/A",
  Key_Benefits = "Death Benefit, Maturity Benefit, Retired Life Income, Periodical Return of Charges, Loyalty Additions, Multiple Investment Options, Waiver of Premium",
  Min_Age_At_Entry = 30,
  Max_Age_At_Entry = 60,
  Min_Policy_Term = 30,  
  Max_Policy_Term = 99,
  Min_Premium_Payment_Term = 5,   # Limited pay options available from 5 years up to 25 years
  Max_Premium_Payment_Term = 25,
  Sum_Assured = 10,  # Minimum 10 times the annualized premium; Adjust based on the policyholder’s age and underwriting policy
  Exclusions = "Suicide Clause, Market Risks, Rider Exclusions",
  Plan_Type = "ULIP",
  stringsAsFactors = FALSE
))

insurance_plans <- rbind(insurance_plans, data.frame(
  Serial_Number = 11,
  Plan_Name = "Bajaj Allianz Life Future Gain II",
  Type = "Unit-Linked, Non-Participating",
  Variants = "N/A",
  Key_Benefits = "Death Benefit, Maturity Benefit, Investment Options, Partial Withdrawals, Loyalty Additions, Settlement Option",
  Min_Age_At_Entry = 1,
  Max_Age_At_Entry = 60,
  Min_Policy_Term = 10,
  Max_Policy_Term = 30,
  Min_Premium_Payment_Term = 5,
  Max_Premium_Payment_Term = 30,
  Sum_Assured = 7,  # Minimum 7 times the annualized premium
  Exclusions = "Suicide Clause, Market Risks",
  Plan_Type = "ULIP",
  stringsAsFactors = FALSE
))

insurance_plans <- rbind(insurance_plans, data.frame(
  Serial_Number = 12,
  Plan_Name = "Bajaj Allianz Life Invest Protect Goal",
  Type = "Unit-Linked, Non-Participating",
  Variants = "N/A",
  Key_Benefits = "Death Benefit, Maturity Benefit, Loyalty Additions, Return of Charges, Fund Maintenance Booster, Investment Strategies",
  Min_Age_At_Entry = 18,
  Max_Age_At_Entry = 60,
  Min_Policy_Term = 20,
  Max_Policy_Term = 40,
  Min_Premium_Payment_Term = 5,
  Max_Premium_Payment_Term = 12,  # Limited pay options from 5 to 12 years, or regular pay for the entire policy term
  Sum_Assured = 7,  # Minimum 7 times the annualized premium
  Exclusions = "Suicide Clause, Market Risks",
  Plan_Type = "ULIP",
  stringsAsFactors = FALSE
))

insurance_plans <- rbind(insurance_plans, data.frame(
  Serial_Number = 13,
  Plan_Name = "Bajaj Allianz Life Fortune Gain",
  Type = "Unit-Linked, Non-Participating",
  Variants = "N/A",
  Key_Benefits = "Death Benefit, Maturity Benefit, Partial Withdrawals, Investment Options, Top-Up Premiums, Settlement Option",
  Min_Age_At_Entry = 1,
  Max_Age_At_Entry = 63,
  Min_Policy_Term = 7,
  Max_Policy_Term = 30,
  Min_Premium_Payment_Term = 1,  
  Max_Premium_Payment_Term = 1, 
  Sum_Assured = 7,  # Minimum 7 times the annualized premium
  Exclusions = "Suicide Clause, Market Risks",
  Plan_Type = "ULIP",
  stringsAsFactors = FALSE
))

insurance_plans <- rbind(insurance_plans, data.frame(
  Serial_Number = 14,
  Plan_Name = "Bajaj Allianz Life Assured Wealth Goal Platinum",
  Type = "Non-Linked, Non-Participating",
  Variants = "N/A",
  Key_Benefits = "Life Cover, Guaranteed Income, Return of Premium (ROP), Flexible Options, Maturity Benefit",
  Min_Age_At_Entry = 6,
  Max_Age_At_Entry = 60,
  Min_Policy_Term = 5,
  Max_Policy_Term = 17,
  Min_Premium_Payment_Term = 5,
  Max_Premium_Payment_Term = 12,
  Sum_Assured = NA,  
  Exclusions = "Suicide Clause, Surrender Value",
  Plan_Type = "Savings Plan",
  stringsAsFactors = FALSE
))

insurance_plans <- rbind(insurance_plans, data.frame(
  Serial_Number = 15,
  Plan_Name = "Bajaj Allianz Life ACE",
  Type = "Non-Linked, Participating",
  Variants = "N/A",
  Key_Benefits = "Flexibility, Early Income, Wealth, Increasing Income, Protection, Maturity Benefit",
  Min_Age_At_Entry = 18,
  Max_Age_At_Entry = 55,
  Min_Policy_Term = 10,
  Max_Policy_Term = 45,
  Min_Premium_Payment_Term = 5,
  Max_Premium_Payment_Term = 12,
  Sum_Assured = 11,  
  Exclusions = "Suicide Clause, Market Risks",
  Plan_Type = "Savings Plan",
  stringsAsFactors = FALSE
))

insurance_plans <- rbind(insurance_plans, data.frame(
  Serial_Number = 16,
  Plan_Name = "Bajaj Allianz Life Flexi Income Goal",
  Type = "Non-Linked, Participating",
  Variants = "N/A",
  Key_Benefits = "Guaranteed Income, Cash Bonus, Maturity Benefit, Death Benefit, Joint Life Cover",
  Min_Age_At_Entry = 0,  
  Max_Age_At_Entry = 55,
  Min_Policy_Term = NA,  
  Max_Policy_Term = NA,  
  Min_Premium_Payment_Term = 5,
  Max_Premium_Payment_Term = 20,
  Sum_Assured = NA,  # Varies based on the chosen sum assured and premium payment term
  Exclusions = "Suicide Clause, Nonforfeiture Benefits",
  Plan_Type = "Savings Plan",
  stringsAsFactors = FALSE
))

insurance_plans <- rbind(insurance_plans, data.frame(
  Serial_Number = 17,
  Plan_Name = "Bajaj Allianz Life Assured Wealth Goal",
  Type = "Non-Linked, Non-Participating",
  Variants = "N/A",
  Key_Benefits = "Life Cover, Guaranteed Income, Return of Premium, Flexible Premium Payment, Multiple Variants, Maturity Benefit",
  Min_Age_At_Entry = 18,
  Max_Age_At_Entry = 75,
  Min_Policy_Term = 9,
  Max_Policy_Term = 99,
  Min_Premium_Payment_Term = 5,
  Max_Premium_Payment_Term = NA,  
  Sum_Assured = NA,  
  Exclusions = "Suicide Clause, Surrender and Paid-Up Policy",
  Plan_Type = "Savings Plan",
  stringsAsFactors = FALSE
))

insurance_plans <- rbind(insurance_plans, data.frame(
  Serial_Number = 18,
  Plan_Name = "Bajaj Allianz Life Guaranteed Savings Goal",
  Type = "Non-Linked, Non-Participating",
  Variants = "N/A",
  Key_Benefits = "Guaranteed Maturity Benefit, Higher Maturity Benefit for Females, Death Benefit",
  Min_Age_At_Entry = 0,
  Max_Age_At_Entry = 60,  
  Min_Policy_Term = 5,
  Max_Policy_Term = 20,
  Min_Premium_Payment_Term = 1,  
  Max_Premium_Payment_Term = 1,  
  Sum_Assured = 1.5,  
  Exclusions = "Suicide Clause, Surrender Benefits",
  Plan_Type = "Savings Plan",
  stringsAsFactors = FALSE
))

insurance_plans <- rbind(insurance_plans, data.frame(
  Serial_Number = 19,
  Plan_Name = "Bajaj Allianz Life POS Goal Suraksha",
  Type = "Non-Linked, Non-Participating",
  Variants = "N/A",
  Key_Benefits = "Guaranteed Maturity Benefit, Death Benefit, Guaranteed Additions, Loan Facility",
  Min_Age_At_Entry = 0,
  Max_Age_At_Entry = 55,
  Min_Policy_Term = 10,
  Max_Policy_Term = 20,
  Min_Premium_Payment_Term = 5,
  Max_Premium_Payment_Term = 12,
  Sum_Assured = 30000,  # Minimum ₹30,000; Maximum ₹25,00,000
  Exclusions = "Suicide Clause, Waiting Period, Policy Lapse and Paid-Up",
  Plan_Type = "Savings Plan",
  stringsAsFactors = FALSE
))

insurance_plans <- rbind(insurance_plans, data.frame(
  Serial_Number = 20,
  Plan_Name = "Bajaj Allianz Life Guaranteed Income Goal",
  Type = "Non-Linked, Non-Participating",
  Variants = "N/A",
  Key_Benefits = "Guaranteed Maturity Benefit, Income Benefit, Lump-Sum Benefit, Extended Life Cover, Death Benefit, Riders",
  Min_Age_At_Entry = 6,
  Max_Age_At_Entry = 60,
  Min_Policy_Term = 5,
  Max_Policy_Term = 20,
  Min_Premium_Payment_Term = 5,
  Max_Premium_Payment_Term = 12,
  Sum_Assured = 100000,  
  Exclusions = "Suicide Clause, Waiting Period",
  Plan_Type = "Savings Plan",
  stringsAsFactors = FALSE
))

insurance_plans <- rbind(insurance_plans, data.frame(
  Serial_Number = 21,
  Plan_Name = "Bajaj Allianz Life Assured Wealth Goal",
  Type = "Non-Linked, Non-Participating Individual Life Insurance Savings Plan",
  Variants = "Single Life, Joint Life",
  Key_Benefits = "Life Cover, Guaranteed Income, Return of Premium, Flexible Premium Payment, Multiple Variants, Maturity Benefit",
  Min_Age_At_Entry = 18,
  Max_Age_At_Entry = 75,
  Min_Policy_Term = 9,
  Max_Policy_Term = 99,
  Min_Premium_Payment_Term = 5,
  Max_Premium_Payment_Term = 99, 
  Sum_Assured = NA, 
  Exclusions = "Suicide Clause, Surrender and Paid-Up Policy",
  Plan_Type = "Child Insurance Plan",
  stringsAsFactors = FALSE
))

insurance_plans <- rbind(insurance_plans, data.frame(
  Serial_Number = 22,
  Plan_Name = "Bajaj Allianz Life Saral Pension",
  Type = "Non-Linked, Non-Participating Individual Immediate Annuity Plan",
  Variants = "Immediate Annuity, Joint Life Annuity",
  Key_Benefits = "Lifelong Guaranteed Income, Return of Purchase Price, Joint Life Option, Flexible Annuity Payouts, Liquidity",
  Min_Age_At_Entry = 40,
  Max_Age_At_Entry = 80,
  Min_Policy_Term = NA, 
  Max_Policy_Term = NA,
  Min_Premium_Payment_Term = NA, 
  Max_Premium_Payment_Term = NA,
  Sum_Assured = NA, 
  Exclusions = "Suicide Clause, Surrender Value",
  Plan_Type = "Retirement Plan",
  stringsAsFactors = FALSE
))

insurance_plans <- rbind(insurance_plans, data.frame(
  Serial_Number = 23,
  Plan_Name = "Bajaj Allianz Life Guaranteed Pension Goal II",
  Type = "Non-Linked, Non-Participating Immediate or Deferred Annuity Plan",
  Variants = "Immediate Annuity, Deferred Annuity, Joint Life Annuity",
  Key_Benefits = "Annuity Options, Deferred Annuity, Return of Purchase Price, Joint Life Annuity, Survival Benefits",
  Min_Age_At_Entry = 30,
  Max_Age_At_Entry = 85, 
  Min_Policy_Term = NA, 
  Max_Policy_Term = NA,
  Min_Premium_Payment_Term = NA, 
  Max_Premium_Payment_Term = NA,
  Sum_Assured = NA, 
  Exclusions = "Suicide Clause, Market Risks",
  Plan_Type = "Retirement Plan",
  stringsAsFactors = FALSE
))


#print_plan <- function(plan) {
#  cat("Plan Name:", plan$Plan_Name, "\n")
#  cat("Type:", plan$Type, "\n")
#  cat("Variants:", plan$Variants, "\n")
#  cat("Key Benefits:", plan$Key_Benefits, "\n")
#  cat("Age at Entry:", plan$Age_At_Entry, "\n")
#  cat("Policy Term:", plan$Policy_Term, "\n")
#  cat("Premium Payment Term:", plan$Premium_Payment_Term, "\n")
#  cat("Sum Assured:", plan$Sum_Assured, "\n")
#  cat("Exclusions:", plan$Exclusions, "\n")
#  cat("\n") # Add an extra line for separation
#}
#apply(insurance_plans, 1, function(row) print_plan(as.list(row)))
insurance_plans <- insurance_plans %>%
  mutate(
    Sum_Assured_Fixed = ifelse(Sum_Assured > 1000000, Sum_Assured, NA),
    Sum_Assured_Multiple = ifelse(Sum_Assured <= 100, Sum_Assured, NA)
  )
sample_annual_premium <- 1000000

insurance_plans <- insurance_plans %>%
  mutate(
    Calculated_Sum_Assured = ifelse(!is.na(Sum_Assured_Multiple), 
                                    Sum_Assured_Multiple * sample_annual_premium, 
                                    Sum_Assured_Fixed)
  )
insurance_plans <- insurance_plans %>%
  mutate(
    Sum_Assured_Fixed = ifelse(Sum_Assured > 1000000, Sum_Assured, NA),
    Sum_Assured_Multiple = ifelse(Sum_Assured <= 100, Sum_Assured, NA)
  )

# Example: Calculate Sum Assured based on a sample annual premium
sample_annual_premium <- 1000000

insurance_plans <- insurance_plans %>%
  mutate(
    Calculated_Sum_Assured = ifelse(!is.na(Sum_Assured_Multiple), 
                                    Sum_Assured_Multiple * sample_annual_premium, 
                                    Sum_Assured_Fixed)
  )

insurance_plans$Non_Linked <- ifelse(grepl("Non-Linked", insurance_plans$Type), 1, 0)


insurance_plans$Non_Participating <- ifelse(grepl("Non-Participating", insurance_plans$Type), 1, 0)


insurance_plans$Sum_Assured <- gsub("[^0-9]", "", insurance_plans$Sum_Assured)
insurance_plans$Sum_Assured <- as.numeric(insurance_plans$Sum_Assured)

# Categorical data - Convert Type, Variants, Exclusions, Plan_Type to factors
insurance_plans$Type <- as.factor(insurance_plans$Type)
insurance_plans$Variants <- as.factor(insurance_plans$Variants)
insurance_plans$Exclusions <- as.factor(insurance_plans$Exclusions)
insurance_plans$Plan_Type <- as.factor(insurance_plans$Plan_Type)
insurance_plans$Non_Linked <- as.factor(insurance_plans$Non_Linked)
insurance_plans$Non_Participating <- as.factor(insurance_plans$Non_Participating)


# Example of splitting Key_Benefits into binary categorical columns
benefit_categories <- c("Death Benefit", "Accidental Death Benefit", "Terminal Illness Benefit", 
                        "Critical Illness Benefit", "Waiver of Premiums", "Return of Premiums", 
                        "Maturity Benefit", "Loyalty Additions", "Guaranteed Income", 
                        "Partial Withdrawals", "Investment Options", "Fund Booster", 
                        "Settlement Option", "Joint Life Cover", "Cash Bonus", 
                        "Return of Charges", "Extended Life Cover", "Riders")

exclusion_categories <- c("Suicide Clause", "Market Risks", "Rider Exclusions", 
                          "Accidental Death Exclusions", "Disability Exclusions", 
                          "45-day Waiting Period", "90-day Waiting Period", "No Maturity Benefit", 
                          "No Surrender Value", "Policy Lapse", "Paid-Up Policy", 
                          "Nonforfeiture Benefits", "Surrender Benefits")

for (benefit in benefit_categories) {
  insurance_plans[[benefit]] <- as.integer(grepl(benefit, insurance_plans$Key_Benefits))
}

for (exclusion in exclusion_categories) {
  insurance_plans[[exclusion]] <- as.integer(grepl(exclusion, insurance_plans$Exclusions))
}

insurance_plans$Exclusions <- NULL
insurance_plans$Type <- NULL

insurance_plans$Key_Benefits <- NULL


duplicates <- duplicated(insurance_plans)

print(insurance_plans[duplicates, ])

insurance_plans <- insurance_plans[!duplicated(insurance_plans), ]
print(insurance_plans)
transposed_plans <- t(insurance_plans)
print(transposed_plans)
transposed_plans <- as.data.frame(transposed_plans, stringsAsFactors = FALSE)

print_transposed_plans <- function(transposed_df) {
  plan_count <- ncol(transposed_df) 
  for (i in 1:nrow(transposed_df)) {
    cat(rownames(transposed_df)[i], ":\n")
    for (j in 1:plan_count) {
      cat(j, ") ", transposed_df[i, j], "\n", sep = "")
    }
    cat("\n") 
  }
}

print_transposed_plans(transposed_plans)

#TO FILTER
#filtered_plans <- filter(insurance_plans, grepl("Unit-Linked",Type))

#transposed_filtered_plans <- t(filtered_plans)
#print_transposed_plans(transposed_filtered_plans)

#UI

library(shiny)

library(DT)

ui <- fluidPage(
  titlePanel("Interactive Insurance Plans"),
  sidebarLayout(
    sidebarPanel(
      selectInput("planType", "Select Plan Type:", choices = unique(insurance_plans$Plan_Type))
    ),
    mainPanel(
      DTOutput("planTable")
    )
  )
)

server <- function(input, output) {
  output$planTable <- renderDT({
    datatable(filter(insurance_plans, Plan_Type == input$planType), options = list(pageLength = 5))
  })
}

shinyApp(ui = ui, server = server)
write.csv(insurance_plans, file = "insurance_plans.csv", row.names = FALSE)
write.csv(insurance_plans, file = "C:/Users/angad/OneDrive/Desktop/insurance_plans.csv", row.names = FALSE)


