# Install necessary packages if not already installed
install.packages("tidyverse")
install.packages("readxl")
install.packages("skimr")
install.packages("GGally")
install.packages("readxlsb")
install.packages("ggcorrplot")
install.packages("VIM") 
install.packages("randomForest")
install.packages("caTools")
install.packages("caret")
install.packages("nnet")

# Load necessary libraries
library(tidyverse)  
library(skimr)     
library(GGally)     
library(readxlsb)   
library(readxl)
library(ggcorrplot)
library(VIM) 
library(randomForest)
library(caTools) 
library(caret) 
library(shiny)
library(nnet)
library(dplyr)

# Load and clean data
file_path <- "C:/Users/angad/Downloads/Data.xlsb"
data <- read_xlsb(file_path, sheet = 1, range = "A1:Z1000")

# Initial Data Exploration
head(data)
str(data)
summary(data)
sapply(data, function(x) sum(is.na(x))) # Check for missing values
skim(data) # Overview of the data

# Data cleaning and transformation
data_clean <- data %>%
  mutate(PPC = ifelse(is.na(PPC), mean(PPC, na.rm = TRUE), PPC),
         Prosperity = ifelse(is.na(Prosperity), mean(Prosperity, na.rm = TRUE), Prosperity),
         PROD_NAME = as.factor(PROD_NAME),
         PH_OCC = as.factor(PH_OCC),
         PH_EDUCATION = as.factor(PH_EDUCATION),
         INCOME_SEGMENT = as.factor(INCOME_SEGMENT),
         PROD_CATEGORY = as.factor(PROD_CATEGORY))

# Splitting the dataset into training and testing sets
set.seed(123)  # For reproducibility
split <- sample.split(data_clean$PROD_NAME, SplitRatio = 0.7)
train_data <- subset(data_clean, split == TRUE)
test_data <- subset(data_clean, split == FALSE)

train_data <- na.omit(train_data)
test_data <- na.omit(test_data)

train_data <- train_data %>%
  filter(PPC < quantile(PPC, 0.99))  # Remove top 1% outliers

rf_model_all_features <- randomForest(PROD_NAME ~ PH_OCC + PH_EDUCATION + INCOME_SEGMENT + PPC + Prosperity +
                                        PH_PROS_BAND + INS_QS + PROD_CATEGORY + PRODUCT_VARIANT,
                                      data = train_data, 
                                      ntree = 100, 
                                      importance = TRUE)

# Feature importance check
importance(rf_model_all_features)
varImpPlot(rf_model_all_features)

# Evaluate the model with all features
rf_all_features_predictions <- predict(rf_model_all_features, newdata = test_data)
confusionMatrix(rf_all_features_predictions, test_data$PROD_NAME)

# Drop less important features based on the new selection
rf_model_refined <- randomForest(PROD_NAME ~ PH_EDUCATION + PH_OCC + INS_QS + PPC + Prosperity,  # Refined set of features
                                 data = train_data, 
                                 ntree = 100, 
                                 importance = TRUE)

# Check feature importance for the refined model
importance(rf_model_refined)
varImpPlot(rf_model_refined)

# Evaluate the refined model
rf_refined_predictions <- predict(rf_model_refined, newdata = test_data)
confusionMatrix(rf_refined_predictions, test_data$PROD_NAME)

# --- HYPERPARAMETER TUNING FOR RANDOM FOREST ---

# Set up a grid of hyperparameters to tune
tune_grid <- expand.grid(mtry = c(2, 3, 4, 5))  # Number of variables tried at each split

# Set up control for cross-validation
control <- trainControl(method = "cv", number = 5)  # 5-fold cross-validation

# Train Random Forest model with hyperparameter tuning
rf_tuned_model <- train(PROD_NAME ~ PH_EDUCATION + PH_OCC + INS_QS + PPC + Prosperity,
                        data = train_data,
                        method = "rf",
                        tuneGrid = tune_grid,
                        ntree = 100,  # Number of trees in the forest
                        trControl = control)

# Check the best tuned model
print(rf_tuned_model)

# Evaluate the tuned model
rf_tuned_predictions <- predict(rf_tuned_model, newdata = test_data)
rf_tuned_conf_matrix <- confusionMatrix(rf_tuned_predictions, test_data$PROD_NAME)

# Print tuned accuracy
rf_tuned_accuracy <- rf_tuned_conf_matrix$overall["Accuracy"]
print(paste("Tuned Random Forest Model Accuracy:", round(rf_tuned_accuracy * 100, 2), "%"))

# --- MODEL COMPARISON SECTION (Random Forest, KNN, Decision Tree) ---

# 1. Random Forest (tuned)
rf_accuracy <- rf_tuned_accuracy

# 2. K-Nearest Neighbors (KNN)
knn_model <- train(PROD_NAME ~ PH_EDUCATION + PH_OCC + INS_QS + PPC + Prosperity,
                   data = train_data, 
                   method = "knn", 
                   tuneLength = 10)

knn_predictions <- predict(knn_model, newdata = test_data)
knn_accuracy <- confusionMatrix(knn_predictions, test_data$PROD_NAME)$overall["Accuracy"]

# 3. Decision Tree
dt_model <- train(PROD_NAME ~ PH_EDUCATION + PH_OCC + INS_QS + PPC + Prosperity,
                  data = train_data, 
                  method = "rpart")

dt_predictions <- predict(dt_model, newdata = test_data)
dt_accuracy <- confusionMatrix(dt_predictions, test_data$PROD_NAME)$overall["Accuracy"]

# Compare the models' accuracy
model_comparison <- data.frame(
  Model = c("Tuned Random Forest", "KNN", "Decision Tree"),
  Accuracy = c(rf_accuracy, knn_accuracy, dt_accuracy)
)

print("Model Comparison based on Accuracy:")
print(model_comparison)

# ---- Shiny App ----

ui <- fluidPage(
  
  # Application title
  titlePanel("Insurance Policy Recommendation System"),
  
  # Sidebar layout
  sidebarLayout(
    sidebarPanel(
      selectInput("PH_OCC", "Occupation:", choices = levels(train_data$PH_OCC)),
      selectInput("PH_EDUCATION", "Education Level:", choices = levels(train_data$PH_EDUCATION)),
      numericInput("INS_QS", "Insurance QS Score:", value = 25, min = 0, max = 50, step = 1),
      numericInput("PPC", "PPC (Policyholder Premium Contribution):", value = 100000, min = 0, step = 10000),
      numericInput("Prosperity", "Prosperity Score:", value = 50, min = 0, max = 100, step = 1),
      actionButton("recommend", "Get Recommendation")
    ),
    
    mainPanel(
      h3("Recommended Product:"),
      verbatimTextOutput("recommended_product"),
      
      h3("Feature Importance:"),
      plotOutput("feature_importance")
    )
  )
)

# Define server logic 
server <- function(input, output) {
  
  observeEvent(input$recommend, {
    customer_profile <- data.frame(PH_OCC = factor(input$PH_OCC, levels = levels(train_data$PH_OCC)),
                                   PH_EDUCATION = factor(input$PH_EDUCATION, levels = levels(train_data$PH_EDUCATION)),
                                   INS_QS = input$INS_QS,
                                   PPC = input$PPC,
                                   Prosperity = input$Prosperity)
    
    predicted_product <- predict(rf_tuned_model, newdata = customer_profile)
    
    output$recommended_product <- renderText({
      paste("The recommended insurance policy is:", predicted_product)
    })
    
    output$feature_importance <- renderPlot({
      varImpPlot(rf_tuned_model$finalModel)
    })
  })
}

# Run the Shiny app
shinyApp(ui = ui, server = server)

#Accuracy of the final model
# Step 1: Make predictions on the test data using the final Random Forest model
rf_final_predictions <- predict(rf_tuned_model, newdata = test_data)

# Step 2: Compare predictions with actual data using confusion matrix
rf_conf_matrix <- confusionMatrix(rf_final_predictions, test_data$PROD_NAME)

# Step 3: Print the confusion matrix and accuracy metrics
print("Confusion Matrix for the Tuned Random Forest Model:")
print(rf_conf_matrix)

# Step 4: Extract and print the accuracy
rf_final_accuracy <- rf_conf_matrix$overall["Accuracy"]
print(paste("Tuned Model Final Accuracy:", round(rf_final_accuracy * 100, 2), "%"))