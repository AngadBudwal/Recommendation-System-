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

# ---- FEATURE IMPORTANCE CHECK ----
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

# Refine model based on feature importance
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

# ---- HYPERPARAMETER TUNING FOR RANDOM FOREST ----

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

# ---- MODEL COMPARISON SECTION (Random Forest, KNN, Decision Tree) ----

# 1. Random Forest (tuned)
rf_accuracy <- rf_tuned_accuracy

# 2. K-Nearest Neighbors (KNN)
train_data_knn <- model.matrix(~ PH_OCC + PH_EDUCATION + INS_QS + PPC + Prosperity - 1, data = train_data) %>%
  as.data.frame()
test_data_knn <- model.matrix(~ PH_OCC + PH_EDUCATION + INS_QS + PPC + Prosperity - 1, data = test_data) %>%
  as.data.frame()

# Add PROD_NAME back to the one-hot encoded data
train_data_knn$PROD_NAME <- train_data$PROD_NAME
test_data_knn$PROD_NAME <- test_data$PROD_NAME

# Since model.matrix creates colnames with unwanted characters, clean up colnames
colnames(train_data_knn) <- make.names(colnames(train_data_knn))
colnames(test_data_knn) <- make.names(colnames(test_data_knn))

# KNN model with one-hot encoded data
knn_model <- train(PROD_NAME ~ .,
                   data = train_data_knn, 
                   method = "knn", 
                   tuneLength = 10)

# KNN predictions and accuracy
knn_predictions <- predict(knn_model, newdata = test_data_knn)
knn_accuracy <- confusionMatrix(knn_predictions, test_data_knn$PROD_NAME)$overall["Accuracy"]

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


# Step 2: Define function to check if predicted product belongs to the correct category
check_category_match <- function(predicted_product, actual_product, category_mapping) {
  # Map the predicted product to its category
  predicted_category <- category_mapping[predicted_product]
  
  # Map the actual product to its category
  actual_category <- category_mapping[actual_product]
  
  # Return TRUE if the predicted product belongs to the same category as the actual product
  return(predicted_category == actual_category)
}

# Step 3: Create a mapping of products to their categories
category_mapping <- train_data %>% 
  select(PROD_NAME, PROD_CATEGORY) %>%
  distinct() %>%
  deframe()  # Convert to a named vector where PROD_NAME is the name, and PROD_CATEGORY is the value

# Step 4: Train the Random Forest model on all features (with a refined selection of features)
rf_model_all_features <- randomForest(PROD_NAME ~ PH_OCC + PH_EDUCATION + INS_QS + PPC + Prosperity,
                                      data = train_data,
                                      ntree = 100,
                                      importance = TRUE)

# Step 5: Make predictions on the test data
rf_all_features_predictions <- predict(rf_model_all_features, newdata = test_data)

# Step 6: Compare if the predicted product category matches the actual product category
correct_category_matches <- mapply(check_category_match, 
                                   predicted_product = rf_all_features_predictions, 
                                   actual_product = test_data$PROD_NAME, 
                                   MoreArgs = list(category_mapping = category_mapping))

# Step 7: Calculate the accuracy of category matching
category_accuracy <- mean(correct_category_matches) * 100
print(paste("Category Prediction Accuracy:", round(category_accuracy, 2), "%"))

# Step 8: Create the Shiny app for interaction
ui <- fluidPage(
  
  # Application title
  titlePanel("Insurance Policy Category Prediction System"),
  
  # Sidebar layout
  sidebarLayout(
    sidebarPanel(
      selectInput("PH_OCC", "Occupation:", choices = levels(train_data$PH_OCC)),
      selectInput("PH_EDUCATION", "Education Level:", choices = levels(train_data$PH_EDUCATION)),
      numericInput("INS_QS", "Insurance QS Score:", value = 25, min = 0, max = 50, step = 1),
      numericInput("PPC", "PPC (Policyholder Premium Contribution):", value = 100000, min = 0, step = 10000),
      numericInput("Prosperity", "Prosperity Score:", value = 50, min = 0, max = 100, step = 1),
      actionButton("recommend", "Get Category Recommendation")
    ),
    
    mainPanel(
      h3("Predicted Product Category:"),
      verbatimTextOutput("predicted_category"),
      
      h3("Feature Importance:"),
      plotOutput("feature_importance")
    )
  )
)

# Server logic for Shiny app
server <- function(input, output) {
  
  observeEvent(input$recommend, {
    customer_profile <- data.frame(PH_OCC = factor(input$PH_OCC, levels = levels(train_data$PH_OCC)),
                                   PH_EDUCATION = factor(input$PH_EDUCATION, levels = levels(train_data$PH_EDUCATION)),
                                   INS_QS = input$INS_QS,
                                   PPC = input$PPC,
                                   Prosperity = input$Prosperity)
    
    # Predict product
    predicted_product <- predict(rf_model_all_features, newdata = customer_profile)
    
    # Get predicted product's category
    predicted_category <- category_mapping[predicted_product]
    
    output$predicted_category <- renderText({
      paste("The predicted insurance product category is:", predicted_category)
    })
    
    output$feature_importance <- renderPlot({
      varImpPlot(rf_model_all_features)
    })
  })
}

# Run the Shiny app
shinyApp(ui = ui, server = server)

# Step 9: Final Confusion Matrix for accuracy comparison
rf_conf_matrix <- confusionMatrix(rf_all_features_predictions, test_data$PROD_NAME)
rf_final_accuracy <- rf_conf_matrix$overall["Accuracy"]
print(paste("Final Model Accuracy on Product Prediction:", round(rf_final_accuracy * 100, 2), "%"))

# Step 10: Print the accuracy of category-based matching
print(paste("Final Model Accuracy on Product Category:", round(category_accuracy, 2), "%"))
