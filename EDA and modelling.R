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
install.packages("e1071")




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
library(e1071)


file_path <- "C:/Users/angad/Downloads/Data.xlsb"
data <- read_xlsb(file_path, sheet = 1, range = "A1:Z50000")

# Initial Data Exploration
head(data)
str(data)
summary(data)
sapply(data, function(x) sum(is.na(x))) # Check for missing values
skim(data) # Overview of the data

# EDA: Plot categorical variables
plot_categorical <- function(column_name) {
  ggplot(data, aes_string(x = column_name)) + 
    geom_bar() + 
    theme_minimal() + 
    ggtitle(paste("Distribution of", column_name)) +
    theme(axis.text.x = element_text(angle = 45, hjust = 1))
}

plot_categorical("PH_OCC")
plot_categorical("PH_EDUCATION")
plot_categorical("PROD_NAME")
plot_categorical("PROD_CATEGORY")
plot_categorical("PRODUCT_VARIANT")
plot_categorical("PH_PROS_BAND")
plot_categorical("INCOME_SEGMENT")

# EDA: Plot continuous variables
plot_continuous <- function(column_name, binwidth = 10) {
  ggplot(data, aes_string(x = column_name)) + 
    geom_histogram(binwidth = binwidth, fill = "blue", color = "black") + 
    theme_minimal() + 
    ggtitle(paste("Distribution of", column_name))
}

plot_continuous("PPC", binwidth = 5000)  # Adjust binwidth if necessary
plot_continuous("Prosperity", binwidth = 5)

# Check for correlations among continuous variables
cor_matrix <- cor(data %>% select_if(is.numeric), use = "complete.obs")
ggcorrplot::ggcorrplot(cor_matrix, method = "circle", type = "lower", 
                       lab = TRUE, title = "Correlation Matrix")

# Pair plot for continuous variables
ggpairs(data %>% select_if(is.numeric))

# Boxplot example for continuous vs categorical variable
ggplot(data, aes(x = PH_OCC, y = Prosperity)) +
  geom_boxplot(fill = "blue") +
  theme_minimal() +
  ggtitle("Prosperity by Occupation") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

ggplot(data, aes(x = PH_PROS_BAND, y = PPC)) +
  geom_boxplot(fill = "blue") +
  theme_minimal() +
  ggtitle("PPC by PH_PROS_BAND") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

# Scatter plot example for relationship between two continuous variables
ggplot(data, aes(x = Prosperity, y = PPC)) +
  geom_point(alpha = 0.5) +
  geom_smooth(method = "lm", color = "red") +
  theme_minimal() +
  ggtitle("Relationship between Prosperity and PPC")


# Data Preparation for Modeling

# Filtering out missing values and ensuring correct data types
# Data cleaning and transformation
data_clean <- data %>%
  mutate(PPC = ifelse(is.na(PPC), mean(PPC, na.rm = TRUE), PPC),
         Prosperity = ifelse(is.na(Prosperity), mean(Prosperity, na.rm = TRUE), Prosperity),
         PROD_NAME = as.factor(PROD_NAME),
         PH_OCC = as.factor(PH_OCC),
         PH_EDUCATION = as.factor(PH_EDUCATION),
         INCOME_SEGMENT = as.factor(INCOME_SEGMENT))

# Splitting the dataset into training and testing sets
set.seed(123)  # For reproducibility
split <- sample.split(data_clean$PROD_NAME, SplitRatio = 0.7)
train_data <- subset(data_clean, split == TRUE)
test_data <- subset(data_clean, split == FALSE)

train_data <- na.omit(train_data)
test_data <- na.omit(test_data)

# Remove outliers
train_data <- train_data %>%
  filter(PPC < quantile(PPC, 0.99))  # Remove top 1% outliers

# Train Random Forest model with all features
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

# --- MODEL COMPARISON SECTION ---

# 1. Random Forest (already trained as rf_model_refined)
rf_accuracy <- confusionMatrix(rf_refined_predictions, test_data$PROD_NAME)$overall["Accuracy"]

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
  Model = c("Random Forest", "KNN", "Decision Tree"),
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
    
    predicted_product <- predict(rf_model_refined, newdata = customer_profile)
    
    output$recommended_product <- renderText({
      paste("The recommended insurance policy is:", predicted_product)
    })
    
    output$feature_importance <- renderPlot({
      varImpPlot(rf_model_refined)
    })
  })
}

# Run the Shiny app
shinyApp(ui = ui, server = server)

#Accuracy of the final model
# Step 1: Make predictions on the test data using the final Random Forest model
rf_final_predictions <- predict(rf_model_refined, newdata = test_data)

# Step 2: Compare predictions with actual data using confusion matrix
rf_conf_matrix <- confusionMatrix(rf_final_predictions, test_data$PROD_NAME)

# Step 3: Print the confusion matrix and accuracy metrics
print("Confusion Matrix for the Final Random Forest Model:")
print(rf_conf_matrix)

# Step 4: Extract and print the accuracy
rf_final_accuracy <- rf_conf_matrix$overall["Accuracy"]
print(paste("Final Model Accuracy:", round(rf_final_accuracy * 100, 2), "%"))



