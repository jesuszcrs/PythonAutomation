# Write a Python function to preprocess a given dataset for analysis.
# Include steps such as handling missing values, removing outliers,
# and standardizing or normalizing numerical features.
import pandas as pd
import numpy as np
from sklearn.preprocessing import StandardScaler

# Create a sample DataFrame
data = {
    'price': [10.2, 15.6, np.nan, 22.3, 18.9],
    'quantity': [25, np.nan, 35, 40, 30],
    'sales': [np.nan, 200, 300, 400, 500]
}
df = pd.DataFrame(data)

def preprocess_data(df):
    # Handle missing values
    df.fillna(df.mean(), inplace=True)
    
    # Remove outliers (example using Z-score)
    z_scores = (df - df.mean()) / df.std()
    df = df[(z_scores < 3).all(axis=1)]
    
    # Define numerical_cols
    numerical_cols = ['price', 'quantity', 'sales']
    
    # Standardize numerical features to have a mean of 0 and a standard deviation of 1
    scaler = StandardScaler()
    df[numerical_cols] = scaler.fit_transform(df[numerical_cols])
    
    return df

# Preprocess the sample DataFrame
preprocessed_df = preprocess_data(df)
print(preprocessed_df)

# Implement a simple algorithm (e.g., linear regression, decision tree)
# from scratch using Python. Test the algorithm on a sample dataset
# and evaluate its performance.
import numpy as np 

class SimpleLinearRegression:
    def __init__(self):
        self.coefficient = None
        self.intercept = None

    # Calculates the coefficients of the regression line based on the input features X and target variable y
    def fit(self, X, y):
        X_mean = np.mean(X)
        y_mean = np.mean(y)
        numerator = np.sum((X-X_mean) * (y - y_mean))
        denominator = np.sum((X - X_mean) ** 2)
        self.coefficient = numerator / denominator
        self.intercept = y_mean - self.coefficient * X_mean

    # Predicts the target variable y for given input features X using the coefficients calculated during fitting
    def predict(self, X):
        return self.coefficient * X + self.intercept

# Test the algorithm
X = np.array([1, 2, 3, 4, 5])
y = np.array([2, 4, 5, 4, 5])
model = SimpleLinearRegression()
model.fit(X, y)
predictions = model.predict(X)
print(predictions)