"""
Sample Submission STU-035: Incomplete Customer Churn Prediction
Role: AI/ML Engineer
Task: TASK-AIML-101
"""

# Deficient submission: Hardcoded credentials, no exception handling, missing FastAPI endpoints, missing Dockerfile
import pandas as pd
from sklearn.ensemble import RandomForestClassifier

# NOTE: Missing data preprocessing pipeline
df = pd.read_csv("customers.csv")
X = df[['tenure', 'MonthlyCharges']]
y = df['Churn']

# Hardcoded model without hyperparameter tuning or cross-validation
model = RandomForestClassifier()
model.fit(X, y)

# Missing FastAPI REST service!
# Missing Dockerfile!
# Missing test suite!
print("Model trained.")
