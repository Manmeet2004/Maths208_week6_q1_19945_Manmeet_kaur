import datetime
import openpyxl
from pathlib import Path
import numpy as np
import matplotlib.pyplot as plt

# Define a function to calculate the mean
def calculate_mean(data):
    return sum(data) / len(data)

# Define a function to calculate the variance
def calculate_variance(data, mean):
    return sum((x - mean) ** 2 for x in data) / (len(data) - 1)

# Define a function to calculate the standard deviation
def calculate_std_dev(variance):
    return variance ** 0.5

# Define a function to calculate the z-scores
def calculate_z_scores(data, mean, std_dev):
    return [(x - mean) / std_dev for x in data]

# Define a function to calculate the quartiles (Q1, median, Q3)
def calculate_quartiles(data):
    data.sort()
    n = len(data)
    q1 = (data[n // 4] + data[(n // 4) - 1]) / 2 if n % 4 == 0 else data[n // 4]
    median = (data[n // 2] + data[(n // 2) - 1]) / 2 if n % 2 == 0 else data[n // 2]
    q3 = (data[(3 * n // 4)] + data[(3 * n // 4) - 1]) / 2 if n % 4 == 0 else data[(3 * n // 4)]
    return q1, median, q3

# Construct the path to the Excel file in your project directory
file_path = Path(__file__).parent / "mydata.xlsx"

# Load the Excel workbook
workbook = openpyxl.load_workbook(file_path)

# Select a specific sheet
sheet = workbook.active

# Read data from the sheet
glucose_data = []
blood_pressure_data = []

for cell in sheet['A']:
    if cell.value is not None:
        try:
            glucose_data.append(int(cell.value))
        except (ValueError, TypeError):
            continue

for cell in sheet['B']:
    if cell.value is not None:
        try:
            blood_pressure_data.append(int(cell.value))
        except (ValueError, TypeError):
            continue

# Calculate statistics for "Glucose"
glucose_mean = calculate_mean(glucose_data)
# Calculate statistics for "BloodPressure"
blood_pressure_mean = calculate_mean(blood_pressure_data)


glucose_variance = calculate_variance(glucose_data, glucose_mean)
glucose_std_dev = calculate_std_dev(glucose_variance)
glucose_z_scores = calculate_z_scores(glucose_data, glucose_mean, glucose_std_dev)
glucose_q1, glucose_median, glucose_q3 = calculate_quartiles(glucose_data)

# Calculate statistics for "BloodPressure"
blood_pressure_mean = calculate_mean(blood_pressure_data)
blood_pressure_variance = calculate_variance(blood_pressure_data, blood_pressure_mean)
blood_pressure_std_dev = calculate_std_dev(blood_pressure_variance)
blood_pressure_z_scores = calculate_z_scores(blood_pressure_data, blood_pressure_mean, blood_pressure_std_dev)
blood_pressure_q1, blood_pressure_median, blood_pressure_q3 = calculate_quartiles(blood_pressure_data)

# Print the statistics
print("Statistics for 'Glucose':")
print(f"Mean: {glucose_mean}")
print(f"Variance: {glucose_variance}")
print(f"Standard Deviation: {glucose_std_dev}")
print(f"Z-Scores: {glucose_z_scores}")
print(f"Q1: {glucose_q1}")
print(f"Median: {glucose_median}")
print(f"Q3: {glucose_q3}")

print("\nStatistics for 'BloodPressure':")
print(f"Mean: {blood_pressure_mean}")
print(f"Variance: {blood_pressure_variance}")
print(f"Standard Deviation: {blood_pressure_std_dev}")
print(f"Z-Scores: {blood_pressure_z_scores}")
print(f"Q1: {blood_pressure_q1}")
print(f"Median: {blood_pressure_median}")
print(f"Q3: {blood_pressure_q3}")

# Create a boxplot
plt.figure(figsize=(8, 6))
plt.boxplot([glucose_data, blood_pressure_data], labels=['Glucose', 'BloodPressure'])
plt.title('Boxplots for Glucose and BloodPressure')
plt.ylabel('Values')
plt.show()

# Close the workbook
workbook.close()

current_datetime = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
print(f"Current date and time: {current_datetime}")