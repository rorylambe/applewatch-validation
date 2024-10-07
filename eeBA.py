# Bland-Altman plot for VO2max cosmed (indirect calorimetry) vs Apple Watch

import pandas as pd
import statsmodels.api as sm
import matplotlib.pyplot as plt
import numpy as np

# Load data from Excel file
file_path = 'cleaned_VO2_data.xlsx'  # Specify file path
sheet_name = 'Sheet1'  # Specify sheet name
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Drop rows with missing values in the measurements
df = df.dropna(subset=['cosmed_kcal', 'aw_kcal'])

# Specify the names of the columns in the imported Excel file
m1 = df['cosmed_kcal']
m2 = df['aw_kcal']

# Calculate Bland-Altman statistics
mean_diff_values = (m1 + m2) / 2
diff = m1 - m2
mean_diff = diff.mean()
std_diff = diff.std()

# Calculate standard error
n = len(m1)
se_diff = std_diff / np.sqrt(n)

# Calculate limits of agreement
loa_upper = mean_diff + 1.96 * std_diff
loa_lower = mean_diff - 1.96 * std_diff

# Calculate 95% CI for the mean difference
ci_upper = mean_diff + 1.96 * se_diff
ci_lower = mean_diff - 1.96 * se_diff

# Calculate mean and standard deviation for Cosmed and Apple Watch
mean_cosmed = m1.mean()
cosmed_std = m1.std()
mean_aw = m2.mean()
aw_std = m2.std()

# Calculate standard error for Cosmed and Apple Watch
se_cosmed = cosmed_std / np.sqrt(n)
se_aw = aw_std / np.sqrt(n)

# Print results
print(f"Mean Cosmed: {mean_cosmed} kcals.\nMean Apple Watch: {mean_aw} kcals.")
print(f"Cosmed standard deviation: {cosmed_std}.\nAW standard deviation: {aw_std}")
print(f'Standard deviation of difference: {std_diff}')
print(f"Cosmed standard error: {se_cosmed}.\nAW standard error: {se_aw}")
print(f'Standard error of difference: {se_diff}')
print(f'Mean difference: {mean_diff}')
print(f'95% CI of mean difference: [{ci_lower}, {ci_upper}]')
print(f'Upper limit of agreement: {loa_upper}')
print(f'Lower limit of agreement: {loa_lower}')

# Create Bland-Altman plot
f, ax = plt.subplots(1, figsize=(8, 5))
sm.graphics.mean_diff_plot(m1, m2, ax=ax)
ax.axhline(mean_diff, color='gray', linestyle='--')
ax.axhline(loa_upper, color='red', linestyle='--')
ax.axhline(loa_lower, color='red', linestyle='--')
plt.title('Bland-Altman Plot')
plt.xlabel('Mean of Cosmed and Apple Watch VO2')
plt.ylabel('Difference between Cosmed and Apple Watch energy expenditure')
plt.show()