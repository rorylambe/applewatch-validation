# Bland-Altman plot for VO2max cosmed (indirect calorimetry) vs Apple Watch

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# Load data from Excel file
file_path = 'cleaned_VO2_data.xlsx'  
sheet_name = 'Sheet1'  
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Drop rows with missing values in the measurements
df = df.dropna(subset=['cosmed_vo2', 'aw_vo2'])

# Drop participants who did not achieve VO2 max
df.drop(df[df['notes'] == 'VO2 peak'].index, inplace=True)

# Specify the names of the columns in the imported Excel file
m1 = df['cosmed_vo2']
m2 = df['aw_vo2']

# Calculate Bland-Altman statistics
mean_diff_values = (m1 + m2) / 2
diff = m1 - m2
mean_diff = diff.mean()
std_diff = diff.std()

# Calculate standard error
n = len(diff)
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
print(f"The number of participants was: {len(m1)}")
print(f"Mean Cosmed: {mean_cosmed} ml/min/kg.\nMean Apple Watch: {mean_aw} ml/min/kg.")
print(f"Cosmed standard deviation: {cosmed_std}.\nAW standard deviation: {aw_std}")
print(f'Standard deviation of difference: {std_diff}')
print(f"Cosmed standard error: {se_cosmed}.\nAW standard error: {se_aw}")
print(f'Standard error of difference: {se_diff}')
print(f'Mean difference: {mean_diff}')
print(f'95% CI of mean difference: [{ci_lower}, {ci_upper}]')
print(f'Upper limit of agreement: {loa_upper}')
print(f'Lower limit of agreement: {loa_lower}')

# Create Bland-Altman plot manually
plt.rcParams['font.family'] = 'serif'
plt.figure(figsize=(8, 5))
plt.scatter(mean_diff_values, diff, alpha=0.5)
plt.axhline(mean_diff, color='gray', linestyle='--', label=f'Mean Difference: {mean_diff:.2f}')
plt.axhline(loa_upper, color='red', linestyle='--', label=f'Upper LoA: {loa_upper:.2f}')
plt.axhline(loa_lower, color='red', linestyle='--', label=f'Lower LoA: {loa_lower:.2f}')
plt.title('Bland-Altman Plot Comparing VO\u2082 max Values')
plt.xlabel('Mean of COSMED and Apple Watch VO\u2082 max')
plt.ylabel('Difference between COSMED and Apple Watch VO\u2082 max')
plt.legend()
plt.show()
