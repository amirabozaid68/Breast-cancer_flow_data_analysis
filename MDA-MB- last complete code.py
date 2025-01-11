import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import statsmodels.api as sm
from statsmodels.formula.api import ols
from statsmodels.stats.multicomp import pairwise_tukeyhsd

# Define the data for MDA-MB-231 Cell Line
data = {
    "Dose": ["Control", "50 µg/mL", "50 µg/mL", "100 µg/mL", "100 µg/mL", "200 µg/mL", "200 µg/mL"],
    "Sample_ID": ["Control", "50-1", "50-2", "100-1", "100-2", "200-1", "200-2"],
    "Live_Cells": [90.75, 33.32, 28.39, 28.52, 28.57, 25.63, 24.01],
    "Early_Apoptotic": [8.02, 14.02, 21.77, 21.96, 25.33, 26.93, 25.54],
    "Late_Apoptotic_Necrotic": [0.12, 43.25, 43.88, 45.50, 42.11, 45.20, 46.43],
    "Necrotic": [1.10, 9.41, 5.95, 4.02, 3.99, 2.24, 4.02]
}

# Convert data to DataFrame
df = pd.DataFrame(data)

# Setting up output paths
output_dir = r'C:\Users\amira\OneDrive - Alexandria University\Desktop\Flow Cytometry\25-12-2024-MDA-MB-231\MDA-MB-231-Outputs'
anova_path = f"{output_dir}/ANOVA_Results.xlsx"
tukey_path = f"{output_dir}/MDA_MB231_Tukeys_HSD_Results.xlsx"
combined_plot_path = f"{output_dir}/Combined_Visualization.png"
phase_comparison_plot_path = f"{output_dir}/Phase_Comparison_Visualization.png"
tukey_plot_path = f"{output_dir}/MDA_MB231_Tukey_HSD_Plot.png"

# Perform ANOVA and save results
writer = pd.ExcelWriter(anova_path, engine='xlsxwriter')
for column in df.columns[2:]:
    model = ols(f'{column} ~ C(Dose)', data=df).fit()
    anova_table = sm.stats.anova_lm(model, typ=2)
    anova_table.to_excel(writer, sheet_name=column)
writer.close()
print(f"ANOVA results saved to {anova_path}")

# Perform Tukey's HSD Test on Live_Cells
tukey = pairwise_tukeyhsd(endog=df['Live_Cells'], groups=df['Dose'], alpha=0.05)
tukey_results_df = pd.DataFrame(data=tukey._results_table.data[1:], columns=tukey._results_table.data[0])
tukey_results_df.to_excel(tukey_path, index=False)
print(f"Tukey's test results saved to {tukey_path}")

# Plot Tukey's HSD results
plt.figure(figsize=(10, 8))
tukey.plot_simultaneous()
plt.title('MDA-MB-231 Tukey HSD Test Results')
plt.savefig(tukey_plot_path)
plt.close()

# Combined bar plot for all cell types
sns.set(style="whitegrid")
fig, axes = plt.subplots(nrows=2, ncols=2, figsize=(12, 10))
fig.suptitle('Treatment Response Across Cell Types')

cell_types = ['Live_Cells', 'Early_Apoptotic', 'Late_Apoptotic_Necrotic', 'Necrotic']
for i, ax in enumerate(axes.flatten()):
    sns.barplot(x='Dose', y=cell_types[i], data=df, ax=ax)
    ax.set_title(cell_types[i].replace('_', ' '))
    ax.set_ylabel('% of Cells')
    ax.set_xlabel('Dose µg/mL')

plt.tight_layout(rect=[0, 0.03, 1, 0.95])
plt.savefig(combined_plot_path)
plt.close()
print(f"Combined visualization saved to {combined_plot_path}")

# Detailed bar plot for each cell type with Sample_ID
plt.figure(figsize=(12, 8))
for i, column in enumerate(['Live_Cells', 'Early_Apoptotic', 'Late_Apoptotic_Necrotic', 'Necrotic'], start=1):
    plt.subplot(2, 2, i)
    sns.barplot(x='Dose', y=column, hue='Sample_ID', data=df, palette='viridis')
    plt.title(f'{column} Response by Dose')
    plt.ylabel('Percentage')
    plt.xlabel('Dose')
    plt.xticks(rotation=45)
    plt.legend(title='Sample ID')

plt.tight_layout()
plt.savefig(phase_comparison_plot_path)
plt.close()
print(f"Phase comparison visualization saved to {phase_comparison_plot_path}")
