import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import statsmodels.api as sm
from statsmodels.formula.api import ols
from statsmodels.stats.multicomp import pairwise_tukeyhsd

# Define the data for MCF-7 Cell Line
data = {
    "Dose": [0, 50, 50, 100, 100, 200, 200],
    "Live_Cells": [63.18, 24.47, 25.38, 21.69, 19.81, 14.34, 12.34],
    "Early_Apoptotic_Cells": [0.66, 34.69, 26.15, 23.31, 17.46, 15.01, 17.65],
    "Late_Apoptotic_Necrotic_Cells": [0.70, 17.87, 19.15, 24.74, 34.00, 31.20, 33.47],
    "Necrotic_Cells": [3.50, 6.04, 8.31, 11.40, 14.03, 20.07, 18.41]
}

# Convert data to DataFrame
df = pd.DataFrame(data)

# Setting up output paths
output_dir = r'C:\Users\amira\OneDrive - Alexandria University\Desktop\Flow Cytometry\01-01-2025 MCF-7\MCF-7 output analysis'
anova_path = f"{output_dir}/ANOVA_Results_MCF7.xlsx"
tukey_path = f"{output_dir}/MCF7_Tukeys_HSD_Results.xlsx"
combined_plot_path = f"{output_dir}/Combined_Visualization_MCF7.png"
phase_comparison_plot_path = f"{output_dir}/Phase_Comparison_Visualization_MCF7.png"
tukey_plot_path = f"{output_dir}/MCF7_Tukey_HSD_Plot.png"

# Perform ANOVA and save results
writer = pd.ExcelWriter(anova_path, engine='xlsxwriter')
for column in df.columns[1:]:
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
plt.title('MCF-7 Tukey HSD Test Results')
plt.savefig(tukey_plot_path)
plt.close()

# Combined bar plot for all cell types
sns.set(style="whitegrid")
fig, axes = plt.subplots(nrows=2, ncols=2, figsize=(12, 10))
fig.suptitle('Treatment Response Across Cell Types for MCF-7')

cell_types = ['Live_Cells', 'Early_Apoptotic_Cells', 'Late_Apoptotic_Necrotic_Cells', 'Necrotic_Cells']
for i, ax in enumerate(axes.flatten()):
    sns.barplot(x='Dose', y=cell_types[i], data=df, ax=ax)
    ax.set_title(cell_types[i].replace('_', ' '))
    ax.set_ylabel('% of Cells')
    ax.set_xlabel('Dose (µg/mL)')

plt.tight_layout(rect=[0, 0.03, 1, 0.95])
plt.savefig(combined_plot_path)
plt.close()
print(f"Combined visualization saved to {combined_plot_path}")

# Detailed bar plot for each cell type with Dose
plt.figure(figsize=(12, 8))
for i, column in enumerate(['Live_Cells', 'Early_Apoptotic_Cells', 'Late_Apoptotic_Necrotic_Cells', 'Necrotic_Cells'], start=1):
    plt.subplot(2, 2, i)
    sns.barplot(x='Dose', y=column, hue='Dose', data=df, palette='viridis', dodge=False)  # Fixed the warning
    plt.title(f'{column} Response by Dose')
    plt.ylabel('Percentage')
    plt.xlabel('Dose (µg/mL)')
    plt.xticks(rotation=45)
    plt.legend(title='Dose', loc='upper right')

plt.tight_layout()
plt.savefig(phase_comparison_plot_path)
plt.close()
print(f"Phase comparison visualization saved to {phase_comparison_plot_path}")
