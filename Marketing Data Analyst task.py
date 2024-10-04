import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy.stats import ttest_ind
# Function to read and prepare the data
def load_and_prepare_data(file_path):
    df = pd.read_excel(file_path)
    df['Product profit'] = df['Product profit'].fillna(0)
    df['Reward'] = df['Reward'].fillna(0)
    return df
# Function to segment data into groups
def segment_data(df):
    group_a = df[df['Group'] == 'A']
    group_b = df[df['Group'] == 'B']
    group_c = df[df['Group'] == 'C']
    return group_a, group_b, group_c
# 1. Function to calculate Target Achievement Rate (TAR)
def calculate_tar(group):
    if 'Target' in group.columns:
        achieved = (group['Number of Finished Products'] >= group['Target']).sum()
        return achieved / len(group) * 100 if len(group) > 0 else 0
    return np.nan  # Group C does not have targets
# 2. Function to calculate Average Profit per Agent (APA)
def calculate_apa(group):
    return group['Product profit'].sum() / len(group) if len(group) > 0 else 0
# 3. Function to calculate Reward Utilization Efficiency (RUE)
def calculate_rue(group):
    total_rewards = group['Reward'].sum()
    total_profit = group['Product profit'].sum()
    return (total_rewards / total_profit) * 100 if total_profit > 0 else np.nan
# 4. Function to calculate Return on Investment (ROI)
def calculate_roi(group):
    total_rewards = group['Reward'].sum()
    total_profit = group['Product profit'].sum()
    return ((total_profit - total_rewards) / total_rewards) * 100 if total_rewards > 0 else np.nan
# 5. Function to calculate Sales Performance (SP)
def calculate_sp(group):
    return group['Number of Finished Products'].sum() / len(group) if len(group) > 0 else 0
# 6. Function to calculate Cost per Unit Sold (CUS)
def calculate_cus(group):
    total_rewards = group['Reward'].sum()
    total_products = group['Number of Finished Products'].sum()
    return (total_rewards / total_products) if total_products > 0 else np.nan
# Function to compare with control group
def compare_with_control(group, control, column):
    t_stat, p_value = ttest_ind(group[column], control[column], nan_policy='omit')
    return p_value
# Function to check if the difference is significant
def check_success(p_value):
    if p_value < 0.05:
        return 'Statistically significant difference compared to Control'
    else:
        return 'No significant difference compared to Control'
# Function to plot distributions of key metrics
def plot_distributions(df):
    plt.figure(figsize=(12, 6))
    sns.histplot(df[df['Group'] == 'A']['Product profit'], color='blue', label='Group A', kde=True)
    sns.histplot(df[df['Group'] == 'B']['Product profit'], color='green', label='Group B', kde=True)
    sns.histplot(df[df['Group'] == 'C']['Product profit'], color='red', label='Group C', kde=True)
    plt.title('Product Profit Distribution Across Groups')
    plt.legend()
    plt.show()
# Function to visualize performance metrics (optional: you can customize it further)
def plot_bar_chart(groups, kpi, title):
    kpi_values = [kpi(g) for g in groups]
    group_names = ['Group A', 'Group B', 'Group C']
    
    plt.figure(figsize=(8, 5))
    plt.bar(group_names, kpi_values, color=['blue', 'green', 'red'])
    plt.title(title)
    plt.ylabel('Value')
    plt.show()
# Main function to execute the analysis and visualization
def analyze_campaign(file_path):
    df = load_and_prepare_data(file_path)
    group_a, group_b, group_c = segment_data(df)

    # Calculate KPIs for each group
    tar_a = calculate_tar(group_a)
    tar_b = calculate_tar(group_b)

    apa_a = calculate_apa(group_a)
    apa_b = calculate_apa(group_b)
    apa_c = calculate_apa(group_c)

    rue_a = calculate_rue(group_a)
    rue_b = calculate_rue(group_b)

    roi_a = calculate_roi(group_a)
    roi_b = calculate_roi(group_b)

    sp_a = calculate_sp(group_a)
    sp_b = calculate_sp(group_b)
    sp_c = calculate_sp(group_c)

    cus_a = calculate_cus(group_a)
    cus_b = calculate_cus(group_b)

    # Statistical test
    p_value_profit_a = compare_with_control(group_a, group_c, 'Product profit')
    p_value_profit_b = compare_with_control(group_b, group_c, 'Product profit')

    decision_a = check_success(p_value_profit_a)
    decision_b = check_success(p_value_profit_b)

    # Output KPIs and decisions
    print(f"Group A:")
    print(f"Target Achievement Rate (TAR): {tar_a:.2f}%")
    print(f"Average Profit per Agent (APA): ${apa_a:.2f}")
    print(f"Reward Utilization Efficiency (RUE): {rue_a:.2f}%")
    print(f"Return on Investment (ROI): {roi_a:.2f}%")
    print(f"Sales Performance (SP): {sp_a:.2f} products per agent")
    print(f"Cost per Unit Sold (CUS): ${cus_a:.2f}")
    print(f"Decision: {decision_a}")

    print("\nGroup B:")
    print(f"Target Achievement Rate (TAR): {tar_b:.2f}%")
    print(f"Average Profit per Agent (APA): ${apa_b:.2f}")
    print(f"Reward Utilization Efficiency (RUE): {rue_b:.2f}%")
    print(f"Return on Investment (ROI): {roi_b:.2f}%")
    print(f"Sales Performance (SP): {sp_b:.2f} products per agent")
    print(f"Cost per Unit Sold (CUS): ${cus_b:.2f}")
    print(f"Decision: {decision_b}")

    print("\nGroup C (Control Group):")
    print(f"Average Profit per Agent (APA): ${apa_c:.2f}")
    print(f"Sales Performance (SP): {sp_c:.2f} products per agent")

    # Plot distributions of product profit
    plot_distributions(df)

    # Plot bar charts for comparison
    plot_bar_chart([group_a, group_b, group_c], calculate_apa, 'Average Profit per Agent (APA)')
    plot_bar_chart([group_a, group_b, group_c], calculate_sp, 'Sales Performance (SP)')
# File path to the Excel file
file_path = r"C:\Users\amira\Downloads\case_study.xlsx"  # Replace with the actual file path
analyze_campaign(file_path)
