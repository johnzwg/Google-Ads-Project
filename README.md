# This Python script provides a comprehensive analysis of simulated Google Ads performance data.
# It covers various aspects including:
# - Data Simulation: Generates realistic-looking data at Campaign, Ad Group, Keyword, and Demographic levels.
# - Metric Calculation: Computes key performance indicators (KPIs) such as CTR, CPC, Conversion Rate, CPA, and ROAS.
# - Performance Aggregation: Aggregates data daily, by campaign, ad group, keyword, and demographic segments (Age, Gender, Location).
# - Time-Series Analysis: Calculates period-over-period changes (e.g., Week-over-Week).
# - Anomaly Detection: Implements a Z-score based method to identify significant performance deviations.
# - Budget Pacing Simulation: Provides a basic model to track simulated budget consumption.
# - Optimization Insights: Generates simulated recommendations based on performance rules and demographic analysis.
# - Data Visualization: Creates and saves various plots (line charts for trends, bar charts for comparisons)
#   to visually represent performance across different dimensions.
# - XLSX Export: Exports the raw simulated data and all aggregated reports into a multi-sheet Excel file,
#   suitable for sharing or further analysis.
# This project demonstrates strong Python skills in data manipulation (pandas), statistical analysis (numpy, scipy),
# data visualization (matplotlib, seaborn), and structured reporting, making it ideal for showcasing capabilities
# in digital marketing analytics.

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime, timedelta
import os
import random
from scipy.stats import zscore # For more robust anomaly detection
import matplotlib.pyplot as plt # For data visualization
import seaborn as sns # For enhanced plot aesthetics

print("--- Starting Comprehensive Google Ads Performance Analysis Project (with Graphs) ---")
print("This script will simulate detailed Google Ads data, perform extensive analysis,")
print("generate an XLSX file with raw and aggregated data, and create various graphs.")

# --- Configuration Parameters ---
# These parameters control the scale and characteristics of the simulated data.
NUM_DAYS = 120  # Simulate data for 120 days (approx. 4 months)
NUM_CAMPAIGNS = 8 # Number of top-level campaigns
AD_GROUPS_PER_CAMPAIGN = 3 # Average number of ad groups per campaign
KEYWORDS_PER_AD_GROUP = 10 # Average number of keywords per ad group

# Anomaly simulation parameters
ANOMALY_PROBABILITY = 0.05 # Probability of a significant anomaly on any given day/campaign
ANOMALY_MAGNITUDE_FACTOR = 0.2 # Clicks/Conversions drop to 20% of expected during anomaly

# Anomaly detection parameters
ANOMALY_DETECTION_WINDOW_SIZE = 14 # Look at the last 14 days for anomaly detection
Z_SCORE_THRESHOLD = 2.5 # Z-score threshold for flagging anomalies (e.g., 2.5 standard deviations)

# --- 1. Data Simulation Module ---
# This module contains functions to generate realistic-looking simulated Google Ads data
# at different hierarchical levels (Campaign -> Ad Group -> Keyword -> Demographic).

def generate_campaign_names(num_campaigns):
    """Generates a list of unique campaign names."""
    campaign_types = ["Search", "Shopping", "Display", "Video", "App"]
    product_categories = ["Electronics", "Apparel", "HomeGoods", "Services", "Travel"]
    names = []
    for i in range(num_campaigns):
        camp_type = random.choice(campaign_types)
        category = random.choice(product_categories)
        names.append(f"{camp_type} - {category} Campaign {i+1}")
    return list(set(names)) # Ensure uniqueness

def generate_ad_group_names(campaign_name, num_ad_groups):
    """Generates ad group names for a given campaign."""
    ad_group_types = ["Brand", "Generic", "Competitor", "Remarketing"]
    names = []
    for i in range(num_ad_groups):
        ad_type = random.choice(ad_group_types)
        names.append(f"{campaign_name} - {ad_type} AG {i+1}")
    return names

def generate_keywords(ad_group_name, num_keywords):
    """Generates simulated keywords for a given ad group."""
    base_keywords = ["buy", "best", "cheap", "online", "deal", "review", "price"]
    # Extract a relevant product term from the ad group name for more realistic keywords
    product_terms_parts = ad_group_name.split(' - ')
    if len(product_terms_parts) > 1:
        product_term_candidate = product_terms_parts[1].replace(' Campaign', '').replace(' AG', '').split(' ')[0]
        if product_term_candidate:
            product_terms = product_term_candidate
        else:
            product_terms = "item" # Fallback
    else:
        product_terms = "item" # Fallback

    keywords = []
    for i in range(num_keywords):
        term = random.choice(base_keywords)
        keywords.append(f"{term} {product_terms} {i+1}")
    return keywords

def simulate_daily_performance(base_impressions, base_clicks, base_cost, base_conversions):
    """Simulates daily performance metrics with some natural variance."""
    impressions = max(10, int(base_impressions * np.random.uniform(0.8, 1.2)))
    clicks = max(1, int(base_clicks * np.random.uniform(0.8, 1.2)))
    cost = round(base_cost * np.random.uniform(0.8, 1.2), 2)
    conversions = max(0, int(base_conversions * np.random.uniform(0.7, 1.3)))
    return impressions, clicks, cost, conversions

def apply_anomaly(clicks, conversions, cost, magnitude_factor):
    """Applies a simulated anomaly (e.g., a drop in performance)."""
    clicks = max(1, int(clicks * magnitude_factor * np.random.uniform(0.1, 0.5))) # More severe drop for visibility
    conversions = max(0, int(conversions * magnitude_factor * np.random.uniform(0.1, 0.5)))
    cost = round(cost * magnitude_factor * np.random.uniform(0.5, 0.8), 2)
    return clicks, conversions, cost

def generate_ads_data_comprehensive(num_days, num_campaigns, ad_groups_per_campaign, keywords_per_ad_group):
    """
    Generates comprehensive simulated Google Ads data across campaigns, ad groups, keywords, and demographics.
    """
    print(f"Generating comprehensive simulated Google Ads data for {num_days} days...")
    all_data = []
    start_date = datetime.now() - timedelta(days=num_days)
    
    campaign_names = generate_campaign_names(num_campaigns)
    
    # Define demographic segments
    age_groups = ["18-24", "25-34", "35-44", "45-54", "55-64", "65+"]
    genders = ["Male", "Female", "Undetermined"]
    locations = ["Urban", "Suburban", "Rural"] # Simplified locations

    for i in range(num_days):
        current_date = start_date + timedelta(days=i)
        
        # Introduce a "global" anomaly day with a certain probability
        is_global_anomaly_day = np.random.rand() < ANOMALY_PROBABILITY / 5 # Less frequent global anomalies

        for campaign_name in campaign_names:
            num_ad_groups_actual = random.randint(max(1, ad_groups_per_campaign - 1), ad_groups_per_campaign + 1)
            ad_group_names = generate_ad_group_names(campaign_name, num_ad_groups_actual)

            for ad_group_name in ad_group_names:
                num_keywords_actual = random.randint(max(1, keywords_per_ad_group - 3), keywords_per_ad_group + 3)
                keywords = generate_keywords(ad_group_name, num_keywords_actual)

                for keyword in keywords:
                    # Simulate performance for each demographic segment
                    for age_group in age_groups:
                        for gender in genders:
                            for location in locations:
                                # Base performance for this keyword and demographic segment
                                # Adjust base performance slightly by demographic to create variations
                                base_impressions = np.random.randint(50, 500)
                                base_clicks = np.random.randint(5, 50)
                                base_cost = round(np.random.uniform(0.5, 5.0), 2)
                                base_conversions = np.random.randint(0, 5)

                                # Introduce some demographic-based performance biases
                                if age_group == "25-34" and gender == "Female" and "Apparel" in campaign_name:
                                    base_conversions = int(base_conversions * 1.5) # Higher conversions
                                elif age_group == "65+" and "Electronics" in campaign_name:
                                    base_conversions = int(base_conversions * 0.7) # Lower conversions

                                impressions, clicks, cost, conversions = simulate_daily_performance(
                                    base_impressions, base_clicks, base_cost, base_conversions
                                )

                                # Apply anomalies based on probability or global anomaly day
                                if is_global_anomaly_day or np.random.rand() < ANOMALY_PROBABILITY:
                                    clicks, conversions, cost = apply_anomaly(clicks, conversions, cost, ANOMALY_MAGNITUDE_FACTOR)

                                all_data.append({
                                    'Date': current_date.strftime('%Y-%m-%d'),
                                    'Campaign Name': campaign_name,
                                    'Ad Group Name': ad_group_name,
                                    'Keyword': keyword,
                                    'Age Group': age_group,
                                    'Gender': gender,
                                    'Location': location,
                                    'Impressions': impressions,
                                    'Clicks': clicks,
                                    'Cost': cost,
                                    'Conversions': conversions,
                                    'Simulated Revenue': round(conversions * np.random.uniform(50, 200), 2) # Simulate revenue per conversion
                                })
    return pd.DataFrame(all_data)

# Generate the comprehensive dataset
ads_df = generate_ads_data_comprehensive(NUM_DAYS, NUM_CAMPAIGNS, AD_GROUPS_PER_CAMPAIGN, KEYWORDS_PER_AD_GROUP)
ads_df['Date'] = pd.to_datetime(ads_df['Date'])

# --- 2. Core Metrics Calculation Module ---
# This module defines functions to calculate standard and advanced Google Ads metrics.

def calculate_metrics(df):
    """Calculates key Google Ads metrics for a given DataFrame."""
    # Ensure 'Impressions' and 'Clicks' are not zero to avoid division by zero errors
    df['CTR'] = (df['Clicks'] / df['Impressions'].replace(0, np.nan)).fillna(0) * 100 # Click-Through Rate (%)
    df['CPC'] = (df['Cost'] / df['Clicks'].replace(0, np.nan)).fillna(0) # Cost Per Click
    df['Conversion Rate'] = (df['Conversions'] / df['Clicks'].replace(0, np.nan)).fillna(0) * 100 # Conversion Rate (%)
    df['CPA'] = (df['Cost'] / df['Conversions'].replace(0, np.nan)).replace([np.inf, -np.inf], np.nan).fillna(0) # Cost Per Acquisition
    df['ROAS'] = (df['Simulated Revenue'] / df['Cost'].replace(0, np.nan)).replace([np.inf, -np.inf], np.nan).fillna(0) # Return On Ad Spend
    return df

# Calculate metrics for the entire dataset
ads_df = calculate_metrics(ads_df.copy()) # Use a copy to avoid SettingWithCopyWarning

# --- 3. Performance Aggregation and Time-Series Analysis Module ---
# This module handles aggregating data at different levels (daily, weekly, monthly, demographic)
# and performing time-based comparisons.

def aggregate_daily_performance(df):
    """Aggregates performance metrics on a daily basis."""
    daily_agg = df.groupby('Date').agg({
        'Impressions': 'sum',
        'Clicks': 'sum',
        'Cost': 'sum',
        'Conversions': 'sum',
        # Corrected syntax: key is string literal, value is aggregation function
        'Simulated Revenue': 'sum' 
    }).reset_index()
    return calculate_metrics(daily_agg)

def aggregate_campaign_performance(df):
    """Aggregates performance metrics by campaign."""
    campaign_agg = df.groupby('Campaign Name').agg({
        'Impressions': 'sum',
        'Clicks': 'sum',
        'Cost': 'sum',
        'Conversions': 'sum',
        # Corrected syntax
        'Simulated Revenue': 'sum'
    }).reset_index()
    return calculate_metrics(campaign_agg)

def aggregate_ad_group_performance(df):
    """Aggregates performance metrics by ad group."""
    ad_group_agg = df.groupby(['Campaign Name', 'Ad Group Name']).agg({
        'Impressions': 'sum',
        'Clicks': 'sum',
        'Cost': 'sum',
        'Conversions': 'sum',
        # Corrected syntax
        'Simulated Revenue': 'sum'
    }).reset_index()
    return calculate_metrics(ad_group_agg)

def aggregate_keyword_performance(df):
    """Aggregates performance metrics by keyword."""
    keyword_agg = df.groupby(['Campaign Name', 'Ad Group Name', 'Keyword']).agg({
        'Impressions': 'sum',
        'Clicks': 'sum',
        'Cost': 'sum',
        'Conversions': 'sum',
        # Corrected syntax
        'Simulated Revenue': 'sum'
    }).reset_index()
    return calculate_metrics(keyword_agg)

def aggregate_demographic_performance(df, group_by_col):
    """Aggregates performance metrics by a demographic column."""
    demographic_agg = df.groupby(group_by_col).agg({
        'Impressions': 'sum',
        'Clicks': 'sum',
        'Cost': 'sum',
        'Conversions': 'sum',
        # Corrected syntax
        'Simulated Revenue': 'sum'
    }).reset_index()
    return calculate_metrics(demographic_agg)

def calculate_period_over_period_change(df, period_type='week'):
    """
    Calculates week-over-week or month-over-month changes for key metrics.
    Assumes 'Date' column is present and is datetime type.
    """
    df_copy = df.copy() # Work on a copy to avoid SettingWithCopyWarning
    if period_type == 'week':
        df_copy['Period'] = df_copy['Date'].dt.to_period('W').dt.start_time
    elif period_type == 'month':
        df_copy['Period'] = df_copy['Date'].dt.to_period('M').dt.start_time
    else:
        raise ValueError("period_type must be 'week' or 'month'")

    period_agg = df_copy.groupby('Period').agg({
        'Impressions': 'sum',
        'Clicks': 'sum',
        'Cost': 'sum',
        'Conversions': 'sum',
        # Corrected syntax
        'Simulated Revenue': 'sum'
    }).reset_index()
    period_agg = calculate_metrics(period_agg)

    # Calculate WoW/MoM change
    for col in ['Clicks', 'Cost', 'Conversions', 'CTR', 'CPA', 'ROAS']:
        period_agg[f'{col}_Prev_Period'] = period_agg[col].shift(1)
        # Handle division by zero for percentage change
        denominator = period_agg[f'{col}_Prev_Period'].replace(0, np.nan) # Replace 0 with NaN for division
        period_agg[f'{col}_Change_Pct'] = (
            (period_agg[col] - period_agg[f'{col}_Prev_Period']) / denominator
        ) * 100
        # If previous period was 0 and current is > 0, set change to 10000% (arbitrarily large)
        period_agg.loc[(period_agg[f'{col}_Prev_Period'] == 0) & (period_agg[col] > 0), f'{col}_Change_Pct'] = 10000 
        period_agg[f'{col}_Change_Pct'] = period_agg[f'{col}_Change_Pct'].fillna(0) # Fill remaining NaNs (e.g., 0/0) with 0

    return period_agg.dropna(subset=[f'Clicks_Prev_Period']) # Remove first period with no previous data

# --- 4. Anomaly Detection Module ---
# This module implements a more sophisticated anomaly detection using Z-scores.

def detect_anomalies(df, metric_col, window_size, z_score_threshold):
    """
    Detects anomalies in a given metric using a rolling Z-score method.
    Anomalies are flagged if the metric value is significantly below the rolling mean.
    """
    df_copy = df.copy() # Work on a copy to avoid modifying original DataFrame
    
    # Calculate rolling mean and standard deviation
    df_copy[f'{metric_col}_Rolling_Mean'] = df_copy[metric_col].rolling(window=window_size, min_periods=1).mean()
    df_copy[f'{metric_col}_Rolling_Std'] = df_copy[metric_col].rolling(window=window_size, min_periods=1).std()

    # Calculate Z-score for drops (negative deviation from mean)
    # Handle cases where std dev is zero to avoid division by zero
    std_dev_safe = df_copy[f'{metric_col}_Rolling_Std'].replace(0, np.nan)
    df_copy[f'{metric_col}_Z_Score'] = (
        (df_copy[metric_col] - df_copy[f'{metric_col}_Rolling_Mean']) / std_dev_safe
    ).fillna(0) # Fill NaN from division by zero or initial periods

    # Flag anomalies: metric is significantly below the rolling mean
    # We look for negative Z-scores that exceed the absolute threshold
    anomalies = df_copy[
        (df_copy[f'{metric_col}_Z_Score'] < -z_score_threshold) &
        (df_copy[f'{metric_col}_Rolling_Std'] > 0) # Only consider if there's some variation
    ].copy()

    return anomalies[['Date', metric_col, f'{metric_col}_Rolling_Mean', f'{metric_col}_Z_Score']]

# --- 5. Budget Pacing Simulation Module ---
# A simple module to simulate and track budget pacing.

def simulate_budget_pacing(df, total_budget, start_date, end_date):
    """
    Simulates budget pacing over a period.
    Assumes daily cost data is available.
    """
    total_days = (end_date - start_date).days + 1
    if total_days == 0: # Handle case of single day or invalid range
        return pd.DataFrame()
    daily_budget_target = total_budget / total_days
    
    pacing_data = []
    cumulative_cost = 0
    
    for _, row in df.sort_values(by='Date').iterrows():
        cumulative_cost += row['Cost']
        days_passed = (row['Date'] - start_date).days + 1
        
        # Calculate expected cumulative budget based on days passed
        expected_cumulative_budget = daily_budget_target * days_passed
        
        pacing_data.append({
            'Date': row['Date'],
            'Daily Cost': row['Cost'],
            'Cumulative Cost': cumulative_cost,
            'Expected Cumulative Budget': expected_cumulative_budget,
            'Budget Variance': cumulative_cost - expected_cumulative_budget
        })
    return pd.DataFrame(pacing_data)

# --- 6. Optimization Insights Module (Simulated) ---
# This module provides simulated recommendations based on performance rules.

def generate_optimization_insights(campaign_perf_df, daily_anomalies_df, demographic_perf_age, demographic_perf_gender):
    """
    Generates simulated optimization insights based on aggregated performance, anomalies, and demographics.
    """
    insights = []

    insights.append("\n--- Simulated Optimization Insights & Recommendations ---")
    
    # Overall performance insights
    avg_cpa = campaign_perf_df['CPA'].mean()
    avg_roas = campaign_perf_df['ROAS'].mean()
    
    insights.append(f"Overall Average CPA: ${avg_cpa:.2f}")
    insights.append(f"Overall Average ROAS: {avg_roas:.2f}x")

    if avg_cpa > 50: # Example threshold
        insights.append("- Recommendation: Review campaigns with high CPA. Consider refining targeting or improving ad copy.")
    if avg_roas < 2.0: # Example threshold
        insights.append("- Recommendation: Focus on improving ROAS for underperforming campaigns. Look into conversion tracking accuracy.")

    # Campaign-specific insights
    insights.append("\nCampaign-Specific Insights:")
    for _, row in campaign_perf_df.iterrows():
        if row['CPA'] > 70 and row['Conversions'] > 0:
            insights.append(f"- Campaign '{row['Campaign Name']}': High CPA (${row['CPA']:.2f}). Investigate keyword bids or ad group performance.")
        if row['ROAS'] < 1.5 and row['Simulated Revenue'] > 0:
            insights.append(f"- Campaign '{row['Campaign Name']}': Low ROAS ({row['ROAS']:.2f}x). Consider pausing underperforming keywords/ad groups.")
        if row['CTR'] < 1.0:
            insights.append(f"- Campaign '{row['Campaign Name']}': Low CTR ({row['CTR']:.2f}%). Improve ad relevance or expand keyword coverage.")
    
    # Anomaly-driven insights
    if not daily_anomalies_df.empty:
        insights.append("\nAnomaly-Driven Insights:")
        for _, row in daily_anomalies_df.iterrows():
            metric_impacted = row.index[1].replace('_', ' ').title()
            insights.append(f"- On {row['Date'].strftime('%Y-%m-%d')}: Significant drop in {metric_impacted}. Investigate external factors or account changes.")
    else:
        insights.append("\nNo significant anomalies detected, indicating stable performance.")

    # Demographic insights
    insights.append("\nDemographic Insights:")
    # Identify best performing age group by Conversion Rate
    best_age_group = demographic_perf_age.loc[demographic_perf_age['Conversion Rate'].idxmax()]
    insights.append(f"- Best performing Age Group by Conversion Rate: '{best_age_group['Age Group']}' ({best_age_group['Conversion Rate']:.2f}%)")
    
    # Identify lowest performing gender by CPA (if conversions > 0)
    if not demographic_perf_gender[demographic_perf_gender['Conversions'] > 0].empty:
        worst_gender_cpa = demographic_perf_gender[demographic_perf_gender['Conversions'] > 0].loc[demographic_perf_gender[demographic_perf_gender['Conversions'] > 0]['CPA'].idxmax()]
        insights.append(f"- Consider reviewing targeting for Gender '{worst_gender_cpa['Gender']}' due to high CPA (${worst_gender_cpa['CPA']:.2f}).")
    else:
        insights.append("- Not enough conversion data to provide gender-specific CPA insights.")

    return "\n".join(insights)

# --- 7. Data Visualization Module ---
# This module contains functions to generate and save various plots.

def create_and_save_plot(df, x_col, y_col, title, filename, plot_type='line', hue_col=None, y_label=None, x_label=None, rotation=0):
    """
    Generates a plot (line or bar) and saves it to the data_exports directory.
    """
    output_dir = 'data_exports'
    os.makedirs(output_dir, exist_ok=True)
    filepath = os.path.join(output_dir, filename)

    plt.figure(figsize=(12, 6))
    sns.set_style("whitegrid") # Set a nice seaborn style

    if plot_type == 'line':
        sns.lineplot(data=df, x=x_col, y=y_col, hue=hue_col)
    elif plot_type == 'bar':
        sns.barplot(data=df, x=x_col, y=y_col, hue=hue_col)
        if rotation > 0:
            plt.xticks(rotation=rotation, ha='right') # Rotate x-axis labels for readability
    
    plt.title(title, fontsize=16)
    plt.xlabel(x_label if x_label else x_col.replace('_', ' ').title(), fontsize=12)
    plt.ylabel(y_label if y_label else y_col.replace('_', ' ').title(), fontsize=12)
    plt.tight_layout() # Adjust layout to prevent labels from overlapping
    
    try:
        plt.savefig(filepath)
        print(f"Graph saved: {filepath}")
    except Exception as e:
        print(f"Error saving graph {filename}: {e}")
    finally:
        plt.close() # Close the plot to free up memory

# --- 8. Reporting Module ---
# This module orchestrates the analysis, prints reports, and calls plotting functions.

def generate_full_report(df):
    """Generates and prints a comprehensive Google Ads performance report, including graphs."""
    print("\n--- Generating Comprehensive Google Ads Performance Report ---")

    # 8.1 Overall Daily Performance
    daily_performance_df = aggregate_daily_performance(df)
    print("\n--- Daily Aggregated Performance (Last 7 Days) ---")
    print(daily_performance_df.tail(7).round(2).to_string())
    
    # Create and save daily trend graphs
    create_and_save_plot(daily_performance_df, 'Date', 'Clicks', 'Daily Clicks Trend', 'daily_clicks_trend.png', y_label='Total Clicks')
    create_and_save_plot(daily_performance_df, 'Date', 'Conversions', 'Daily Conversions Trend', 'daily_conversions_trend.png', y_label='Total Conversions')
    create_and_save_plot(daily_performance_df, 'Date', 'Cost', 'Daily Cost Trend', 'daily_cost_trend.png', y_label='Total Cost ($)')


    # 8.2 Campaign Performance Summary
    campaign_performance_df = aggregate_campaign_performance(df)
    print("\n--- Campaign Performance Summary ---")
    print(campaign_performance_df.round(2).to_string())
    
    # Create and save campaign performance graphs
    create_and_save_plot(campaign_performance_df.sort_values(by='Conversions', ascending=False).head(5),
                         'Campaign Name', 'Conversions', 'Top 5 Campaigns by Conversions', 'top_campaigns_conversions.png',
                         plot_type='bar', rotation=45)
    create_and_save_plot(campaign_performance_df.sort_values(by='CPA', ascending=False).head(5),
                         'Campaign Name', 'CPA', 'Top 5 Campaigns by CPA (Highest)', 'top_campaigns_cpa.png',
                         plot_type='bar', rotation=45, y_label='Cost Per Acquisition ($)')


    # 8.3 Ad Group Performance Summary (Top 5 by Conversions)
    ad_group_performance_df = aggregate_ad_group_performance(df)
    print("\n--- Top 5 Ad Groups by Conversions ---")
    print(ad_group_performance_df.sort_values(by='Conversions', ascending=False).head(5).round(2).to_string())

    # 8.4 Keyword Performance Summary (Top 5 by Clicks)
    keyword_performance_df = aggregate_keyword_performance(df)
    print("\n--- Top 5 Keywords by Clicks ---")
    print(keyword_performance_df.sort_values(by='Clicks', ascending=False).head(5).round(2).to_string())

    # 8.5 Week-over-Week Performance Change
    wow_performance_df = calculate_period_over_period_change(df, period_type='week')
    print("\n--- Week-over-Week Performance Changes (Last 3 Weeks) ---")
    # Display relevant columns for WoW
    print(wow_performance_df[['Period', 'Clicks', 'Clicks_Change_Pct', 'Conversions', 'Conversions_Change_Pct', 'Cost', 'Cost_Change_Pct', 'ROAS', 'ROAS_Change_Pct']].tail(3).round(2).to_string())

    # 8.6 Anomaly Detection Report
    print("\n--- Anomaly Detection Report ---")
    clicks_anomalies = detect_anomalies(daily_performance_df, 'Clicks', ANOMALY_DETECTION_WINDOW_SIZE, Z_SCORE_THRESHOLD)
    conversions_anomalies = detect_anomalies(daily_performance_df, 'Conversions', ANOMALY_DETECTION_WINDOW_SIZE, Z_SCORE_THRESHOLD)

    all_anomalies = pd.concat([clicks_anomalies, conversions_anomalies]).drop_duplicates(subset=['Date']).sort_values(by='Date')

    if not all_anomalies.empty:
        print(f"Identified {len(all_anomalies)} potential anomaly days:")
        for _, row in all_anomalies.iterrows():
            metric_impacted = row.index[1].replace('_', ' ').title() # e.g., 'Clicks' or 'Conversions'
            print(f"- Date: {row['Date'].strftime('%Y-%m-%d')}, Metric: {metric_impacted}, Value: {row[metric_impacted]:,.0f}, Z-Score: {row[f'{metric_impacted}_Z_Score']:.2f}")
    else:
        print("No significant anomalies detected based on the Z-score threshold.")

    # 8.7 Budget Pacing Report
    print("\n--- Budget Pacing Simulation ---")
    total_budget_for_period = daily_performance_df['Cost'].sum() * 1.1 # Assume 10% buffer for total budget
    budget_pacing_df = simulate_budget_pacing(daily_performance_df, 
                                              total_budget_for_period, 
                                              df['Date'].min(), 
                                              df['Date'].max())
    if not budget_pacing_df.empty:
        print("Daily Budget Target: $", round(total_budget_for_period / NUM_DAYS, 2))
        print("Cumulative Cost vs. Expected Budget (Last 7 Days):")
        print(budget_pacing_df[['Date', 'Cumulative Cost', 'Expected Cumulative Budget', 'Budget Variance']].tail(7).round(2).to_string())
        
        final_budget_variance = budget_pacing_df['Budget Variance'].iloc[-1] if not budget_pacing_df.empty else 0
        print(f"\nFinal Budget Variance: ${final_budget_variance:.2f} (positive means overspent, negative means underspent)")
    else:
        print("Not enough data to simulate budget pacing.")

    # 8.8 Demographic Performance
    demographic_perf_age = aggregate_demographic_performance(df, 'Age Group')
    demographic_perf_gender = aggregate_demographic_performance(df, 'Gender')
    demographic_perf_location = aggregate_demographic_performance(df, 'Location')

    print("\n--- Demographic Performance (by Age Group) ---")
    print(demographic_perf_age.sort_values(by='Conversions', ascending=False).round(2).to_string())
    print("\n--- Demographic Performance (by Gender) ---")
    print(demographic_perf_gender.sort_values(by='Conversions', ascending=False).round(2).to_string())
    print("\n--- Demographic Performance (by Location) ---")
    print(demographic_perf_location.sort_values(by='Conversions', ascending=False).round(2).to_string())

    # Create and save demographic performance graphs
    create_and_save_plot(demographic_perf_age, 'Age Group', 'Conversions', 'Conversions by Age Group', 'conversions_by_age.png', plot_type='bar')
    create_and_save_plot(demographic_perf_gender, 'Gender', 'Conversions', 'Conversions by Gender', 'conversions_by_gender.png', plot_type='bar')
    create_and_save_plot(demographic_perf_location, 'Location', 'Conversions', 'Conversions by Location', 'conversions_by_location.png', plot_type='bar')
    create_and_save_plot(demographic_perf_age, 'Age Group', 'CPA', 'CPA by Age Group', 'cpa_by_age.png', plot_type='bar', y_label='Cost Per Acquisition ($)')


    # 8.9 Optimization Insights
    insights_report = generate_optimization_insights(campaign_performance_df, all_anomalies, demographic_perf_age, demographic_perf_gender)
    print(insights_report)

    return daily_performance_df, campaign_performance_df, ad_group_performance_df, keyword_performance_df, wow_performance_df, all_anomalies, budget_pacing_df, demographic_perf_age, demographic_perf_gender, demographic_perf_location

# --- Main Execution Flow ---
if __name__ == "__main__":
    print("\n--- Starting Data Generation ---")
    # ads_df is already generated globally at the top of the script.
    # This allows all functions to access the primary DataFrame.
    print(f"Generated {len(ads_df):,} rows of simulated Google Ads data.")

    # Generate the comprehensive report and get all aggregated DataFrames
    daily_perf, campaign_perf, ad_group_perf, keyword_perf, wow_perf, anomalies_df, budget_pacing_df, demographic_perf_age, demographic_perf_gender, demographic_perf_location = generate_full_report(ads_df)

    # --- 9. Export Raw Data and Aggregated Reports to XLSX ---
    # This section saves the simulated raw data and various aggregated reports to an XLSX file.
    # This file can be uploaded to GitHub to show the data source and analysis results.

    output_dir = 'data_exports'
    os.makedirs(output_dir, exist_ok=True) # Create directory if it doesn't exist
    xlsx_filename = os.path.join(output_dir, 'google_ads_comprehensive_simulated_data_with_demographics.xlsx')

    try:
        wb = Workbook()
        
        # Sheet 1: Raw Simulated Data
        ws_raw = wb.active
        ws_raw.title = "Raw Ads Data"
        ws_raw.append(ads_df.columns.tolist())
        for r_idx, row in enumerate(dataframe_to_rows(ads_df, index=False, header=False)):
            ws_raw.append(row)

        # Sheet 2: Daily Aggregated Performance
        ws_daily = wb.create_sheet("Daily Performance")
        ws_daily.append(daily_perf.columns.tolist())
        for r_idx, row in enumerate(dataframe_to_rows(daily_perf, index=False, header=False)):
            ws_daily.append(row)

        # Sheet 3: Campaign Performance
        ws_campaign = wb.create_sheet("Campaign Performance")
        ws_campaign.append(campaign_perf.columns.tolist())
        for r_idx, row in enumerate(dataframe_to_rows(campaign_perf, index=False, header=False)):
            ws_campaign.append(row)

        # Sheet 4: Ad Group Performance
        ws_adgroup = wb.create_sheet("Ad Group Performance")
        ws_adgroup.append(ad_group_perf.columns.tolist())
        for r_idx, row in enumerate(dataframe_to_rows(ad_group_perf, index=False, header=False)):
            ws_adgroup.append(row)

        # Sheet 5: Keyword Performance
        ws_keyword = wb.create_sheet("Keyword Performance")
        ws_keyword.append(keyword_perf.columns.tolist())
        for r_idx, row in enumerate(dataframe_to_rows(keyword_perf, index=False, header=False)):
            ws_keyword.append(row)
            
        # Sheet 6: Anomaly Report
        if not anomalies_df.empty:
            ws_anomalies = wb.create_sheet("Anomaly Report")
            ws_anomalies.append(anomalies_df.columns.tolist())
            for r_idx, row in enumerate(dataframe_to_rows(anomalies_df, index=False, header=False)):
                ws_anomalies.append(row)

        # Sheet 7: Budget Pacing
        if not budget_pacing_df.empty:
            ws_pacing = wb.create_sheet("Budget Pacing")
            ws_pacing.append(budget_pacing_df.columns.tolist())
            for r_idx, row in enumerate(dataframe_to_rows(budget_pacing_df, index=False, header=False)):
                ws_pacing.append(row)

        # Sheet 8: Demographic Performance (Age)
        ws_age_demo = wb.create_sheet("Demographics - Age")
        ws_age_demo.append(demographic_perf_age.columns.tolist())
        for r_idx, row in enumerate(dataframe_to_rows(demographic_perf_age, index=False, header=False)):
            ws_age_demo.append(row)

        # Sheet 9: Demographic Performance (Gender)
        ws_gender_demo = wb.create_sheet("Demographics - Gender")
        ws_gender_demo.append(demographic_perf_gender.columns.tolist())
        for r_idx, row in enumerate(dataframe_to_rows(demographic_perf_gender, index=False, header=False)):
            ws_gender_demo.append(row)

        # Sheet 10: Demographic Performance (Location)
        ws_location_demo = wb.create_sheet("Demographics - Location")
        ws_location_demo.append(demographic_perf_location.columns.tolist())
        for r_idx, row in enumerate(dataframe_to_rows(demographic_perf_location, index=False, header=False)):
            ws_location_demo.append(row)


        wb.save(xlsx_filename)
        print(f"\nSimulated Google Ads data and aggregated reports successfully exported to: {xlsx_filename}")
        print("This .xlsx file contains multiple sheets with raw data and various aggregated views.")
        print("You can upload this .xlsx file to GitHub to show your data source and analysis results.")

    except Exception as e:
        print(f"Error exporting data to XLSX: {e}")

    print("\n--- Comprehensive Google Ads Performance Analysis Project Finished ---")
