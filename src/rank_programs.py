import pandas as pd
import matplotlib.pyplot as plt
import os

# Ensure outputs folder exists
os.makedirs("outputs", exist_ok=True)

# Read dataset
file_path = "data/Grad Program Exit Survey Data (2).xlsx"
df = pd.read_excel(file_path)

# Drop first row (Qualtrics metadata)
df = df.iloc[1:]

# Select ranking columns
rank_cols = [col for col in df.columns if "Rank" in col]

# Convert to numeric
df[rank_cols] = df[rank_cols].apply(pd.to_numeric, errors="coerce")

# Calculate average rank
mean_ranks = df[rank_cols].mean().sort_values()

# Clean column names for readability
mean_ranks.index = mean_ranks.index.str.replace(" - Rank", "")

# Save ranking to CSV
mean_ranks.to_csv("outputs/rank_order.csv")

# Create bar chart
plt.figure()
mean_ranks.plot(kind="bar")
plt.ylabel("Average Rank (Lower = Better)")
plt.title("Course Ranking Based on Student Preferences")
plt.xticks(rotation=45, ha="right")
plt.tight_layout()

# Save figure
plt.savefig("outputs/rank_order.png")

