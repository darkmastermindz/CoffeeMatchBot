## CoffeeMatchShuffle
# Takes a excel file with labels, Name and Email and matches people randomly
# - Created by Hansel Wei
##

import pandas as pd
import numpy as np

df = pd.read_excel(open('input.xlsx', 'rb'))

## OPTIONS:
MAX_LEN = len(df)
OUTPUT_FILENAME = "output"
##

print("☕📋✨ Time to brew some coffee... your input.xlsx file reads:")
print(df, end="\n\n")
df0 = df.sample(frac=1)

print("☕💫✨ shuffling...")
print(df0, end="\n\n")
df0.columns.name = None
np0 = df0.to_numpy()

# Shuffle Split
print("💫✨☕ --- Splitting into groups and shuffle again!", end="\n\n")
dfs = np.array_split(df0, 2)

print("☕💠😊 Group 1:")
df_group0 = pd.DataFrame(dfs[0]).sample(frac=1).reset_index(drop = True)
df_group0.set_axis(["💠 Name", "💠 Email"], axis=1)
print(df_group0, end="\n\n")

print("☕🔶😊 Group 2:", end="\n\n")
df_group1 = pd.DataFrame(dfs[1]).sample(frac=1).reset_index(drop = True)
df_group1.set_axis(["🔶 Name", "🔶 Email"], axis=1)
print(df_group1, end="\n\n")

print("☕🔮✨ --- Result:", end="\n\n")
groups = [df_group0, df_group1]
result = pd.concat(groups, axis=1, join='inner')

print(result, end="\n\n")
print("☕🔮📋 --- Saving Result as Excel File", end="\n\n")
df.to_excel(OUTPUT_FILENAME + ".xlsx", index=False)

print(format("✨☕ Your Magical Coffee is Done! Check %s.xlsx file" % OUTPUT_FILENAME))
print("-Hansel", end="\n\n")
