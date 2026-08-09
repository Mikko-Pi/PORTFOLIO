import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

df_table = xl("C:D", headers=True)
avg = df_table["Monthly Average"].mean()

sns.set_theme(style='whitegrid')
ax = sns.barplot(df_table,x='End of Month', y='Monthly Average')

ax.tick_params(axis='x',rotation=80,labelsize=6)

ax.tick_params(axis='y', labelsize=8)
ax.axhline(avg, color='black', linestyle='--', linewidth=1, zorder=10)
ax.text(
    x=0.02,
    y=avg + (avg * 0.01),        # small upward offset (1% of value)
    s=f"{avg:.0f} €",
    color='black',
    va='bottom',                 # anchor text BELOW the point → text appears above
    ha='left',
    fontsize=8,
    backgroundcolor='white'
)

ax

#ANOTHER CODE STARTS FROM HERE:

import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

df_table=xl("A:B", headers=True)
avg=df_table["Account events"].mean()
sns.set_theme(style='whitegrid')
ax = sns.barplot(df_table,x='End of Month', y='Account events')

ax.tick_params(axis='x',rotation=80,labelsize=6)

ax.tick_params(axis='y', labelsize=8)
ax.axhline(avg, color='black', linestyle='--', linewidth=1, zorder=10)
ax.text(
    x=0.02,
    y=avg + (avg * 0.01),        # small upward offset (1% of value)
    s=f"{avg:.0f}",
    color='black',
    va='bottom',                 # anchor text BELOW the point → text appears above
    ha='left',
    fontsize=8,
    backgroundcolor='white'
)

ax