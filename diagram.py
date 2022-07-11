from tabnanny import verbose
import pandas as pd
from openpyxl import load_workbook

# Grab Dataframe
df = pd.read_excel('test.xlsx', sheet_name=None)

# Plot the Dataframe - types = scatter, line, bar, plot.pie
df.plot(x = 'COUNTRY', y = 'POP', type='bar')

plt.show()