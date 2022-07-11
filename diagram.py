from tabnanny import verbose
from pyxll import xl_func, plot
import pandas as pd
from openpyxl import load_workbook
import matplotlib.pyplot as plt
import xlsxwriter


# Notes: If you specify sheet_name in read_excel it turns it into a dict, if not it will be a dataframe

# Grab Dataframe
df = pd.read_excel('test.xlsx')

# Plot the Dataframe - kind  = scatter, line, bar, plot.pie
# for plot have options of: title= , color={'line'='color'}, subplots = T/F, sharex= T/F, sharey = T/F, layout= (rows, columns), figize=(2,3), use_index=T/F, grid=T/F, legend = T/F, style=[]

# Plots a bar graph
ax = df.plot(x = 'COUNTRY', y = 'POP', kind ='bar')
plt.show()

fig = ax.get_figure()
# sht.pictures.add(
#     fig,
#     name = "Pandas",
#     update = True,
#     left = sht.range("A21").left,
#     top = sht.range("A21").top,
#     height = 200,
#     width = 300,
# )

plt.tight_layout()
fig.savefig('graph.jpg', dpi=199)

workbook = xlsxwriter.Workbook('test.xlsx')
worksheet = workbook.add_worksheet()
worksheet.insert_image('L2', 'graph.jpg')
