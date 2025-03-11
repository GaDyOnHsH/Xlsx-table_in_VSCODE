from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

# Create a new Excel workbook
wb = Workbook()
ws = wb.active

# Add data to the worksheet
data = [
    ["A", "B", "C", "D"],
    ["item1", 4, 5, 6],
    ["item2", 7, 8, 9],
    ["item3", 10, 11, 12]
    ]

for row in data:
    ws.append(row)

# Create a bar chart
chart = BarChart()
chart.title = "random_name"
chart.x_axis.title = "X"
chart.y_axis.title = "Y"

data = Reference(ws, min_col=2, min_row=1, max_col=4, max_row=len(data))
categories = Reference(ws, min_col=1, min_row=2, max_row=len(data))

chart.add_data(data, titles_from_data=True)
chart.set_categories(categories)

ws.add_chart(chart, "E2")

wb.save("random_name.xlsx")