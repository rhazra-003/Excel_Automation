wb = load_workbook('report_2021.xlsx')
sheet = wb['Report']

# barchart
barchart = BarChart()

#locate data and categories
data = Reference(sheet,
                 min_col=min_column+1,
                 max_col=max_column,
                 min_row=min_row,
                 max_row=max_row) #including headers
categories = Reference(sheet,
                       min_col=min_column,
                       max_col=min_column,
                       min_row=min_row+1,
                       max_row=max_row) #not including headers

# adding data and categories
barchart.add_data(data, titles_from_data=True)
barchart.set_categories(categories)

#location chart
sheet.add_chart(barchart, "B12")
barchart.title = 'Sales by Product line'
barchart.style = 5 #choose the chart style

wb.save('report_2021.xlsx')
