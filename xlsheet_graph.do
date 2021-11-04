python:
import __main__ as m
m.max_row = len(m.df)
m.max_col = len(m.df.columns)

from xlsxwriter.utility import xl_rowcol_to_cell
from xlsxwriter.utility import xl_range_abs

# m.col = 2

m.cat = 0


# Check custom data labels

m.custom_data_labels = []


# If custom data labels are provided, add them. The data label is specified in the first column of the dataset.

if m.use_custom_data_labels == True:
	for x in range(m.row_offset+1, m.row_offset+m.max_row+1):
		m.col = 0
		m.row = x
		m.custom_data_labels.append({'value': "='" + m.sheet_name + "'!" + xl_rowcol_to_cell(m.row, m.col, row_abs=True, col_abs=True)})
	
	m.series_options['data_labels']['custom'] = m.custom_data_labels




# Create a chart object.

def generate_chart(m):
	chart = m.workbook.add_chart({'type': m.type, 'subtype': m.subtype})

	# Configure the series of the chart from the dataframe data.


	
	if m.type == "scatter" and m.alternate == True:
	
		# for scatter plots, and alternate is True, the data should be structered as follows:
		#   A  |   B  |   C  |   D  |   E  |  
		#      | SERIES 1    |  SERIES 2   |
		#      | XVAR | YVAR | XVAR | YVAR |
		
		for x in range(m.colstart, m.colend+1, 2):
			print(x)
			
			m.col = x - 1
			m.series = m.series_options
			
			m.series['name'] = [m.sheet_name, m.row_offset, m.col]
			m.series['categories'] = [m.sheet_name, m.row_offset+1, m.col, m.row_offset+m.max_row, m.col]
			m.series['values'] = [m.sheet_name, m.row_offset+1, m.col+1, m.row_offset+m.max_row, m.col+1]

			chart.add_series(m.series)
	
	
	# If type is not scatter:
	
	
	# If error bars are present, they should be coded as follows:
	#   A  |   B  |   C  |   D  |   E  |   F  |   G   |
	#      | SERIES 1           |  SERIES 2           |
	#      | VAL  | LOW  | HIGH | VAL  | LOW  | HIGH  |
	
	
	elif m.transpose == False and m.errorbars == True:
		if m.errorbarsymmetry == True:
			m.step = 2
		else:
			m.step = 3
	
		for x in range(m.colstart, m.colend+1, m.step):
			m.col = x - 1
			m.series = m.series_options
			m.series['name'] = [m.sheet_name, m.row_offset, m.col]
			m.series['categories'] = [m.sheet_name, m.row_offset+1, m.cat, m.row_offset+m.max_row, m.cat+m.catcols-1]
			m.series['values'] = [m.sheet_name, m.row_offset+1, m.col, m.row_offset+m.max_row, m.col]
			
			m.bartype = 'y_error_bars'
			if m.xbars == True:
				m.bartype = 'x_error_bars'
			
			
			if m.errorbarsymmetry == True:
				# m.series[m.bartype] = {'type': 'custom', 'minus_values': [m.sheet_name, m.row_offset+1, m.col+1, m.row_offset+m.max_row, m.col+1], 'plus_values': [m.sheet_name, m.row_offset+1, m.col+1, m.row_offset+m.max_row, m.col+1]}
				m.series[m.bartype] = {'type': 'custom', 'minus_values': "='" + m.sheet_name + "'!" + xl_range_abs(m.row_offset+1, m.col+1, m.row_offset+m.max_row, m.col+1), 'plus_values': "='" + m.sheet_name + "'!" + xl_range_abs(m.row_offset+1, m.col+1, m.row_offset+m.max_row, m.col+1)}
				
				
			else:
				m.series[m.bartype] = {'type': 'custom', 'minus_values': "='" + m.sheet_name + "'!" + xl_range_abs(m.row_offset+1, m.col+1, m.row_offset+m.max_row, m.col+1), 'plus_values': "='" + m.sheet_name + "'!" + xl_range_abs(m.row_offset+1, m.col+2, m.row_offset+m.max_row, m.col+2)}
				
				
				# m.series[m.bartype] = {'type': 'custom', 'minus_values': [m.sheet_name, m.row_offset+1, m.col+1, m.row_offset+m.max_row, m.col+1], 'plus_values': [m.sheet_name, m.row_offset+1, m.col+2, m.row_offset+m.max_row, m.col+2]}

			chart.add_series(m.series)

		
	
	
	elif m.transpose == False:
		for x in range(m.colstart, m.colend+1):
			m.col = x - 1
			m.series = m.series_options
			m.series['name'] = [m.sheet_name, m.row_offset, m.col]
			m.series['categories'] = [m.sheet_name, m.row_offset+1, m.cat, m.row_offset+m.max_row, m.cat+m.catcols-1]
			m.series['values'] = [m.sheet_name, m.row_offset+1, m.col, m.row_offset+m.max_row, m.col]

			chart.add_series(m.series)

	# If transpose is true, use rows instead
	elif m.transpose == True:
		for x in range(m.row_offset+1, m.row_offset+m.max_row+1):
			m.col = 0
			m.row = x
			m.series = m.series_options
			
			m.series['name'] = [m.sheet_name, m.row, m.col, m.row, m.catcols-1]
			m.series['categories'] = [m.sheet_name, m.row_offset, m.catcols, m.row_offset, m.colend-1]
			m.series['values'] = [m.sheet_name, m.row, m.catcols, m.row, m.colend-1]
			chart.add_series(m.series)


			
	if m.title != "":
		chart.set_title ({'name': m.title})
	elif m.title == "none":
		chart.set_title ({'none': True})


	chart.set_style(2)
		
	# Insert the chart into the worksheet.

	chart.set_x_axis(m.x_axis_options)
	chart.set_y_axis(m.y_axis_options)

	chart.set_legend({'position': m.legend_position})
	
	# Set the size. Standard size is 5.0 by 3.0 inches.
	
	chart.set_size({'x_scale': m.xsize/5.0 , 'y_scale': m.ysize/3.0})

	return chart

m.chart = generate_chart(m)
m.worksheet.insert_chart(m.row_offset - 3, m.col_offset, m.chart)


print("Creating graph")

#if m.usegridsheet == True:

#	m.gridchart = generate_chart(m)

#	m.gridsheet.insert_chart(m.gridy, m.gridx, m.gridchart)

#	m.gridy = m.gridy + 15


end
