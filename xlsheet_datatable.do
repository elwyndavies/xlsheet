python:
import __main__ as m
m.max_row = len(m.df)
m.max_col = len(m.df.columns)





# Add formatting
for x in range(m.row_offset+1, m.row_offset+m.max_row+1):
	if m.collapse == True:
		m.worksheet.set_row(x, None, m.cell_format, {'level': 1, 'hidden': True})
	else:
		m.worksheet.set_row(x, None, m.cell_format)
	
# Collapse dataset as well

if m.collapse == True:
	m.worksheet.set_row(m.row_offset, None, None, {'level': 1, 'hidden': True})
	m.worksheet.set_row(m.row_offset+m.max_row+1, None, None, {'collapsed': True})
		

end
