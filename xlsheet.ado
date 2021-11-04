// XLSHEET
// Version 0.1
// Elwyn Davies edavies@worldbank.org


cap program drop xlsheet
program define xlsheet
	syntax anything [using /], [text(str) name(str) x(varlist) y(varlist) title(str) subtitle(str) colstart(int 2) colend(int 2) allcols catcols(int 1) type(str) subtype(str) reversey reversex DATALABels(str) DATALABELFormat(str) DATALABELPOSition(str) format(str) TRANSpose ALSOTRANSpose ALTernate xtitle(str) ytitle(str) LEGENDPOSition(str) legend(str asis) customdatalabels collapse xlwings ERRORBARs XERRorbars YERRorbars noSYMERRorbars XLABELPOSition(str) xsize(real 5) ysize(real 3) gap(int 0) order(str) *]
	
	* First check if Python is installed
	cap python query
	if `"r(libpath)"' != "" & "${blockpython}" != "1" {
		global xlsheet_usepython 1
	}
	else {
		global xlsheet_usepython 0
	}
	
	
	
	
	
	if "${xlsheet_usepython}" == "1" & "`xlwings'" == "" & "${xlsheet_usexlwings}" != "1" {
	
	
		python: import __main__ as m
		python: import pandas as pd
		


		if "`anything'" == "newbook" {
			* Generates a new workbook
			
			disp  "If this gives an error message, please use forward slashes (/) in the filename, instead of backslashes."
			
			python: m.writer = pd.ExcelWriter('`using'', engine='xlsxwriter')
			python: m.workbook  = m.writer.book
			
			
			/*
			python: m.gridsheet = m.workbook.add_worksheet("Grid")
			
			python: m.gridx = -8
			python: m.gridy = 0
			
			python: m.usegrid = False
			python: m.gridsheet.hide()
			*/
			
			
			global xlsheet_newbook 1
			
		}


		if "`anything'" == "newsheet" {
			
			python: m.next_row = 5
			
			
			python: m.sheet_name = "`name'"
			python: m.worksheet = m.workbook.add_worksheet(m.sheet_name)
			python: m.writer.sheets[m.sheet_name] = m.worksheet
			python: m.bold = m.workbook.add_format({'bold': True})
			
			python: m.worksheet.write_string(0,0, "`title'", m.bold)
			
			python: m.worksheet.set_column(0, 0, 22)
			
			
			/*
			python: m.gridx = m.gridx + 8
			python: m.gridy = 0
			*/
			
			python: m.worksheet.activate()
			
			
			global xlsheet_newbook "`name'"
			
		}


		if "`anything'" == "text" {
			* Write some text
		
			python: m.row_offset = m.next_row
			python: m.worksheet.write_string(m.row_offset - 3 ,0, "`text'")

			
			python: m.next_row = m.row_offset + 1
		}


		if "`anything'" == "importdata" {
		
			* Imports the data as a table
		

			* Get the part after the comma
			gettoken left right : 0, parse(",") 
			gettoken left right : right, parse(",")
			char _dta[tablecmd] `right'
			
			//disp "right: `right'"
			

			python: m.row_offset = m.next_row
			python: m.collapse = False

			
			if "`using'" == "" {
				cap export excel using "temp.xlsx", firstrow(varlabels) replace
				python: m.df=pd.read_excel("temp.xlsx")	
			
			}
			else {
				* Use the using path:
				python: m.df=pd.read_excel("`using'")	
			
			}

			* Copy the data to the excel sheet
			python: m.df.to_excel(m.writer, sheet_name=m.sheet_name, startrow = m.row_offset, startcol = 0, index = False)
			python: m.worksheet = m.writer.sheets[m.sheet_name]
			
			
			* If the format option is set
			if "`format'" != "" {
				python: m.cell_format = m.workbook.add_format({'num_format': "`format'"})
			}
			else {
				python: m.cell_format = m.workbook.add_format()
			}
			
			if "`collapse'" != ""  python: m.collapse = True
			
			* Run the file xlsheet_datatable
			cap findfile xlsheet_datatable.do
			cap include "`r(fn)'"	
			
			** Specify data format:
			

			
			** python: m.worksheet.add_table(m.row_offset, 0, m.row_offset + len(m.df), len(m.df.columns)-1)
			
			* Add a title

			cap python: m.worksheet.write_string(m.row_offset - 3 ,0, "`title'", m.bold)
			cap python: m.worksheet.write_string(m.row_offset - 2 ,0, "`subtitle'")
			
			python: m.col_offset = max(len(m.df.columns) + 2, 6)
			
			* Specify the next gap
			
			python: m.next_row = m.row_offset + max(10, len(m.df)) + 5 + `gap'
			
			
			* cap python: m.gridsheet.write_string(m.gridy ,m.gridx, "`title'", m.bold)
			* cap python: m.gridsheet.write_string(m.gridy+1 ,m.gridx, "`subtitle'")
			
			* python: m.gridy = m.gridy + 2
			

		}

		if "`anything'" == "graph" | "`anything'" == "chart" {
		
			* Get the part after the comma
			gettoken left right : 0, parse(",") 
			gettoken left right : right, parse(",")
			char _dta[graphcmd] `right'

			python: m.catcols = `catcols'
			python: m.colstart = m.catcols + 1
			python: m.colend = len(m.df.columns)
			
			
			python: m.x_axis_options = {}
			python: m.y_axis_options = {}
			python: m.series_options = {}
			
			python: m.y_axis_options['major_gridlines'] = False
			python: m.x_axis_options['major_gridlines'] = False
			
			if "`xlabelposition'" != "" {
				python: m.x_axis_options['label_position'] = '`xlabelposition''
			}
			else {
				python: m.x_axis_options['label_position'] = 'low'
			}
			
			python: m.alternate = False
			python: m.use_custom_data_labels = False
			
			python: m.legend_position = "bottom"
			
			python: m.errorbars = False

			
			if "`reversex'" != ""   python: m.x_axis_options['reverse'] = True
			if "`reversey'" != ""   python: m.y_axis_options['reverse'] = True
			
			if "`alternate'" != ""  python: m.alternate = True
			
			
			if "`xtitle'" != ""  python: m.x_axis_options['name'] = "`xtitle'"
			if "`ytitle'" != ""  python: m.y_axis_options['name'] = "`ytitle'"
			
			
			* Legend position, valid options: top bottom left right overlay_left overlay_right none
			
			if inlist("`legendposition'", "top", "bottom", "left", "right", "overlay_left", "overlay_right", "none"){
				python: m.legend_position = "`legendposition'"
			}
			
			* ERRORBARs XERRorbars YERRorbars SYMERRorbars
			
			python: m.xbars = False
			python: m.errorbarsymmetry = True
			
			if "`errorbars'" != ""  python: m.errorbars = True
			if "`xerrorbars'" != ""  python: m.xbars = True
			if "`symerrorbars'" == "nosymerrorbars"  python: m.errorbarsymmetry = False

			
			if "`datalabels'" == "value" | "`datalabels'" == "custom" {
				python: m.series_options['data_labels'] = {}
				python: m.series_options['data_labels']['value'] = True
				if "`format'" != ""  python: m.series_options['data_labels']['num_format'] = "`format'"
				if "`datalabelformat'" != ""  python: m.series_options['data_labels']['num_format'] = "`datalabelformat'"
				if "`datalabelposition'" != ""  python: m.series_options['data_labels']['position'] = "`datalabelposition'"

			}
			if "`datalabels'" == "custom" {
				python: m.use_custom_data_labels = True
			
			}
			
			
			if "`type'" == "" {
				local type "column"
			}
			
			*Ensure every bar has a unit label
			if "`type'" == "column" {
				python: m.x_axis_options['interval_unit'] = 1
			}
			
			if "`type'" == "bar" {
				python: m.y_axis_options['interval_unit'] = 1
			}
			
			
			if "`transpose'" != "" {
				python: m.transpose = True 
			}
			else {
				python: m.transpose = False
			}
			

			
			python: m.type = "`type'"
			python: m.subtype = "`subtype'"
			python: m.title = "`title'"
			
			python: m.xsize = `xsize'
			python: m.ysize = `ysize'
			
			/*
			python: m.colstart = `colstart'
			python: m.colend = `colend'
			*/
			
			cap findfile xlsheet_graph.do
			cap include "`r(fn)'"
			
			python: m.col_offset = m.col_offset + 8
			
			* If the option is to transpose the graph, create another graph with the axes transposed.
			if "`alsotranspose'" != "" & "`transpose'" == "" {
				xlsheet `0' transpose
			}
			

		}


		if "`anything'" == "close" {
			python: m.writer.save()
		}

	}
	else if "`xlwings'" != "" | "${xlsheet_usexlwings}" == "1" {
		* Use xlwings for interactive manipulation of Excel
		
		
		
	
	}
	else {
		* Wrapper for putexcel if Python is not installed
		
		
		disp "Not using Python"
		
		if "`anything'" == "newbook" {
		
			global xlsheet_filename `"`using'"'
			cap putexcel clear
			cap rm "${xlsheet_filename}"
			putexcel set "${xlsheet_filename}", replace open
			
			
		}
		if "`anything'" == "newsheet" {
			cap putexcel save
			putexcel set "${xlsheet_filename}", open modify sheet(`name')
			global xlsheet_row 1
			putexcel A1 = `"`title'"', bold
			
			global xlsheet_row = ${xlsheet_row} + 1
			
		}
		if "`anything'" == "text" {
			putexcel A${xlsheet_row} = `"`title'"'
			
			global xlsheet_row = ${xlsheet_row} + 1
		}
		if "`anything'" == "graph" | "`anything'" == "chart" {
			
			* Get the part after the comma, and save it into the dataset
			gettoken left right : 0, parse(",") 
			gettoken left right : right, parse(",")
			
			char _dta[chartcmd] `right'
		}
		if "`anything'" == "importdata" {
			
			* Get the part after the comma, and save it into the dataset
			gettoken left right : 0, parse(",") 
			gettoken left right : right, parse(",")
			char _dta[tablecmd] `right'
			
			global xlsheet_row = ${xlsheet_row} + 1
			
			putexcel A${xlsheet_row} = `"`title'"', bold
			
			global xlsheet_row = ${xlsheet_row} + 1
			
			putexcel A${xlsheet_row} = `"`subtitle'"'
			
			global xlsheet_row = ${xlsheet_row} + 2
			
			
			local ncol = 1
			* Print var lists
			foreach var of varlist * {
				local col: word `ncol' of `c(ALPHA)'
				
				cap putexcel `col'${xlsheet_row} = "`:var label `var''", border(all) bold
				local row = ${xlsheet_row}
				
				forvalues i = 1(1)`=_N' {	
					local row = ${xlsheet_row} + `i'
					
					
					
					* If string, put down the value
					
					cap confirm string var `var'
					
					if !_rc {
						* it is a string:
						
						cap putexcel `col'`row' = `"`=`var'[`i']'"'
						
					}
					else {
						* it is a numerical value. First check if the value label exists.
						if "`:  val label `var''" != "" {
							cap putexcel `col'`row' = "`: label (`var')  `=`var'[`i']''"
						}
						else {
							cap putexcel `col'`row' = `=`var'[`i']', nformat(`format') 
						}
					
					}
					
					

					
					
					local ++row
				}

				local ++ncol
			}
			
			global xlsheet_row = ${xlsheet_row} + _N + 3
			
			
			
		}
		if "`anything'" == "close" {
			putexcel save
		}
		
		
		
		
	
	}
	
	
end
