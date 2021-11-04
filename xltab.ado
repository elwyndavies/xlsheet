// Version 0.1
// Elwyn Davies edavies@worldbank.org


cap program drop xltab
program define xltab
	
	syntax varlist [if] [in] [fweight  aweight  pweight  iweight], [row] [rowtotal] [PERCentage] [SHAREs] [missing] [EXClude(str asis)] [CUMULative] [stat(str)] [VARiable(str))] [nozero]  [graph(str asis)] [table(str asis)] [cmd(str asis)] [index(str asis)] [savecmd(str asis)] [savecmd2(str asis)] [savecmd3(str asis)] [save(str asis)] [*]
	
	
	if `"`graph'"' != "" | `"`table'"' != "" {
		preserve
	}
	
	
	
	
	if "`weight'" != "" {
		local weightopt [`weight'`exp']
	
	}
	
	* If no collapse stat is given, use "count"
	if "`stat'" == ""  local stat count
	if "`stat'" == "asis"  local stat first
	
	tokenize "`varlist'"
	
	* Count the number of variables
	local numvars : list sizeof varlist
	
	disp "`numvars'"
	
	
	
	* Drop missing values
	foreach var of local varlist {
		if "`missing'" == "" drop if missing(`var')
		if `"`exclude'"' != "" {
			foreach excludeval in `exclude' {
				cap drop if `var' == `excludeval'
				cap drop if `var' == "`excludeval'"
			}
		}
	}
	
	
	* Check number of observations
	
	count
	
	if r(N) == 0 {
		exit
	}
	
	
	* Tabulate one-way
	
	if `numvars'  == 1 {
		cap gen one = 1
		if "`variable'" == ""  local variable one
		
		collapse (`stat') count_=`variable' `if' `in' `weightopt', by(`1')
		
		* Check if you need to drop missing values
		
		
		label var count_ "Frequency"
		
		* Check if you need to report shares:
		if "`shares'" != "" {
			label var count_ "Percentage"
			
			egen Total = sum(count_)
			
			replace count_ = count_/Total
				
			if "`percentage'" != "" replace count_ = count_* 100
			
			cap drop Total
		}
	
	}
	
	
	* Tabulate two-way
	if `numvars'  == 2 | `numvars' == 3 {
		cap gen one = 1
		
		if "`variable'" == ""  local variable one
		
		collapse (`stat') count_=`variable' `if' `in' `weightopt', by(`varlist')
		
		* In case three variables are specified:
		if `numvars' == 3 {
			local 1  `1' `2'
			local 2  `3'
		
		}
		
		
		
		xltab_reshape count_, i(`1') j(`2')
		
		tokenize "`varlist'"
		if `numvars' == 3   xltab_creategaps `1', byvar(`2')
		
		* If the row option is specified, calculate shares
		if "`row'" != "" {
			egen Total = rowtotal(count_*)
			foreach var of varlist count_* {
				replace `var' = `var'/Total
				
				if "`percentage'" != "" replace `var' = `var' * 100
			}
			cap drop Total
		}
	}
	
	if "`rowtotal'" != "" {
		egen count_total = rowtotal(count_*)
		label var count_total "Total"
	}
	
	
	if "`cumulative'" != "" {
	
		* If cumulative numbers should be given:
		
		foreach var of varlist count_* {
			replace `var' = sum(`var')
		}
	
	}
	
	
	if "`zero'" != "nozero" {
	
		* Make sure there are zeroes when no values are present
		
		foreach var of varlist count_* {
			replace `var' = 0   if missing(`var')
		}
	
	}
	
	** Index numbers
	
	if "`index'" != "" {
		tempvar index_value
		cap gen `index_value' = count_`index'
		foreach var of varlist count_* {
			cap replace `var' = `var'/`index_value'
		
		}
		cap drop `index_value'
	
	}
	
	
	* Run any command if necessary
	`cmd'
	
	if `"`graph'"' != "" | `"`table'"' != "" {
		list
		
		xlsheet importdata, `table'
		if `"`graph'"' != ""  xlsheet graph, `graph'
		
		
	}
	
	* Run any command before saving
	
	if `"`save'"' != "" {
		`savecmd'
		`savecmd2'
		`savecmd3'
		
		notes _dta: table(`table') graph(`graph') 
		
		char _dta[tablecmd] table(`table')
		char _dta[graphcmd] graph(`graph')
		
		save `save', replace
	
	}
	
	if `"`graph'"' != "" | `"`table'"' != "" {
		restore
	}
	
	
	

end




