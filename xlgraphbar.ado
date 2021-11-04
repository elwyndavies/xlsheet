cap program drop xlgraphbar
program define xlgraphbar
	syntax [anything] [if] [in] [fweight  aweight  pweight  iweight], [Over(str asis)]  [display] [*]
	
	local orig `0'
	local rest `options'
	
	
	
	preserve

	
	
	* Loop through all potential over options (maximum is 9)
	
	forvalues i = 1(1)9 {
		* Check the current over option
		local 0 `over'
		syntax [anything], [reverse] [*]	
		local over`i' `anything'
		local over`i'_asyvars `asyvars'
		local over`i'_missing `missing'
		
		local byvar `byvar' `anything'
		
		* Go to the next one
		local 0 `", `rest' "'
		syntax [anything], [Over(str asis)] [*]
		local rest `options'
	}
	
	* Now, go back to the full command
	
	local 0 `orig'
	syntax [anything] [if] [in] [fweight  aweight  pweight  iweight], [MISSing] [noGAPs] [ASYvars] [ASCategories] [blabel(str)] [type(str)] [format(str)] [factor(int 1)] [percentages] [keepcollapselabels] [COLLAPSEcmd(str)] [noGAPs] [MULTIover(str)] [subtype(str)] [stack] [yvallabel(str)] [title(str)] [perclabel(str)] [countlabel(str)] [*]
	
	* Default option collapse command:
	if `"`collapsecmd'"' == ""  local collapsecmd collapse
	
	* Default option type command:
	if `"`type'"' == ""  local type column
	
	* Default label for the _perc var:
	if `"`perclabel'"' == ""  local perclabel Percentage
	if `"`countlabel'"' == ""  local countlabel Count
	
	* Deal with weights
	if "`weight'" != "" {
		local weightopt [`weight'`exp']
	}
	
	* Append the contents of multiover() to the variables indicated with over:
	local byvar `byvar' `multiover'
	
	if `"`title'"' == "" & "`byvar'" != "" {
		local tabletitle "`anything', over `byvar'"
	}
	else {
		local tabletitle `title'
	}
	
	* Save the variable labels
	foreach v of var * {
		local l`v' : variable label `v'
		if `"`l`v''"' == "" {
			local l`v' "`v'"
		}
	}
	
	
	* if there is no over var, create a variable named _one
	if "`byvar'" == "" {
		gen _one = 1
		local byvar _one
		label var _one "All"
	}
	
	* Check first option of anything
	gettoken firststat rest: anything
	
	if "`missing'" == "" {
		foreach var in `byvar' {
			quietly drop if missing(`var')
		}
	}
	
	* If this is (asis), then just keep the variables
	if "`firststat'" == "(asis)" {
		keep `byvar' `rest'
		order `byvar' `rest'
	}
	else if "`firststat'" == "(count)" {
		tempvar one
		gen `one' = 1
		`collapsecmd' (sum) _freq=`one' `if' `in' `weightopt', by(`byvar')
		label var _freq "`countlabel'"
	}
	else if "`firststat'" == "(percent)" | "`anything'" == "" {
		tempvar one
		cap keep `if' `in'
		
		gen `one' = 1/_N
		`collapsecmd' (sum) _perc=`one' `if' `in' `weightopt', by(`byvar')
		label var _perc "`perclabel'"
		
		if "`format'" == "" local format "0%"
	}
	else {
		`collapsecmd' `anything' `if' `in' `weightopt', by(`byvar')
	}
	
	* Restore the variable labels (they usually get lost in collapse)
	if "`keepcollapselabels'" == "" {
		foreach v of var * {
			if "`l`v''" != ""  label var `v' "`l`v''"
		}
	}
	

	
	* if ascategories is set, reshape...
	
	if "`ascategories'" != "" {
		unab allvars: *
		local vars : list allvars - byvar
		
		tempvar n 
		gen `n' = _n
		
		local i = 1
		* Save the current labels
		foreach var of local vars {
			local l`i' : variable label `var'
			if `"`l`i''"' == "" {
				local l`i' "`var'"
			}
		
			rename `var' _var`i'
			local ++i
		}
		
		* Create the _yval variable:
		reshape long _var, i(`n') j(_yval)
		
		cap drop `n'
		
		* Reassign the labels
		local i = 1
		cap label drop _yval
		foreach var of local vars {
			label define _yval `i' `"`l`i''"', add
			local ++i
		}
		
		label val _yval _yval
		
		if "`byvar'" == "_one" {
			drop _one
			local byvar ""
		}
		
		* _yval is now a new byvar, if it is not yet defined:
		
		local byvar _yval `byvar'
		
		*if `: list posof "_yval" in byvar' == 0 {
		*}
		
		
		
		label var _yval `"`yvallabel'"'
	}
	
	
	
	* Reverse the bylist
	local reversebyvar ""
	
	foreach var of varlist `byvar' {
		order `var'
		local reversebyvar `var' `reversebyvar'
	}
	
	sort `reversebyvar'
	
	* Count the number of category variables:
	local catcols : list sizeof byvar 
	
	
	
	
	
	* Apply the factor option: multiply all variables, except for the byvar, with the specified factor
	* This is useful when the data is in fractions and you want to conver them to percentages (e.g. 0.05 -> 5)
	* Note that Excel can also do this by applying formatting. The Excel format 0.0% will display 0.05 as 5.0%
	
	
	
	
	foreach var of local vars {
		cap replace `var' = `var' * `factor'
	}
	
	

	* If the asyvars option is selected, the dataset gets reshaped
	
	if "`asyvars'" != "" {
		unab allvars: *
		local vars : list allvars - byvar
		
		gettoken byvar1 byvarrest : byvar
		
		* Check the size of byvarrest, if it is zero, we need an additional variable for the reshape to work
		
		local byvarrest_N : list sizeof byvarrest 
		
		if `byvarrest_N' == 0 {
			tempvar onedummy
			gen `onedummy' = 1
			local byvarrest `onedummy'
			
			local ++catcols
		}
		
		xltab_reshape `vars', i(`byvarrest') j(`byvar1')
		
		/* if `byvarrest_N' == 0 {
			drop `onedummy'
		}*/
		
		* The number of categories is one lower
		local --catcols
		
		* The byvar variable is now without the reshaped variable
		local byvar `byvarrest'
	}
	
	
	unab allvars: *
	local vars : list allvars - byvar
	
	if "`percentages'" != "" {
		tempvar sum
		egen `sum' = rowtotal(`vars')
		
		foreach var of local vars {
			cap replace `var' = `var' / `sum'
		}
		
		*list
		
		drop `sum'
		
		if "`format'" == "" local format "0%"
		
	}
	
	
	
	
	if `"`blabel'"' == "bar" 	local labelopt "datalabels(value)"
	if `"`format'"' != "" 		local labelopt "datalabelformat(`format')"
	
	* Next, create gaps:
	
	local byvarrest_N : list sizeof byvarrest
	
	* If the number of byvar vars are larger than 1, we will need to create gaps
	if `catcols' > 1 {
		
		local i = 1
		foreach var of varlist * {
		
			* Only do this for the variables at category variables, excluding the final one
			if `i' < `catcols' {
				tempvar repeatedval
				disp "Removing empty values of `var'"
				
				gen `repeatedval' = 1	 if `var' == `var'[_n-1]
				
				* Replace this with an empty value ("" or ., depending on the variable type)
				cap noisily replace `var' = "" 	if `repeatedval' == 1
				cap noisily replace `var' = . 	if `repeatedval' == 1
				
				drop `repeatedval'
			}
			
			local ++i
			
		}
	}
	
	* If the stack option is specified, set the subtype to "stacked"
	if "`stack'" != "" & `"`subtype'"' != "" & (`"`type'"' == "column" | `"`type'"' == "bar" | `"`type'"' == "line" | `"`type'"' == "area") local subtype stacked
	
	if "$xlgraphbarcmd" == "1" {
		local tablesubtitle xlgraphbar `0'
	}
	
	xlsheet importdata, format(`format') title(`tabletitle') subtitle(`tablesubtitle')
	xlsheet graph, type(`type') subtype(`subtype') catcols(`catcols') `labelopt' `options' title(`title')
	
	
	* If the dataset needs to be displayed:
	
	if `"`display'"' != ""  list
	
	restore
	
	
	*collapse `anything', by(`byvar')
	
end
