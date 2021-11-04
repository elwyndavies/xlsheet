cap program drop xlreg
program define xlreg
	syntax [anything], [new] [*]
	
	if "`new'" != "" {
		cap frame drop xlreg_combined
	}
	
	local regname `anything'
	matrix __`regname'_ = r(table)'
	svmat __`regname'_, name(matcol)
	
	local rownames : rowfullnames __`regname'_
	local c : word count `rownames'

	gen __`regname'_var = ""
	
	forvalues i = 1/`c' {
		replace __`regname'_var = "`:word `i' of `rownames''" in `i'
	}
	
	split __`regname'_var if __`regname'_var != "", parse(.)
	
	cap gen __`regname'_var2 = ""
	
	replace __`regname'_var2 = __`regname'_var1 if __`regname'_var2 == ""
	
	replace __`regname'_var1  = subinstr(__`regname'_var1,"bn","",.)
	
	destring __`regname'_var1, force replace
	
	
	
	*list __* if _n <10
	
	gen __`regname'_varlabel = ""
	gen __`regname'_vallabel = ""
	
	forvalues i = 1/`c' {
		cap replace __`regname'_varlabel = `" `: var label `=__`regname'_var2[`i']''"'  if _n == `i'
		
		cap replace __`regname'_varlabel = __`regname'_var2  if __`regname'_varlabel == ""
		cap replace __`regname'_vallabel = `" `: label (`=__`regname'_var2[`i']') `=__`regname'_var1[`i']''  "'  if _n == `i'
	}
	
	gen __`regname'_errorbar = abs(__`regname'_b-__`regname'_ul)
	
	
	cap frame drop xlreg_`regname'
	
	frame put __`regname'_varlabel __`regname'_vallabel __`regname'_b __`regname'_errorbar, into(xlreg_`regname')
	
	frame xlreg_`regname': drop if missing(__`regname'_b, __`regname'_errorbar)
	
	frame xlreg_`regname': rename __`regname'_varlabel varlabel
	frame xlreg_`regname': rename __`regname'_vallabel vallabel
	
	drop __`regname'_*
	
	
	tempfile newreg
	
	frame xlreg_`regname': save `newreg'
	
	* Also add this to xlreg_combined
	
	cap confirm frame xlreg_combined
	
	disp _rc
	if _rc {
		* The xlreg_combined frame does not exist yet
		frame copy xlreg_`regname' xlreg_combined
	}
	else {
		frame xlreg_combined: merge 1:1 varlabel vallabel using `newreg', gen(_regmerge)
		
		frame xlreg_combined: drop _regmerge
	}
	

end
