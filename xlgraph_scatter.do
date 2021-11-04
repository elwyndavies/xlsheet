cap program drop xlgraph_scatter
program define xlgraph_scatter
	syntax [varlist] [if] [in] [fweight  aweight  pweight  iweight], [by(str asis) order(str asis) title(str asis) mlabel(str asis)] [*]
	
	tokenize `varlist'
	
	local yvar `1'
	local xvar `2'
	
	// Create scatter plot
	preserve
		
		// Keep the observations to use
		marksample touse
		keep if `touse' == 1
		
		if `"`by'"' != "" {
			local agg `by'
		}
		else {
			tempvar agg 
			gen `agg' = 1
		}
		
		if `"`order'"' != "" {
			local timevar `order'
		}
		else {
			tempvar timevar 
			gen `timevar' = _n
		}
		
		if `"`order'"' == "`mlabel'" {
			tempvar timevar
			gen `timevar' = `order'
		}
		
		keep `mlabel' `agg' `timevar' `xvar' `yvar' 
		order `mlabel' `agg' `timevar' `xvar' `yvar' 
		
		local ytitle: var label `yvar'
		local xtitle: var label `xvar'
		
		xltab_reshape `xvar' `yvar' , i(`timevar') j(`agg')
		
		drop `timevar'
		
		order `mlabel'
	
		xlsheet importdata, title(`title')
		// subtype(straight_with_markers)
		xlsheet graph, type(scatter)  alternate ytitle(`ytitle') xtitle(`xtitle') legendposition(bottom) datalabels(custom) title(`title')

	restore
	
end