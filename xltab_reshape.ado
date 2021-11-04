
cap program drop xltab_reshape
program define xltab_reshape

	syntax varlist, i(varlist) j(varlist)
	
	local var `j'
	
	* Check if it is string
	cap confirm string var `j'
	
	if !_rc {
		* if the variable is in string format
		
		rename `j' `j'_tmp
		
		* Use senconde to ensure that the order is kept the same
		sencode `j'_tmp, generate(`j')
		
		drop `j'_tmp
		
	}
	
	
	
	local tovar `varlist'
	levelsof `var', local(`var'_levels)       /* create local list of all values of `var' */
	foreach val of local `var'_levels {       /* loop over all values in local list `var'_levels */
      	 local val2 `val'
		 if `val' < 0   local val2 "_`=abs(`val')'"
		 disp "`val2'"
		 
		 
		 foreach tovar of varlist `varlist' {
			cap local lab`tovar'`val2' : label (`var') `val'    /* create macro that contains label for each value */
		 }
		 
		* disp "lab`tovar'`val' : label `var' `val'"
	}
	
	drop if `var' == .
	reshape wide `varlist', i(`i') j(`j')
	
	foreach var of varlist * {             
		 if "`lab`var''" != "" {
			 label variable `var' "`lab`var''"
		 }
		
	 }
end
