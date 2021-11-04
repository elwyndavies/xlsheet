* Example: creatgaps Year, by(Sector)
cap program drop xltab_creategaps
program define xltab_creategaps
	
	syntax varlist, [BYvar(varlist)]

	drop if missing(`varlist')
	sort `varlist' `byvar'
	order `varlist' `byvar'
	
	cap confirm string variable `varlist'
	
	if !_rc {
		bysort `varlist' (`byvar'): replace `varlist' = "" if _n != 1
	}
	else {
		bysort `varlist' (`byvar'): replace `varlist' = . if _n != 1
	}
	
	
end
