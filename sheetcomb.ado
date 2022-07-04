*! version 1.0.0 Mehrab Ali 07jul2022

cap prog drop sheetcomb
prog sheetcomb 

	syntax using/, [key(varlist) force  ///
			CELLRAnge(string)				///
			FIRSTrow					///
		///	ALLstring(string)				///
			ALLstring					///
			clear]
			
	if !mi("`cellrange'") loc cellrange = "cellrange(`cellrange')"
/*
	if !mi("`allstring'") loc `allstring' = "allstring(`allstring')"
*/
	
	clear 
	tempfile allsheet
	save `allsheet', emptyok		
		
		qui import excel "`using'", describe
		
		qui forval x=1/`=r(N_worksheet)' {
		    loc sh_`x' = "`=r(worksheet_`x')'"
		}
		
		
		forval x=1/`=r(N_worksheet)' {
		    import excel "`using'", sh("`sh_`x''") `clear' `firstrow' `cellrange' `allstring'
			
			if mi("`varlist'") {
				append using `allsheet', `force'
			}
			else if !mi("`varlist'") {
				merge m:m `varlist' using `allsheet', nogen `force' 
			}
			
			save `allsheet', replace
		}
		
		u `allsheet', clear

end

