*! version 1.0.0 Mehrab Ali 07jul2022

prog sheetcomb 

	syntax using/, [key(varlist)] [force]
	
	clear 
	tempfile allsheet
	save `allsheet', emptyok		
		
		qui import excel "`using'", describe
		
		qui forval x=1/`=r(N_worksheet)' {
		    loc sh_`x' = "`=r(worksheet_`x')'"
		}
		
		
		forval x=1/`=r(N_worksheet)' {
		    import excel "`using'", sh("`sh_`x''") clear first 
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

