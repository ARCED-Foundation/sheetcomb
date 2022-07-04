*! version 1.0.0 Mehrab Ali 07jul2022

prog sheetcomb 

	syntax using/, [force]
	
	clear 
	tempfile allsheet
	save `allsheet', emptyok		
		
		qui import excel "`using'", describe
		
		qui forval x=1/`=r(N_worksheet)' {
		    loc sh_`x' = "`=r(worksheet_`x')'"
		}
		
		
		forval x=1/`=r(N_worksheet)' {
		    import excel "`using'", sh("`sh_`x''") clear first 
			append using `allsheet', `force'
			save `allsheet', replace
		}
		
		u `allsheet'

end

