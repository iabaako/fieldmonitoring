*! version 1.0 Ishmail Azindoo Baako, 11oct2018

* imports contract request form and outputs monitoring inputs file

program define _monreport_create_inputs
	syntax using/, saving(string)
	qui {
		tempfile _staff 
		
		* default contract request extension to .xlsx if extension is not specified
		if !regexm("`using'", "(\.xlsx|\.xls|\.xlsm)$") loc using "`using'.xlsx"
		
		* confirm that file exist, else err 601 
		confirm file "`using'"

		* import contract request form
		cap import excel using "`using'", sheet("`sheet'") cellrange(A14) ///
			firstrow case(lower) allstring clear
			keep type - contract_end
				
		egen nonmiss = rownonmiss(_all), strok
		drop if nonmiss <= 3 | type == "."
		
		* generate row number to mark observations
		gen row = _n + 14
		
		destring uniqueid, replace
		* check that ID var is numeric and is nonmissing
		cap confirm numeric var uniqueid
		if _rc == 7 {
			levelsof row if !regexm(uniqueid, "[0-9]"), loc (rows) clean
			di as err "uniqueid contains non numeric characters on row(s) `rows'."
			ex 7
		}
		cap assert !missing(uniqueid)
		if _rc == 9 {
			levelsof row if missing(uniqueid), loc (rows) clean
			di as err "uniqued_id contains missing values on row(s) `rows'"
			ex 9
		}
		cap isid uniqueid
		if _rc == 459 {
			duplicates tag uniqueid, gen (dups)
			levelsof row if dup, loc (rows) clean
			di as err "variable uniqueid has duplicate observations on rows `rows'"
			exit 459
		}

		replace fullname = proper(fullname)
		
		keep uniqueid fullname role 
		ren (uniqueid fullname role) ///
			(id name role)
		
		order id name role

		* import role id
		_monreport_labels, get(roles)

		replace role = proper(role)
		
		encode role, gen(role_id) label(roles) noext
		order role_id, after(role)
		_strip_labels role_id
		sort role_id id

		gen int can_be_monitored 	= 1
		gen int can_monitor 		= inlist(role_id, 6, 13, 14)
		gen int mon_pin             = runiformint(1000, 9999) if can_monitor

		gen int weekly_target				= .
		gen int weekly_submissions			= .

		* export file
		export delim using "`saving'", replace 
	}
end
