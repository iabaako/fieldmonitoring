*! version 1.0 Ishmail Azindoo Baako, 11oct2018

* imports and prepares field montoring input file input file
program define _monreport_import_inputs
	syntax using/

	qui {
		import 		delim using "`using'", clear
		destring 	id role_id can_be_monitored can_monitor mon_pin weekly_target weekly_submissions, replace
	}

end
 
