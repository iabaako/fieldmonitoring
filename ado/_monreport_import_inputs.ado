*! version 1.0 Ishmail Azindoo Baako, 11oct2018

* imports and prepares field montoring input file input file
program define _monreport_import_inputs
	syntax using/

	qui {
		import 		excel using "`using'", clear firstrow
		destring 	id role_id team_id mon_pin exp_weekly_subms_by exp_weekly_subms_for, replace

		* encode variables can be monitored and can monitor
		replace 	can_monitor 		= proper(can_monitor)
		replace 	can_be_monitored    = proper(can_be_monitored)
		label 		define _yesno 1 "Yes" 0 "No"

		encode 		can_monitor, gen(can_monitor_tmp) label(yesno)
		encode 		can_be_monitor, gen(can_be_monitor_tmp) label(yesno)
		drop 		can_monitor can_be_monitored
		rename 		*_tmp *
	}

end
 
