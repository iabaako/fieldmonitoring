*! Version 1.1.0 Ishmail Azindoo Baako Feb 18, 2018
* Version 1.0.0 Ishmail Azindoo Baako Jan 9, 2018

program define ipacheckmonitor
	#d;
	syntax using/, 
		infile(string) 
		outfile(string)
		xlsform(string)
		[commentdata(string)]
		;
	#d cr
	
	qui {
		
		/* ---------
		   TEMPFILES
		------------ */
		tempfile _master _data _corrdata _transit _comments
		
		* save data in memory
		save `_master', emptyok 
		
		/* ---------------------------------
		IMPORT ADDITIONAL INFO FROM XLS FORM
		------------------------------------ */	
		import excel using "`xlsform'", firstrow sheet(choices) clear
		keep if list_name == "equipment"
		keep list_name value label
		loc equip_count	`=_N' 
		levelsof value, loc (equips) clean
		
		* create local to hold labels for equipment
		loc equip_labels ""
		forval i = 1/`equip_count' {
			loc equip_value = value[`i']
			loc equip_name	= label[`i']
			
			loc equip_labels = `"`equip_labels' `equip_value' "`equip_name'""'
		}
		
		/* ----------------------------------
		   IMPORT AND PREPARE MONITORING DATA
		------------------------------------- */
		* use main dataset
		use "`using'", clear
		* gen week variable
		sort submissiondate
		cap gen subdate 	= dofc(submissiondate)
		cap gen startdate 	= dofc(starttime)
		cap gen enddate 	= dofc(enddate)
		format subdate startdate enddate %td
		gen date_week	= week(subdate)
		egen week 		= group(date_week)
		drop date_week
	
		* gen var to keep track of submissions
		gen 	count_ 	= 1
					
		save `_data'

		/* ------------------
		   SAVE COMMENTS DATA
		--------------------- */
		* check if data is imported in long or wide format
		cap conf longvar: ac_ind_less_r*
		if !_rc loc long "long"
		else loc wide "wide"
		if "`long' `wide'" ~= "" {
			* if data is in wide format, take out comments and reshape to long
			cap confirm var setofac_rpt
			if !_rc {
				count if !missing(setofac_rpt)
				loc comm_count `r(N)'
				if "`wide'" ~= "" & `comm_count' > 0 {
					keep key setofac_rpt ac_*
					destring ac_ind_r_*, replace
					#d;
					reshape long 	ac_ind_less_r_ 
									ac_add_comment_r_ 
									ac_to_improve_r_ 
									ac_to_improve_recom_r_ 
									ac_section_r_
								,	i(key) j(ac_ind_r_)
								;
					#d cr
					* drop unneeded rows
					drop if !ac_add_comment_r
					
					* generate parent key variable
					ren (*_) (*)
					replace key = key + "/ac_rpt[" +  string(ac_ind_r) + "]"
					
					save `_comments'
				}
			
				* if data is in long format, import and prep long dataset
				if "`long'" ~= "" & `comm_count' > 0 {
					use "`commentdata'", clear
					drop if !ac_add_comment_r
					drop key
					ren parent_key key
					
					save `_comments'
				}
				loc comm_var "setofac_rpt"
			}
			else loc comm_count 0
		}
		else loc comm_count 0
	
		/* --------------------------------
		   GENERATE AVERAGES FOR INDICATORS
		----------------------------------- */
		* Check for the number of times monitors used NO or DUL
			* Drop if Monitor Used NO/DUL more than 30% of the time
			* Exclude string variable or yesno variables from list to check
		use `_data', clear
		
		unab vars	: c_* ce_* p_* t_* r_* w_*
		unab exclude: p_equipment p_missing_equipment *_general
		loc evalvars: list vars - exclude
			
		* replace inperson = 1 if missing
		replace inperson = 1 if missing(inperson)
				
		* Generate Variable to hold number of non missing variable per observation
		egen nonmiss_count 	= rownonmiss(`evalvars')
			
		* Generate variable to hold number of NO and DULs
		egen nodul_count		= anycount(`evalvars'), values(-111 -222)
		gen nodul_percentage	= round((nodul_count/nonmiss_count) * 100)
		gen valid_submissions	= nodul_percentage < 30
					
		* recode NO and DUL to missing vals before calculating averages
			* recode -111 to .n
			* recode -222 to .l

		recode `evalvars' (-111 = .n) (-222 = .l)

		* generate average scores for each indicator
		
			* comunication
			unab c_vars: c_*
			loc c_use: list evalvars & c_vars
			egen communication 	= rowmean(`c_use')
			label var	 		communication 	"Communication"	
			* Compliance and Effectiveness
			unab ce_vars: ce_*
			loc ce_use: list evalvars & ce_vars
			egen compliance 	= rowmean(`ce_use')
			label var			compliance		"Compliance and Effectiveness"
			* Professionalism
			unab p_vars: p_*
			loc p_use: list evalvars & p_vars
			egen professionalism = rowmean(`p_use')
			label var			professionalism "professionalism"
			* Team Work
			unab t_vars: t_*
			loc t_use: list evalvars & t_vars
			egen teamwork 		= rowmean(`t_use')
			label var	 		teamwork 		"Team Work"
			* Reliability
			unab r_vars: r_*
			loc r_use: list evalvars & r_vars
			egen reliability	= rowmean(`r_use')
			label var	 		reliability 	"Reliability (Independence)"
			* Writing Skills
			unab w_vars: w_*
			loc w_use: list evalvars & w_vars
			egen writing		= rowmean(`w_use')
			label var	 		writing 		"Writing Skills"

		* keep only relevant vars
			#d;
			keep	mon_id - absent_comments 
					check_type
					survey_type
					inperson
					communication 
					compliance 
					professionalism 
					teamwork 
					reliability 
					writing 
					subdate
					startdate
					week
					count_
					key
					recommendation*
					`comm_var'
					p_equipment 
					p_missing_equipment
					valid_submissions
					nodul_percentage
					;
			#d cr
		
		* Relabel Survey Type
		label var survey_type Questionnaire
		
		/* ---------------------------
		   REPORT SUBMISSION SUMMARIES
		------------------------------ */
		* DAILY SUBMISSIONS
		* Export Headers
		cap putexcel close
		putexcel set "`outfile'", sheet("Summaries") replace			
		putexcel A1:D1 = "SUBMISSION BY DATE", 									 ///
			merge hcenter font(calibri, 12, red) bold border(bottom, double)		
		putexcel A2 = "Date" B2 = "Submissions" C2 = "Accompaniments" 			 ///
			D2 = "Spotchecks", hcenter font(calibri, 11) bold border(bottom)
		
		levelsof subdate, loc (dates) clean
		loc cell 3
		foreach date in `dates' {
			* format date
			loc date_fmt: disp %td `date'
			* Number of submissions for date
			count if subdate	 == `date'
			loc submissions 	`r(N)'	
			* Number of accompaniments for date 
			count if subdate 	== `date' & check_type == 1
			loc accompaniments 	`r(N)'
			* Number of Spotchecks for date
			count if subdate 	== `date' & check_type == 2
			loc spotchecks 		`r(N)'
			* Export Numbers to cell
			putexcel A`cell' = "`date_fmt'"		B`cell' = `submissions' 		 ///
					 C`cell' = `accompaniments' D`cell'	= `spotchecks', 		 ///
					 hcenter font(calibri, 11)	
			loc ++cell
		}
		
		* Export Totals
		* Total number of accompaniments
		count if check_type == 1
		loc accompaniments 	`r(N)'
	
		* Total number of spotchecks
		count if check_type == 2
		loc spotchecks		`r(N)'
			
		* Export Totals
		putexcel A`cell' = "TOTAL" 			B`cell' = `=_N'                          /// 
				 C`cell' = `accompaniments' D`cell' = `spotchecks', 				 ///
				 hcenter font(calibri, 11) bold border(top)
	
		* WEEKLY SUBMISSIONS
		* Export Headers		
		putexcel F1:I1 = "SUBMISSION BY WEEK", ///
			merge hcenter font(calibri, 12, red) bold border(bottom, double)
		putexcel F2 = "week" G2 = "submissions" H2 = "first date" I2 = "last date",  ///
			hcenter font(calibri, 11) bold border(bottom)
				
		levelsof week, loc (weeks) clean
		loc cell 3		
		foreach week in `weeks' {
			* Get first and last dates for each week
			su subdate if week == `week'
			loc first_date: disp %td `r(min)'
			loc last_date: disp %td `r(max)'
			
			* Total number of submissions for week
			count if week 	== `week'
			loc submissions `r(N)'
					
			* Export Weekly Numbers
			putexcel F`cell' = `week'			G`cell' = `submissions'  		 ///
					 H`cell' = "`first_date'" 	I`cell' = "`last_date'",		 ///
					 hcenter font(calibri, 11)
			
			loc ++cell
		}
	
		* Export Totals
		putexcel F`cell' = "TOTAL" 		G`cell' = `=_N', ///
			hcenter font(calibri, 11) bold border(top)
			
		* putexcel close

		* save data
		save `_data', replace
	
		/* --------------------------------------
		   REPORT SUBMISSION SUMMARIES BY MONITOR
		----------------------------------------- */
		* Export Headers
		putexcel set "`outfile'", sheet("Monitors") modify
			
		* Save weeks in local and count the number of weeks
		levelsof week, loc (weeks)
		loc week_count = wordcount("`weeks'")
		
		* Get end column letter.
		loc col = char(67 + `week_count')	
		putexcel A1:`col'1 = "WEEKLY SUBMISSION PER MONITOR", 		 				 ///
			merge hcenter font(calibri, 12, red) bold border(bottom, double)
		
		* collapse data by monitor id and week
		collapse (last) mon_name mon_role (sum) count_, by(mon_id week)
			
		* reshape data to wide format by mon_id
		reshape wide count_, i(mon_id) j(week)
			
		* clean week vars
		ren (count_*) (week_*)
		foreach var of varlist week_* {
			replace `var' = 0 if missing(`var')
			loc label = proper(subinstr("`var'", "_", " ", 1))
			lab var `var' "`label'"
		}
		
		* relabel other variables
			lab var mon_id 		"ID"
			lab var mon_name	"Name"
			lab var mon_role	"Position"
			
		* order variables
		order mon_id mon_name mon_role week_*, sequential
	
		* export data
		export excel using "`outfile'", sheet("Monitors") sheetmodify 		 	 ///
			cell(A2) first(varlab)
		
		* export headers
		putexcel A2:`col'2, bold border(bottom)
	
		* Output Submissions and averages for each field staff	
		* re-import data
		use `_data', clear
			
		* collapse data by enumerator_id, enumerator_role and week
			* This is to capture fieldstaff with multiple roles in one survey round
		
		* save data
		save `_transit', replace
		
		collapse 	(last) 	enumerator_name 									 	 ///
					(sum)	count_,											 	 	 ///
					by (enumerator_id enumerator_role week)

		/* -----------------------------------------
		   OUTPUT AVERAGES FOR EVALUATION INDICATORS
		-------------------------------------------- */	
		
		use `_transit', clear
		collapse 	(last) 	enumerator_name 									 	 ///
					(sum)	count_ if valid_submissions,							 ///
					by (enumerator_id enumerator_role week)

		keep enumerator_id enumerator_name enumerator_role count_ week
		reshape wide count_, i(enumerator_id enumerator_role) j(week) 
			
		isid enumerator_id enumerator_role
		egen submissions = rowtotal(count_*)
			
		* relabel other variables
			lab var enumerator_id 		"ID"
			lab var enumerator_name		"Name"
			lab var enumerator_role		"Position"
			lab var submissions			"All"
			
		* clean week variables
		ren (count_*) (week_*)
		foreach var of varlist week_* {
			replace `var' = 0 if missing(`var')
			loc label = proper(subinstr("`var'", "_", " ", 1))
			lab var `var' "`label'"
		}
				
		* order variables
		order enumerator_id enumerator_name enumerator_role submissions week_*,  	 ///
			sequential
			
		* Output Header
		putexcel set "`outfile'", sheet("Field Staff") modify
			
		loc col = char(68 + `week_count')
		putexcel A1:`col'1 = "WEEKLY SUBMISSION PER ENUMERATOR", 				 	 ///
			merge hcenter font(calibri, 12, red) bold border(bottom, double)			
		putexcel D2:`col'2 = "Submissions", 				 					 	 ///
			merge hcenter font(calibri, 12) bold border(bottom)
			
		* export data
		sort enumerator_role enumerator_name
		export excel using "`outfile'", sheet("Field Staff") sheetmodify 	 	 ///
			cell(A3) first(varlab)
			
		putexcel A3:`col'3, hcenter txtwrap bold border(bottom)
			
		* Output average scores for indicator averages
		* re-import data
		use `_data', clear
			
		* collapse data within enumerator_id and position
		collapse 	(last) 	enumerator_name 									 	 ///
					(sum)	count_												 	 ///
					(mean)	communication - writing,						 	 	 ///
					by (enumerator_id enumerator_role)
			
		* order vars
		order enumerator_id enumerator_name enumerator_role count_ communication 	 ///
			compliance professionalism teamwork reliability writing
		
		* relable variables
			lab var enumerator_id 	"ID"
			lab var enumerator_name "Name"
			lab var enumerator_role "Position"
			lab var count_ 			"Submissions"
			lab var communication 	"Communication"
			lab var compliance 		"Compliance and Effectiveness on job"
			lab var professionalism "Professionalism"
			lab var teamwork 		"Team Work"
			lab var reliability 	"Reliability (Independence)"
			lab var writing			"Writing"
		
		
		putexcel set "`outfile'", sheet("Evaluations") modify
		putexcel A1:J1 = "EVALUATION DETAILS", 				 						 ///
					merge hcenter font(calibri, 12, red) bold border(bottom, double)
		putexcel E2:J2 = "Indicators", merge hcenter bold border(bottom)
		
		export excel using "`outfile'", sheet("Evaluations") sheetmodify         ///
			cell(A3) first(varlab)
		
		putexcel A3:J3, hcenter bold border(bottom)
		loc rowcount = 4 + `=_N'
		putexcel E4:J`rowcount', nformat(number_d2)
		
		/* -----------------------
		OUTPUT INVALID SUBMISSIONS
		-------------------------- */
		use `_data', clear
		
		* label variables
			lab var subdate					"Submission Date"
			lab var mon_id 					"Monitor ID"
			lab var mon_name				"Monitor Name"
			lab var mon_role				"Monitor Position"
			lab var enumerator_id 			"Field Staff ID"
			lab var enumerator_name 		"Field Staff Name"
			lab var enumerator_role 		"Field Staff Position"
			lab var recommendation_comment	"Retrain Comment"
			lab var present					"With Permission"
			lab var nodul_percentage		"% NO/DUL"
			lab var check_type				"Monitoring Type"
			lab var inperson				"In Person"
			
		gsort -subdate -present enumerator_id mon_id
		count if nodul_percentage > 30 & !missing(nodul_percentage)
		loc nodul_count `r(N)'
		if `r(N)' {
			noi disp "Exporting `r(N)' submissions excluded from evaluation score"
			* Output Header
			putexcel set "`outfile'", sheet("Invalid Submissions") modify
			putexcel A1:M1 = "SUBMISSIONS NOT CONSIDERED IN EVALUATIONS",	 ///
				merge hcenter font(calibri, 12, red) bold border(bottom, double)

			export excel subdate											 ///	
				enumerator_id enumerator_name enumerator_role			 	 ///
				mon_id mon_name mon_role					 			 	 ///
				inperson check_type survey_type							 	 ///
				nodul_percentage										     ///
				using "`outfile'" if nodul_percentage > 30	&		 		 ///
				!missing(nodul_percentage),	 								 ///
				sheet("Invalid Submissions") sheetmodify cell(A2) first(varlab)
				
			putexcel A2:K2, hcenter bold border(bottom)
			loc row = `nodul_count' + 2
			putexcel A3:A`row'	, nformat(date_d_mon_yy)
		}
		
		/* --------------------
		   OUTPUT ABSENTEE LIST
		----------------------- */	
				
		* Check if there were absentees
		cap assert present != 2 & present != 3
		if !_rc {
			noi disp "No Absentee list to export"
		}
		else { 
			count if present == 2 | present == 3
			loc absent_count `r(N)'
			noi disp "Exporting `r(N)' observation to sheet Absent"
			
			* label absence as yes or no
			lab define present 2 "Yes" 3 "No", modify
			lab val present present
			
			* Output Header
			putexcel set "`outfile'", sheet("Absent") modify
			putexcel A1:M1 = "ABSENT DURING FIELD MONITORING",			    		 ///
				merge hcenter font(calibri, 12, red) bold border(bottom, double)
				
			gsort -subdate -present enumerator_id mon_id
			noi disp
			export excel subdate												 	 ///	
						 enumerator_id enumerator_name enumerator_role			 	 ///
						 mon_id mon_name mon_role					 			 	 ///
						 inperson check_type survey_type							 ///
						 present absent_comments									 ///
						 using "`outfile'" if present == 2 | present == 3,		 ///
						 sheet("Absent") sheetmodify cell(A2) first(varlab)
			
			putexcel A2:M2, hcenter bold border(bottom)	
			loc row = `absent_count' + 2
			putexcel A3:A`row'	, nformat(date_d_mon_yy)
		}
		
		/* -------------------------------
		   OUTPUT RETRAIN AND REPLACE LIST
		---------------------------------- */	
			
		* check if there were recommended re-trainings
		cap assert recommendation == 1
		if !_rc {
			noi disp 
			noi disp "No Recommendations to Export"
		}
		
		* export list to retrain
		else {
			* export retraining list
			count if recommendation == 2
			if `r(N)' > 0 {
				noi disp
				noi disp in red "Exporting `r(N)' recommended retraining sheet retrain"
			
				* Output Header
				putexcel set "`outfile'", sheet("Retrain") modify
				putexcel A1:I1 = "LIST OF FIELD STAFF RECOMMENDED FOR RETRAINING", 	 ///
					merge hcenter font(calibri, 12, red) bold border(bottom, double)			
			
				
				gsort -subdate enumerator_id mon_id
				* export data
				export excel subdate												 ///	
							 enumerator_id enumerator_name enumerator_role 			 ///						 
							 mon_id mon_name mon_role 								 ///
							 recommendation_comment									 ///
							 using "`outfile'" if recommendation == 2, 			 ///
							 sheet("Retrain") sheetmodify cell(A2) first(varlab)
				
				putexcel A2:I2, hcenter bold border(bottom)
				* format date column
				count if recommendation == 2
				loc row = `r(N)' + 2
				putexcel A3:A`row', nformat(date_d_mon_yy)
			}
			
			* export replacement list
			count if recommendation == 3
			if `r(N)' > 0 {
				noi disp
				noi disp in red "Exporting `r(N)' recommended replacement to sheet replace"
			
				* Output Header
				putexcel set "`outfile'", sheet("Replace") modify
				putexcel A1:I1 = "LIST OF FIELD STAFF RECOMMENDED FOR REPLACEMENT",  ///
					merge hcenter font(calibri, 12, red) bold border(bottom, double)			
			
				
				gsort -subdate enumerator_id mon_id
				* export data
				export excel subdate												 ///	
							 enumerator_id enumerator_name enumerator_role			 ///
							 mon_id mon_name mon_role 								 ///
							 recommendation_comment									 ///
							 using "`outfile'" if recommendation == 3, 			 ///
							 sheet("Replace") sheetmodify cell(A2) first(varlab)
							 
				putexcel A2:I2, hcenter bold border(bottom)
				* format date column
				loc row = `=_N' + 2
				putexcel A3:A`row', nformat(date_d_mon_yy)
			}
			
		}
	
		* save data
		save `_data', replace
	
		/* ---------------
		   OUTPUT COMMENTS
		------------------ */
		if `comm_count' > 0 {
		
			* keep only observations with comments
			keep if !missing(setofac_rpt)
			
			if `=_N' == 0 {
				noi disp
				noi disp "No comments to Export"
			}
		
			else {
				* merge in comments
				drop if mi(setofac_rpt)
				merge 1:m key using `_comments', assert(match) nogen
				
				* Relabel some variables
				lab var inperson 				"Inperson"
				lab var check_type 				"Monitoring Type"
				lab var survey_type 			"Survey"
				lab var ac_to_improve_r 		"Urgently Needs Improvement"
				lab var ac_to_improve_recom_r 	"Recommendations"
				lab var ac_section_r			"Evaluation Indicator"
				
				putexcel set "`outfile'", sheet("Comments") modify
				putexcel A1:M1 = "COMMENTS AND RECOMMENDATIONS FROM MONITORS",  	 ///
					merge hcenter font(calibri, 12, red) bold border(bottom, double)
					
				gsort -subdate enumerator_id mon_id
				noi disp
				noi disp "Exporting `=_N' comments to sheet comments"
				export excel subdate												 ///	
							 enumerator_id enumerator_name enumerator_role			 ///
							 mon_id mon_name mon_role					 			 ///
							 inperson check_type survey_type						 ///
							 ac_section_r ac_to_improve_r ac_to_improve_recom_r		 ///
							 using "`outfile'", 				 					 ///
							 sheet("Comments") sheetmodify cell(A2) first(varlab)
			
				putexcel A2:M2, hcenter bold border(bottom)
				* format date
				loc row = `=_N' + 2
				putexcel A3:A`row', nformat(date_d_mon_yy)
			}
		}
		
		/* --------------------------
		   OUTPUT EQUIPMENT SUMMARIES
		----------------------------- */
		* re-import data
		use `_data', clear
					
		* Create dummy variables to mark all variable
		replace p_missing_equipment = "_" + subinstr(p_missing_equipment, " ", "_", .) ///
			+ "_"
		
		* gen dummies for equipment
		foreach equip in `equips' {
			gen equip_`equip' 		= regexm(p_missing_equipment, "_`equip'_")
		}
		
		* check the number of missing equipment
		count if !p_equipment
		loc miss_equip_count 	`r(N)'
		
		* save data
		save `_data', replace

	
		* collapse data within enumerator_id and position
		collapse 	(mean)	equip_*,										 	 	 ///
					by (enumerator_role)
		
		* change to percentages
		foreach var of varlist equip_* {
			replace `var' = `var' * 100
		}
		
		* save positions in locals
		loc pos_count `=_N'
		forval i = 1/`pos_count' {
			loc position_`i' = enumerator_role[`i']
		}
		
		* transpose data
		xpose, varname format(%5.2f) clear
		drop in 1
		* Relabel variables
		forval i = 1/`pos_count' {
			lab var v`i' "`position_`i''"
		}
		lab var _varname "Equipment"
		
		* change equipment names
		label define equipment_list `equip_labels'
		foreach equip in `equips' {
			loc equip_lab		= "`:lab equipment_list `equip''"	
			replace _varname 	= "`equip_lab'" if _varname == "equip_`equip'"
		}
		
		* Export Headers
		putexcel set "`outfile'", sheet("Equipment") modify
		loc col = char(65 + `pos_count')
		putexcel A1:`col'1 = "PERCENTAGE OF MISSING EQUIPMENT BY POSITION",  		 ///
				merge hcenter txtwrap font(calibri, 12, red) bold border(bottom, double)
		order _varname v*, sequential
		export excel using "`outfile'", sheet("Equipment") sheetmodify cell(A2)  ///
			first(varlab)
		putexcel A2:`col'2, hcenter bold border(bottom)
		loc rowcount = 2 + `=_N'
		putexcel A3:`col'`rowcount', nformat(number_d2)
		
		* Export Field Staff with missing equipment
		if `miss_equip_count' >= 0 {
			* re-import data
			noi disp "Exporting `miss_equip_count' cases of missing equipment to sheet Equipment"
			use `_data', clear
			
			* keep only observations with missing equipment
			keep if !p_equipment
			
			* keep relevant vars
			#d;
			keep startdate
				 mon_id 
				 mon_name 
				 mon_role 
				 inperson 
				 check_type 
				 survey_type 
				 enumerator_id 
				 enumerator_name 
				 enumerator_role 
				 equip_*
				 ;
			#d cr
		}

	
		* reshape data to long format
		reshape long equip_, i(startdate enumerator_id mon_id) j(index)
		* drop cases where equipments were not missing
		drop if !equip_
		
		* generate equipment label
		gen equip_label = ""
		lab define	equipment_list `equip_labels'
		foreach equip in `equips' {
			loc equip_lab			= "`:lab equipment_list `equip''"	
			replace equip_label 	= "`equip_lab'" if index == `equip'
		}
		
		bys startdate mon_id enumerator_id: replace equip_label = string(_n) + ". " + equip_label
		bys startdate mon_id enumerator_id: gen count = _N
		drop index
		bys startdate mon_id enumerator_id: gen index = _n
		
		* Relabel some variables
			lab var inperson 				"Inperson"
			lab var check_type 				"Monitoring Type"
			lab var survey_type 			"Survey"
			lab var count					"# of Missing Equip"
			lab var equip_label				"Missing Equipment"

		* Define columns
		loc start 	= char(65 + `pos_count' + 2)
		loc end		= char(65 + `pos_count' + 2 + 6)
		loc row		= `=_N' + 2
		
		putexcel `start'1:`end'1 = "LIST OF MISSING EQUIPMENT",				  		 ///
				merge hcenter txtwrap font(calibri, 12, red) bold border(bottom, double)
		
		gsort -startdate enumerator_id mon_id index
		export excel startdate enumerator_id enumerator_name mon_id mon_name count	 ///
			equip_label using "`outfile'", sheet("Equipment") sheetmodify cell(`start'2)  ///
			first(varlab)
		
		putexcel `start'2:`end'2		, bold border(bottom)
		putexcel `start'3:`start'`row'	, nformat(date_d_mon_yy)
	
		* restore master dataset
		use `_master', clear
	}

end
