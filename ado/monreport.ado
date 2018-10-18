*! Version 1.0 (beta) Ishmail Azindoo Baako May, 2018
	* Stata program to analyse and report on field montoring data captured
	* using the IPA Field Monitoring Form
	
program define monreport
	#d;
	syntax using/, 
		outfile(string)
		xlsform(string)
		[commentdata(string) 
		languagedata(string)
		wide
		long]
		;
	#d cr
	
	/* Syntax description. 
		using - specifies main monitoring dataset (.dta format)
		outfile - specifies output file name and location (.xlsx format)
		xlsform - specifies surveycto xls form (.xlsx format) 
		commentdata - specifies comment repeat dataset long format datasets only. 
			If dataset is in wide format
		languagedata - specifies language repeat dataset for long format datasets 
			only. 
		wide - specifies if the data is in wide format
		long - specifies if the data is in long format
	*/
	
	qui {

		/*--------------
			CHECK SYNTAX
		---------------- */
		* check that option long or wide is specified
		if "`long'`wide'" == "" {
			nois disp as err "Must specify either long or wide option"
			ex 198
		} 
		* check that options long and wide are not specified together
		if "`long'" ~= "" & "`wide'" ~= "" {
			nois disp as err "options long and wide are mutually exclusive"
			ex 198
		}
		* check that options commentdata and languagedata are specified with option wide
		if "`long'" ~= "" & "`commentdata'" == "" {
			noi disp as err "option commentdata expected with long format"
			ex 198
		}
		if "`long'" ~= "" & "`languagedata'" == "" {
			noi disp as err "option languagedata expected with long format"
			ex 198
		}

		/* ---------
		   TEMPFILES
		------------ */
		tempfile _master _data _corrdata _transit _comments _language
		
		* save data in memory
		save `_master', emptyok 
		
		/* ---------------------------------
		IMPORT ADDITIONAL INFO FROM XLS FORM
		------------------------------------ */	
		
		* import the choices sheet and save names and labels of equipment 
		import excel using "`xlsform'", firstrow sheet(choices) clear
		keep if list_name == "equipment"
		keep list_name value label
		loc equip_count	`=_N' 
		levelsof value, loc (equips) clean
		
		* Create local to hold labels for equipment
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
		
		* drop observations from same monitor_id and enumerator_id which have
		* the same starttime, this will indicate a duplicate submission
		duplicates drop mon_id enumerator_id starttime, force
		
		* gen week variable. Weeks are from Sunday - Saturday
		sort submissiondate
		cap gen subdate 	= dofc(submissiondate)
		cap gen startdate 	= dofc(starttime)
		cap gen enddate 	= dofc(endtime)
		format subdate startdate enddate %td
		gen date_week	= week(subdate)
		egen week 		= group(date_week)
		drop date_week
	
		* gen var to keep track of submissions
		gen 	count_ 	= 1
		
		* update onsite variable
		destring onsite_mode, replace
		replace onsite = onsite_mode
			
		save `_data'

		/* ----------------------------------
		   IMPORT COMMENT AND LANGUAGE DATA
		------------------------------------- */

		* import comment and language data for long formatted data

		* check if data is imported in long or wide format
		count if ac_comment_yn == 1
		loc comm_count `r(N)'	
		if "`long'" ~= "" {
			cap assert !ac_comment_yn
			if _rc == 9 {
				use "`commentdata'", clear
				drop if !ac_add_comment_r
				drop key
				ren parent_key key
					
				save `_comments'
			}

			use `_data', clear
			destring c_languages_fs_rpt_count, replace

			cap assert !c_languages_fs_rpt_count
			if _rc == 9 {
				use "`languagedata'", clear
				drop key
				ren parent_key key
				save `_language'
			}
		}

		if "`wide'" ~= "" {
			if "`wide'" ~= "" & `comm_count' > 0 {
				keep key ac_*
				destring ac_ind_r_*, replace

				#d;
				reshape long 	ac_ind_r_
								ac_ind_less_r_ 
								ac_add_comment_r_ 
								ac_to_improve_r_ 
								ac_to_improve_recom_r_ 
								ac_section_r_,
								i(key) j(ac_ind_r)
							;
				#d cr
				* drop unneeded rows
				drop if missing(ac_section_r_)

				drop ac_ind_r_*
				
				* generate parent key variable
				ren (*_) (*)
				
				save `_comments'
			}	
			
			use `_data', clear
			
			count if !missing(c_languages_fs)
			if !_rc {
				count if !missing(c_languages_fs)
				loc lang_count `r(N)'
				if `lang_count' > 0 {
					keep key c_languages_fs_*
					destring c_languages_fs_ind_r, replace
					#d;
					reshape long 	c_languages_fs_ind_r 
									c_languages_fs_lab_r 
									c_languages_fs_prof_r
								,	i(key) j(c_languages_fs_ind_r)
								;
					#d cr
					
					* generate parent key variable
					ren (*_) (*)
					replace key = key + "/c_languages_fs_rpt[" +  string(ac_ind_r) + "]"
					
					save `_language'
				}
			}
			else loc lang_count 0	
		}
		else loc lang_count 0

		/* ----------------------
		   Export Project Details
		------------------------- */
		use `_data', clear
		loc project 	= project[1]
		loc acronym 	= project_acronym[1]
		loc phase 		= project_phase[1]
		loc capi_paper	= capi_or_paper[1] 
		loc software 	= capi_software[1] 
		loc area 		= research_area[1] 
		loc fm			= field_manager[1] 
		loc head		= project_head[1] 
		loc manager		= project_manager[1]
	
		putexcel set "`outfile'", sheet("Project Details") replace
		putexcel A1:B1 = "PROJECT DETAILS", merge hcenter font(calibri, 12) bold border(bottom, thick)
		
		putexcel A2 = "PROJECT NAME" 	A3 = "PROJECT ACRONYM"  ///
				 A4 = "PROJECT PHASE"	A5 = "CAPI/PAPER"		///
				 A6 = "CAPI SOFTWARE"	A7 = "RESEARCH AREA"	///
				 A8 = "FIELD MANAGER"	A9 = "PROJECT HEAD"		///
				 A10 = "PROJECT MANAGER" A11 = "DATE", bold
		
		putexcel B2 = "`project'" 	B3 = "`acronym'"  		///
				 B4 = "`phase'"		B5 = "`capi_paper'"		///
				 B6 = "`software'"	B7 = "`area'"			///
				 B8 = "`fm'"		B9 = "`head'"			///
				 B10 = "`manager'"	B11 = "`c(current_date)'"

		putexcel A11:B11, border(bottom, thick)
		putexcel A2:A11, border(left, thick)
		putexcel B2:B11, border(right, thick)

		* adjust column width
		* mata:	adjust_column("`outfile'", "Project Details", )
	

		/* ---------------------------
		REPORT ON SUBMISSION
		------------------------------ */
		* DAILY SUBMISSIONS
		* Export Headers
		putexcel set "`outfile'", sheet("Submissions") modify			
		putexcel A1:B1 = "SUBMISSION PER DATE", 									 ///
			merge hcenter font(calibri, 12) bold border(bottom, thick)		
		putexcel A2 = "Date" B2 = "Submissions", hcenter font(calibri, 11) bold border(bottom, thick)
		
		levelsof subdate, loc (dates) clean
		loc cell 3
		foreach date in `dates' {
			* format date
			loc date_fmt: disp %td `date'
			* Number of submissions for date
			count if subdate	 == `date'
			loc submissions 	`r(N)'	
			* Export Numbers to cell
			putexcel A`cell' = "`date_fmt'"		B`cell' = `submissions', hcenter font(calibri, 11)	
			loc ++cell
		}

		putexcel A`cell' = "Total" B`cell' = `=_N', bold hcenter border(top, thick)

		putexcel A`cell':B`cell', border(bottom, thick)
		putexcel A2:A`cell', border(left, thick)
		putexcel B2:B`cell', border(right, thick)

		* WEEKLY SUBMISSIONS
		* Export Headers		
		putexcel D1:G1 = "SUBMISSION PER WEEK", ///
			merge hcenter font(calibri, 12) bold border(bottom, thick)
		putexcel D2 = "week" E2 = "submissions" F2 = "first date" G2 = "last date",  ///
			hcenter font(calibri, 11) bold border(bottom, thick)
				
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
			putexcel D`cell' = `week'			E`cell' = `submissions'  		 ///
					 F`cell' = "`first_date'" 	G`cell' = "`last_date'",		 ///
					 hcenter font(calibri, 11)
			
			loc ++cell
		}
	
		* Export Totals
		putexcel D`cell' = "Total" 		E`cell' = `=_N', ///
			hcenter font(calibri, 11) bold

		putexcel D`cell':E`cell', border(bottom, thick) 
		putexcel D`cell':G`cell', border(top, thick) 
		putexcel E2:E`cell', border(right, thick) 
		putexcel D2:D`cell', border(left, thick)
		loc --cell
		putexcel G2:G`cell', border(right, thick) 
		
		* save data
		destring *id*, replace
		save `_data', replace

		/* --------------------------------------
		   REPORT SUBMISSION SUMMARIES PER MONITOR
		----------------------------------------- */
		* Export Headers
		putexcel set "`outfile'", sheet("Monitors") modify
			
		* Save weeks in local and count the number of weeks
		levelsof week, loc (weeks)
		loc week_count = wordcount("`weeks'")
		
		* Get end column letter.
		loc col = char(67 + `week_count')	
		putexcel A1:`col'1 = "WEEKLY SUBMISSION PER MONITOR", 		 				 ///
			merge hcenter font(calibri, 12) bold border(bottom, thick)
		
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
			relabel mon_id mon_name mon_role
			
		* sort vars and data
		gsort -mon_id
		order mon_id mon_name mon_role week_*, sequential
		
		* export data
		export excel using "`outfile'", sheet("Monitors") sheetmodify 		 	 ///
			cell(A2) first(varlab)
		
		* export headers
		putexcel A2:`col'2, bold border(bottom, thick)
		putexcel A2:A`=`=_N'+2', border(left, thick)
		putexcel `col'2:`col'`=`=_N'+2', border(right, thick)
		putexcel A`=`=_N'+2':`col'`=`=_N'+2', border(bottom, thick)
	
		* Output Submissions and averages for each field staff	
		* re-import data
		use `_data', clear

		/* -----------------------------------------
		   OUTPUT AVERAGES FOR EVALUATION INDICATORS
		-------------------------------------------- */	
		use `_data', clear
		
		* Exclude string variable or yesno variables from list to check
		unab vars	: c_* ce_* p_* t_* i_* w_*
		unab exclude: p_equipment p_missing_equipment *_general c_language_main c_language_label c_language_mon ///
			c_ul_mode c_languages_fs c_languages_fs_rpt_count
		loc evalvars: list vars - exclude
								
		* recode NO and DUL to missing vals before calculating averages
			* recode -111 to .n
			* recode -222 to .l
		
		recode `evalvars' (-111 = .n) (-222 = .l)
		
		* change label for interview_mode
		label define interview_mode_sel 3 "In-Person/On Phone", modify

		* generate a new var to represent interview interview_mode
		destring interview_mode, replace

		levelsof enumerator_id, loc (enumids) clean
		foreach id in `enumids' {
			tab interview_mode if enumerator_id == `id'
			if `r(r)' > 1 replace interview_mode_sel = 3 if enumerator_id == `id'
		}
		
		* collapse data by enumerator and position
		collapse 	(first) enumerator_name	interview_mode_sel			///
					(sum) submissions = count_							///
					(mean) `evalvars', by(enumerator_id enumerator_role)

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
			label var			professionalism "Professionalism"
			* Team Work
			unab t_vars: t_*
			loc t_use: list evalvars & t_vars
			egen teamwork 		= rowmean(`t_use')
			label var	 		teamwork 		"Team Work"
			* Independence
			unab i_vars: i_*
			loc i_use: list evalvars & i_vars
			egen independence	= rowmean(`i_use')
			label var	 		independence 	"Independence"
			* Writing Skills
			unab w_vars: w_*
			loc w_use: list evalvars & w_vars
			egen writing		= rowmean(`w_use')
			label var	 		writing 		"Writing Skills"
			
		foreach var of varlist communication - writing {
			replace `var' = int(round(`var'))
		}
		
		putexcel set "`outfile'", sheet("Staff Evaluations") modify
		putexcel E1:K1 = "STAFF EVALUATIONS", 				 						 ///
					merge hcenter font(calibri, 12) bold border(bottom, thick)
		putexcel E2:K2 = "Indicators", merge hcenter bold border(all, thick)
		
		relabel enumerator_* interview_mode_sel communication compliance professionalism teamwork ///
			submissions independence writing

		label values interview_mode_sel interview_mode_sel

		sort enumerator_role enumerator_id 
		export excel enumerator_id enumerator_name enumerator_role interview_mode_sel submissions communication - writing using "`outfile'", ///
			sheet("Staff Evaluations") sheetmodify         ///
			cell(A3) first(varlab)
		
		putexcel A3:K3, hcenter bold border(top, thick)
		putexcel A3:K3, hcenter bold border(bottom, thick)
		putexcel A3:A`=`=_N'+3', border(left, thick)
		putexcel K2:K`=`=_N'+3', border(right, thick)
		putexcel A`=`=_N'+3':K`=`=_N'+3', border(bottom, thick)

		/* -----------------------------------------
		OUTPUT LANGUAGE SCORES
		-------------------------------------------- */	
		use `_data', clear
	
		decode c_language_main, gen (language)
		ren c_language_main_prof proficiency
		keep enumerator_id enumerator_name enumerator_role count_ language ///
			c_languages_fs c_languages_fs_rpt_count proficiency
		
		destring c_languages_fs, replace

		if "`_language'" ~= "" & `lang_count' > 0 {
			keep if c_languages_fs == 1
			* merge with language data
			merge 1:m key using `_language', nogen keep(match master)
			drop language
			rename (c_languages_fs_lab_r) (language)
			save `_transit'	
			use `_data', clear
			ren c_language_main_prof proficiency
			decode c_language_main, gen (language)
			keep if c_languages_fs == 1
			append using `_transit'
		
		
		
			* drop languages with missing proficiency. This will happen if monitor does not understand language
			drop if missing(proficiency)
			* collapse data by enumerator and position
			collapse 	(first) enumerator_name								///
						(sum) submissions = count_							///
						(mean) proficiency, by(enumerator_id enumerator_role language)
			
			* round proficiency to the nearest whole number
			replace proficiency = round(proficiency)
			
			label val proficiency c_language_main_prof
			relabel enumerator_* 
			lab var proficiency "Proficiency (Oral)"
			lab var language	"Language" 
			lab var submissions "submissions"
				
			putexcel set "`outfile'", sheet("Language") modify
			putexcel A1:F1 = "LANGUAGE PROFICIENCY (ORAL)", 				 						 ///
						merge hcenter font(calibri, 12) bold border(bottom, thick)
			putexcel A2:F2, hcenter bold border(bottom, thick)
				
			sort enumerator_role enumerator_id 
			export excel enumerator_id enumerator_name enumerator_role submissions language proficiency using "`outfile'", ///
				sheet("Language") sheetmodify         ///
				cell(A2) first(varlab)

			putexcel A2:A`=`=_N'+2', border(left, thick)
			putexcel F2:F`=`=_N'+2', border(right, thick)
			putexcel A`=`=_N'+2':F`=`=_N'+2', border(bottom, thick)
		}
		else {
			keep if c_languages_fs == 1
		}


		/* -------------------------------
		   OUTPUT RETRAIN AND REPLACE LIST
		---------------------------------- */	
		use `_data', clear	
		* check if there were recommended re-trainings
		cap assert recommendation == 1
		if !_rc {
			noi disp "No Recommendations to Export"
		}
		
		* export list to retrain
		else {
			* export retraining list
			count if recommendation == 2
			if `r(N)' > 0 {
				noi disp "Exporting `r(N)' recommended retraining sheet retrain"
			
				* Output Header
				putexcel set "`outfile'", sheet("Retrain") modify
				putexcel A1:H1 = "LIST OF FIELD STAFF RECOMMENDED FOR RETRAINING", 	 ///
					merge hcenter font(calibri, 12) bold border(bottom, thick)			
			
				lab var recommendation_comment "comment"

				gsort -subdate enumerator_id mon_id

				relabel subdate  enumerator_id enumerator_name enumerator_role mon_id mon_name mon_role

				* export data
				export excel subdate												 ///	
							 enumerator_id enumerator_name enumerator_role 			 ///						 
							 mon_id mon_name mon_role 								 ///
							 recommendation_comment									 ///
							 using "`outfile'" if recommendation == 2, 			 ///
							 sheet("Retrain") sheetmodify cell(A2) first(varlab)
				
				putexcel A2:H2, hcenter bold border(bottom, thick)
				* format date column
				count if recommendation == 2
				loc row = `r(N)' + 2
				putexcel A3:A`row', nformat(date_d_mon_yy)

				putexcel A2:A`row', border(left, thick)
				putexcel H2:H`row', border(right, thick)
				putexcel A`row':H`row', border(bottom, thick)
			}
			
			* export replacement list
			count if recommendation == 3
			if `r(N)' > 0 {
				noi disp "Exporting `r(N)' recommended replacement to sheet replace"
			
				* Output Header
				putexcel set "`outfile'", sheet("Replace") modify
				putexcel A1:H1 = "LIST OF FIELD STAFF RECOMMENDED FOR REPLACEMENT",  ///
					merge hcenter font(calibri, 12) bold border(bottom, thick)			
				
				lab var recommendation_comment "comment"
				
				gsort -subdate enumerator_id mon_id

				relabel subdate  enumerator_id enumerator_name enumerator_role mon_id mon_name mon_role

				* export data
				export excel subdate												 ///	
							 enumerator_id enumerator_name enumerator_role			 ///
							 mon_id mon_name mon_role 								 ///
							 recommendation_comment									 ///
							 using "`outfile'" if recommendation == 3, 			 ///
							 sheet("Replace") sheetmodify cell(A2) first(varlab)
							 
				putexcel A2:H2, hcenter bold border(bottom)
				* format date column
				loc row = `=_N' + 2
				putexcel A3:A`row', nformat(date_d_mon_yy)

				putexcel A2:A`row', border(left, thick)
				putexcel H2:H`row', border(right, thick)
				putexcel A`row':H`row', border(bottom, thick)
			}
			
		}

		/* --------------------
		   OUTPUT ABSENTEE LIST
		----------------------- */	
	
		use `_data', clear		
		
		* re-label variables
		relabel mon_id mon_name mon_role onsite instrument
			lab var present 		"Absent Type"
			lab var absent_comments "Comment"
			lab var recommendation_comment "Comment"
			
		* Check if there were absentees
		cap assert present == 1 | missing(present)
		if !_rc {
			noi disp "No Absentee list to export"
		}
		else { 
			count if present < 1
			loc absent_count `r(N)'
			noi disp "Exporting `r(N)' observation to sheet Absent"
						
			* Output Header
			putexcel set "`outfile'", sheet("Absent") modify
			putexcel A1:K1 = "ABSENT DURING FIELD MONITORING",			    		 ///
				merge hcenter font(calibri, 12, red) bold border(bottom, double)
								
			gsort -subdate enumerator_id mon_id
			export excel subdate												 	 ///	
						 enumerator_id enumerator_name enumerator_role			 	 ///
						 mon_id mon_name mon_role					 			 	 ///
						 present onsite	instrument									 ///
						 absent_comments											 ///
						 using "`outfile'" if present < 1,		 ///
						 sheet("Absent") sheetmodify cell(A2) first(varlab)
			
			putexcel A2:K2, hcenter bold border(bottom)	
			loc row = `absent_count' + 2
			putexcel A3:A`row'	, nformat(date_d_mon_yy)
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
				merge hcenter txtwrap font(calibri, 12) bold border(bottom, thick)
		order _varname v*, sequential
		export excel using "`outfile'", sheet("Equipment") sheetmodify cell(A2)  ///
			first(varlab)
		putexcel A2:`col'2, hcenter bold border(bottom, thick)
		loc rowcount = 2 + `=_N'
		putexcel A3:`col'`rowcount', nformat(percent_d2)

		putexcel A2:A`rowcount', border(left, thick)
		putexcel `col'2:`col'`rowcount', border(right, thick)
		putexcel A`rowcount':`col'`rowcount', border(bottom, thick)
		
		* Export Field Staff with missing equipment
		if `miss_equip_count' > 0 {
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
				 onsite
				 instrument
				 enumerator_id 
				 enumerator_name 
				 enumerator_role 
				 equip_*
				 ;
			#d cr
	
			* reshape data to long format
			duplicates drop startdate enumerator_id mon_id, force
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
				relabel onsite instrument
				lab var count					"# of Missing Equip"
				lab var equip_label				"Missing Equipment"

			* Define columns
			loc start 	= char(65 + `pos_count' + 2)
			loc end		= char(65 + `pos_count' + 2 + 6)
			loc row		= `=_N' + 2
			
			putexcel `start'1:`end'1 = "LIST OF MISSING EQUIPMENT",				  		 ///
					merge hcenter txtwrap font(calibri, 12) bold border(bottom, thick)

			relabel startdate enumerator_id enumerator_name mon_id mon_name
			
			gsort -startdate enumerator_id mon_id index
			export excel startdate enumerator_id enumerator_name mon_id mon_name count	 ///
				equip_label using "`outfile'", sheet("Equipment") sheetmodify cell(`start'2)  ///
				first(varlab)
			
			putexcel `start'2:`end'2		, bold border(bottom, thick)
			putexcel `start'3:`start'`row'	, nformat(date_d_mon_yy)

			putexcel `start'2:`start'`row', border(left, thick)
			putexcel `end'2:`end'`row', border(right, thick)
			putexcel `start'`row':`end'`row', border(bottom, thick)
		
		}
		
		/* ---------------
		   OUTPUT COMMENTS
		------------------ */
		use `_data', clear
		if `comm_count' > 0 {
			* keep only observations with comments
			keep if ac_comment_yn == 1
			
			if `=_N' == 0 {
				noi disp
				noi disp "No comments to Export"
			}
		
			else {
				* merge in comments
				keep if ac_comment_yn == 1
				merge 1:m key using `_comments', keep (match) nogen
				
				* Relabel some variables
				relabel onsite instrument
				lab var ac_to_improve_r 		"Comment"
				lab var ac_section_r			"Indicator"
				
				putexcel set "`outfile'", sheet("Comments") modify
				putexcel A1:K1 = "COMMENTS AND RECOMMENDATIONS FROM MONITORS",  	 ///
					merge hcenter font(calibri, 12) bold border(bottom, thick)

				relabel enumerator_id enumerator_name enumerator_role mon_id mon_name mon_role
					
				gsort -subdate enumerator_id mon_id
				noi disp "Exporting `=_N' comments to sheet comments"
				export excel subdate												 ///	
							 enumerator_id enumerator_name enumerator_role			 ///
							 mon_id mon_name mon_role					 			 ///
							 onsite instrument										 ///
							 ac_section_r ac_to_improve_r 		 					 ///
							 using "`outfile'", 				 					 ///
							 sheet("Comments") sheetmodify cell(A2) first(varlab)
			
				putexcel A2:K2, hcenter bold border(bottom, thick)
				* format date
				loc row = `=_N' + 2
				putexcel A3:A`row', nformat(date_d_mon_yy)

				putexcel A2:A`row', border(left, thick)
				putexcel K2:K`row', border(right, thick)
				putexcel A`row':K`row', border(bottom, thick)
			}
		}
				
		* restore master dataset
		use `_master', clear
	}
		
end

* define program for relabelling variables
program define relabel
	syntax varlist
	
	loc	communication_lab 	"Communication"	
	loc	compliance_lab		"Compliance and Effectiveness"
	loc	professionalism_lab "Professionalism"
	loc	teamwork_lab 		"Team Work"
	loc	independence_lab 	"Independence"
	loc	writing_lab 		"Writing Skills"
	loc subdate_lab 		"Date" 
	loc enumerator_id_lab 	"Field Staff ID"
	loc enumerator_name_lab "Field Staff Name"
	loc enumerator_role_lab	"Field Staff Position"
	loc mon_id_lab 			"Monitor ID"
	loc mon_name_lab 		"Monitor Name"
	loc mon_role_lab		"Monitor Position"					 			 	
	loc onsite_lab			"Onsite/Remote"	
	loc instrument_lab		"Instrument"
	loc submissions_lab		"Submissions"
	loc interview_mode_sel_lab  "Interview Mode"
	
	foreach var of varlist `varlist' {
		lab var `var' "``var'_lab'"
	}
end
