!* monreport_labels Ishmail Azindoo Baako October 2018
* defines value labels

program define monreport_labels
	syntax[, get(namelist) all]

	qui {
		if regexm("`get'", "roles") | "`all'" ~= "" {

			#d;
			label define 
					roles
					1	"Surveyor"
					2	"Auditor (Field/Phone)"
					3	"Auditor (Audio)"
					4	"Editor"
					5	"Note Taker"
					6	"Facilitator"
					7	"Indepth Interviewer"
					8	"Observer"
					9	"Video Coder"
					10	"Transcriptionist"
					11	"Translator"
					12	"Data Entry Officer"
					13	"Team Leader"
					14	"Field Supervisor"
				;
			#d cr
		}

		if regexm("`get'", "yesno") | "`all'" ~= "" label define yesno	1 "Yes"	0 "No"

		if regexm("`get'", "languages") | "`all'" {
			#d;
			label define 
				languages
					1	"English"
					2	"Ahanta"
					3	"Akuapem Twi"
					4	"Asanti Twi"
					5	"Asen"
					6	"Banda"
					7	"Bassari"
					8	"Birifor"
					9	"Buli"
					10	"Chamba"
					11	"Chokosi"
					12	"Chumburung"
					13	"Dagaare"
					14	"Dagbani"
					15	"Dangme"
					16	"Efutu"
					17	"Ewe"
					18	"Fante"
					19	"Fulfulde"
					20	"Ga"
					21	"Gonja"
					22	"Gruni"
					23	"Guan"
					24	"Kamara"
					25	"Kassem"
					26	"Koma"
					27	"Likpakpan"
					28	"Mampruli"
					29	"Mo/Deg"
					30	"Moar"
					31	"Moshie"
					32	"Nabt"
					33	"Nafaanra"
					34	"Nankam"
					35	"Nawuri"
					36	"Ntrubo"
					37	"Nzema"
					38	"Sefwi"
					39	"Sissali"
					40	"Talen"
					41	"Tampulma"
					42	"Vagla"
					43	"Wali"
					44	"Wassa"
				;
			#d cr
		}		
	}

end
