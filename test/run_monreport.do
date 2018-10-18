* Run test
clear all
cls
* install monreport
net install monreport, all replace from("C:\Users\Ishamial Boako\Box Sync\git\fieldmonitoring\ado")


/* tif midline

monreport using "X:\Dropbox\New DM Systems\TIF Midline HFC\05_data\04_monitoring/IPAGH Field Staff Monitoring Form - TIF Midline", ///
	outfile("C:\Users\Ishamial Boako\Box Sync\git\fieldmonitoring\outputs/monreport_tif_v2.xlsx") ///
	xlsform("X:\Dropbox\New DM Systems\TIF Midline HFC\01_instruments\03_xls/IPAGH Field Staff Monitoring Form_TIF_Midline.xlsx") ///
	commentdata("X:\Dropbox\New DM Systems\TIF Midline HFC\05_data\04_monitoring/IPAGH Field Staff Monitoring Form - TIF Midline-ac_rpt") ///
	languagedata("X:\Dropbox\New DM Systems\TIF Midline HFC\05_data\04_monitoring/IPAGH Field Staff Monitoring Form - TIF Midline-language_rpt") ///
	long 

*/
* lme endline

monreport using "X:\Dropbox\IPA-IFS shared folder LM project\03_Questionnaires&Data\11_LME Endline_HFC\05_data\04_monitoring/ipagh_field_staff_monitoring_form", ///
	outfile("C:\Users\Ishamial Boako\Box Sync\git\fieldmonitoring\outputs/monreport_lme.xlsx") ///
	xlsform("X:\Dropbox\IPA-IFS shared folder LM project\03_Questionnaires&Data\11_LME Endline_HFC\01_instruments\03_xls/IPAGH Field Staff Monitoring Form_v2") ///
	wide
