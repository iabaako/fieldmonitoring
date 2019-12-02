# FIELD MONITORING OUTPUT

## Overview

IPA Ghana field montoring template is the result of an efforts to standardize field staff evaluations and synchronize field staff evaluations in Ghana with the new indicators from HQ. These programs analyze data from the IPA-Ghana Field Montoring Form and produces excel sheets showing the progress and field evaluations of Short Term Staff. 


## Installaion (Beta)

```stata
* monreport can be installed from github

net install monreport, all replace ///
	from("https://raw.githubusercontent.com/iabaako/fieldmonitoring/master/ado")
```

## Syntax
```stata
monreport using filename, outfile(string) xlsform(string) [commentdata(string) languagedata(string) wide long]

options
	outfile 	- field monitoring output
	xlsform         - SurveyCTO xls form for field monitoring
	enumdata        - enumerator data (.csv, .xlsx, .dta) format
	commentdata 	- comment data (.dta) format [required with long formatted data]
	languagedata	- language data (.dta) format [required with long formatted data]

```

## Example Syntax
```stata
* Long Formatted Dataset
monreport using "IPAGH Field Staff Monitoring Form.dta", ///
	outfile("monitoring_output.xlsx") ///
	xlsform("IPAGH Field Staff Monitoring Form.xlsx") ///
	enumdata("enumerator_details.csv") ///
	commentdata("IPAGH Field Staff Monitoring Form-ac_rpt.dta") ///
	languagedata("IPAGH Field Staff Monitoring Form-language_rpt.dta") ///
	long 

* Wide Formatted Dataset 
monreport using "IPAGH Field Staff Monitoring Form.dta", ///
	outfile("monitoring_output.xlsx") ///
	enumdata("enumerator_details.csv") ///
	xlsform("IPAGH Field Staff Monitoring Form.xlsx") ///
	wide
```

Please report all bugs/feature request to the [github issues page](https://github.com/PovertyAction/high-frequency-checks/issues)
