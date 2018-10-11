* Run test
clear all
include ../ado/_monreport_import_inputs.ado

tempfile inputs 
_monreport_import_inputs using "../inputs/monreport_input_template.xlsx"
