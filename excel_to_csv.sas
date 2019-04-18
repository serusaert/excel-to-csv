*********************************************************************************************;
** excel_to_csv.sas
** Kevin Beverly 2018-06-24
*********************************************************************************************;
** Convert spreadsheet workbook to CSV via Python script
**   1. Python program: excel_to_csv.py
**   2. Remove embedded newlines within cells
**   3. Consolidate whitespace
**   4. Convert text to utf-8 (non-translatable unicode characters replaced with "?")
**   5. Import to SAS via PROC IMPORT DBMS=CSV with GUESSINGROWS=max and GETNAMES
**   6. Check that row and column counts are non-zero and match between CSV and SAS dataset
** Example call: 
%**let spreadsheet = \\path\name-of-spreadsheet.xlsx;
%**let worksheet = Sheet1;
%**let csv = \\path\meds_coded.csv;
%**let dset = medcode;
%**excel_to_csv ( spreadsheet=&spreadsheet, worksheet=&worksheet, csv=&csv, dset=&dset );
*********************************************************************************************;

%macro excel_to_csv
  ( spreadsheet=, 
    worksheet=, 
    csv=, 
    dset=, 
    python=C:\Python\Python36\python.exe,  /* where is python? */
    pyprog=[full path to excel_to_csv.py], /* where is the python program paired with this macro? */
    verbose=0 );

  %**put **INF: &python "&pyprog" "&spreadsheet" "&worksheet" "&outfile" ;
  %** Example x command: ;
  **  x &python "&pyprog" "&spreadsheet" "&worksheet" "&csv" ;

filename myreturn pipe "&python ""&pyprog"" ""&spreadsheet"" ""&worksheet"" ""&csv"" ";
data _null_;
  infile myreturn truncover;
  input;
  length return_text call spreadsheet worksheet csv python pyprog $1000 dset $32 row_count col_count 8.;
  return_text = _infile_;
  csv="&csv";
  spreadsheet = "&spreadsheet";
  worksheet = "&worksheet";
  dset = "&dset";
  python = "&python";
  pyprog = "&pyprog";
  call = "&python ""&pyprog"" ""&spreadsheet"" ""&worksheet"" ""&csv"" ";
  if substr(_infile_,1,9) NE 'Finished:' then do;
    put "Er" "ror: Python script did not finish. Pipe command:";
    put "&python ""&pyprog"" ""&spreadsheet"" ""&worksheet"" ""&csv"" ";
    abort;
  end;
  row_count = input(scan(substr(_infile_,21),1,';'), best.) - 1;
  col_count = input(scan(substr(_infile_,20),2,'column_count='), best.);
  call symputx('row_count',row_count);
  call symputx('col_count',col_count);
  if &verbose then do;
    put "**INF: " spreadsheet =;
    put "**INF: " worksheet =;
    put "**INF: " csv =;
    put "**INF: " dset =;
    put "**INF: python EXE:    " python =;
    put "**INF: python script: " pyprog =;
    put "**INF: script: "        call =;
    put "**INF: command: "       return_text =;
    put "**INF: spreadsheet "    row_count =;
    put "**INF: spreadsheet "    col_count =;
  end;
run;
  %if %sysevalf(&row_count=0,bool) %then %do;
    %put Error: Zero rows present in exported CSV file!;
  %end;
  %if %sysevalf(&col_count=0,bool) %then %do;
    %put Error: Zero columns present in exported CSV file!;
  %end;
proc import datafile="&csv" out=&dset dbms=csv replace; GUESSINGROWS=max; GETNAMES=YES; run;
data _NULL_;
	if 0 then set &dset nobs=n;
	call symputx('nrows',n);
	stop;
run;
  %if %sysevalf(&verbose=1,bool) %then %do;
    %put **INF: OBS in &dset =&nrows;
  %end;
data _NULL_;
	set &dset(obs=1);
	dummy_num = 1;
	dummy_char = 'one';
  Array nums(*) _numeric_;
  array chrs(*) _character_;
  number_of_numerics = dim( nums ) - 1;
  if missing(number_of_numerics) then number_of_numerics = 0;
  number_of_charvars = dim( chrs ) - 1;	
  if missing(number_of_charvars) then number_of_charvars = 0;
	vnum = number_of_charvars + number_of_numerics;
	call symputx('number_of_numerics',number_of_numerics);
	call symputx('number_of_charvars',number_of_charvars);
	call symputx('vnum',vnum);	
run;
  %if %sysevalf(&verbose=1,bool) %then %do;
    %put **INF: number of numerics in &dset =&number_of_numerics;
    %put **INF: number of charvars in &dset =&number_of_charvars;
    %put **INF: vnum=&vnum;
  %end;
  %if %sysevalf(&verbose=1,bool) %then %do;
ods proclabel="&dset, &csv";
title1 "&dset, &csv";
proc contents data=&dset; run;
proc print data=&dset; run;
  %end;

* Test that row and column counts are same between CSV and imported SAS dataset;
  %if %sysevalf(&row_count=&nrows,bool) %then %do;
    %if %sysevalf(&nrows=0,bool) %then %do;
      %put Error: Zero rows were imported from spreadsheet!;
    %end;
    %else %if %sysevalf(&verbose=1,bool) %then %do;
      %put **INF: Row count (&row_count) agrees between spreadsheet and imported SAS dataset;
    %end;
  %end;
  %else %do;
    %put Error: CSV and SAS dataset row counts do not agree: csv=&row_count, dset=&nrows;
    %if %sysevalf(&nrows=0,bool) %then %do;
      %put Error: Zero rows were imported from CSV file!;
    %end;
  %end;
  %if %sysevalf(&col_count=&vnum,bool) %then %do;
    %if %sysevalf(&vnum=0,bool) %then %do;
      %put Error: Zero columns were imported from spreadsheet!;
    %end;
    %else %if %sysevalf(&verbose=1,bool) %then %do;
      %put **INF: Column count (&vnum) agrees between spreadsheet and imported SAS dataset;
    %end;
  %end;
  %else %do;
    %put Error: CSV and SAS dataset column counts do not agree: csv=&col_count, dset=&vnum;
    %if %sysevalf(&vnum=0,bool) %then %do;
      %put Error: Zero columns were imported from CSV file!;
    %end;
  %end;
%mend excel_to_csv;
