Ods excel merging cells after spanned column header and before column names

github
https://tinyurl.com/y74otqzn
https://github.com/rogerjdeangelis/utl_ods_excel_merging_cells_after_column_header_and_before_column_names


  Combination of SAS and  WPS/Proc R or IML/R
  I tried layout, grid ...

  SAS does not support merging cells after a header and before column names


see
https://tinyurl.com/y9xwe3mj
https://communities.sas.com/t5/ODS-and-Base-Reporting/Moved-spanned-columns-on-other-rows/m-p/456381



INPUT
=====

  WORK.CLASS total obs=19

      SEX    AGE    HEIGHT    WEIGHT

       M      14     69.0      112.5
       F      13     56.5       84.0
       F      13     65.3       98.0
       F      14     62.8      102.5
    ....

  EXAMPLE OF WHAT I WANT IN EXCEL

  d:/xls/class.xlsx

      <
      +---------------------------------------------------+
      |    A       |     B      |    C       |     D      |
      +---------------------------------------------------+
   1  |       DIMENSION         |       FACTS             |
      +---------------------------------------------------+
   2  |                       SCHEMA                      |  * This is the problem
      +---------------------------------------------------+  * cannot be easily done in SAS
   3  |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |
      +------------+------------+------------+------------+
   4  |    M       |    14      |    69      |   112      |
      +------------+------------+------------+------------+
   5  |    F       |    13      |    56      |    84      |
      +------------+------------+------------+------------+

[SHEET1]


PROCESS
=======


ods excel file="d:/xls/class.xlsx" style=statistical;
ods excel options(sheet_name="sheet1");
options missing="";
* note the addition of the header below the first header;
* we need to merge this header;
PROC REPORT DATA=sashelp.class missing headskip;
COLUMN  ( (  (("Dimension" (" "  AGE SEX ))) (("Factors" ("Schema"  WEIGHT HEIGHT )))   ) );
DEFINE  AGE    / display  ;
DEFINE  SEX    / display  ;
DEFINE  WEIGHT / display  ;
DEFINE  HEIGHT / display  ;
RUN;quit;
ods excel close;

* use IML/R or WPS/PROC R to fix the spanning row;
* you can use style option tp change fon color bold;
%utl_submit_wps64('
libname sd1 "d:/sd1";
options set=R_HOME "C:/Program Files/R/R-3.3.1";
proc r;
submit;
source("C:/Program Files/R/R-3.3.1/etc/Rprofile.site", echo=T);
library(XLConnect);
wb <- loadWorkbook("d:/xls/class.xlsx", create = FALSE);
unmergeCells(wb, sheet = "sheet1", reference = "A2:B2");
unmergeCells(wb, sheet = "sheet1", reference = "C2:D2");
mergeCells(wb, sheet = "sheet1", reference = "A2:D2");
writeWorksheet(wb,c("              SCHEMA"),sheet="sheet1",startRow=2,startCol=1,header=FALSE,rownames=FALSE);
saveWorkbook(wb);
endsubmit;
run;quit;
');

OUTPUT
====

  d:/xls/class.xlsx

      <
      +---------------------------------------------------+
      |    A       |     B      |    C       |     D      |
      +---------------------------------------------------+
   1  |       DIMENSION         |       FACTS             |
      +---------------------------------------------------+
   2  |                       SCHEMA                      |  * This is the problem
      +---------------------------------------------------+  * cannot be easily done in SAS
   3  |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |
      +------------+------------+------------+------------+
   4  |    M       |    14      |    69      |   112      |
      +------------+------------+------------+------------+
   5  |    F       |    13      |    56      |    84      |
      +------------+------------+------------+------------+


*                _              _       _
 _ __ ___   __ _| | _____    __| | __ _| |_ __ _
| '_ ` _ \ / _` | |/ / _ \  / _` |/ _` | __/ _` |
| | | | | | (_| |   <  __/ | (_| | (_| | || (_| |
|_| |_| |_|\__,_|_|\_\___|  \__,_|\__,_|\__\__,_|

;

  just use sashep.class

*          _       _   _
 ___  ___ | |_   _| |_(_) ___  _ __
/ __|/ _ \| | | | | __| |/ _ \| '_ \
\__ \ (_) | | |_| | |_| | (_) | | | |
|___/\___/|_|\__,_|\__|_|\___/|_| |_|

;

see above
