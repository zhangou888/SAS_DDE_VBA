/********************************************************************************
 *  Project: Using DDE to Populate values to formatted 
             EXCEL file for Standard Setting
    Programmer: Ou Zhang
    Initial code: 09-29-2018
    Goals: To populate sas data to a formated excel file (different tabs)

    Input: - Required: 1. sas data
                       2. formatted excel file
                      
    Output: Populated values to excel file
    Macros: 1. Transpose/ modify data by Round Module
            2. DDE Module     
            3. DDE to paste summary statistics to table (apply macro 1,2)
            4. Apply to multiple subjects and grade bands (apply macro 3) 

    Notes: STEP 1: Manually create/format a single excel table shell.
           STEP 2: Apply VBA macro to replicate multiple excel tabs using table shell from step 1.
           STEP 3: Apply SAS code and DDE to populate data to EXCEL file.

    !!! Special directory manipulation exists in paste_tble macro, this section has to be edited 
        before running the entire thing !!!!!
       
********************************************************************************/
/**  STEP 1: Define Work Environment  **/
proc datasets lib = work nolist kill; run;
ods listing;
dm 'log' clear;dm 'output' clear;
%let message = "WARNING:";
%let day = %sysfunc(date(),date7.);
options orientation=portrait ls=200 ps=89 pageno=1 formdlim=""
        FORMCHAR="|-,.,|+|'''+=|-/\<>*" nocenter;
options compress=binary ps=1000;
ods listing;
options SPOOL mprint mlogic;
ods listing gpath="%sysfunc(getoption(work))";

/* Delete all user-defined macro variables from the global symbol table */
%SYMDEL;
%let datetime = %sysfunc(compress(%sysfunc(today(),yymmddN7.)_%sysfunc(time(),hhmm6.), ':'));

/* STEP 2: custom format for variables */
proc format;
    value $labelfmt
                    'PLVL1' = 'B             '    /* Beginning      */
                    'PLVL2' = 'I             '    /* Intermediate  */
                    'PLVL3' = 'A             '    /* Advanced      */
                    'PLVL4' = 'AH            '    /* Advanced High */
                    ;        
run;

proc format ;
    value $colorfmt 
                'PLVL1 ' = 'Blue    '
                'PLVL2 ' = 'Red     '
                'PLVL3 ' = 'Yellow  '
                'PLVL4 ' = 'Green   '
                ;
run;

/** ----  Macro Section  ---- **/
/* Macro 1: Transpose/ modify data by Round Module */
%macro dat_trans(dat,out);
	
	/* dat - input data
       out - output data */
	%macro dummy();%mend dummy;

    /* dat - input data, 
       out - output data */

    data tmp1(keep= mean min q1 rd_median q3 max);
        set &dat.;
        rd_median = round(MEDIANR_,1); 
             mean = round(Mean_, .01);
                N = N_;
               SD = round(SD_,1);
              max = round(max_,1);
              min = round(min_,1);
               q1 = round(q1_,1);
               q3 = round(q3_,1);
    run;

    /* reorder variables */
    data tmp1;
        retain mean min q1 rd_median q3 max;
        set tmp1;
    run;

    /* transpose output */
    proc transpose data=tmp1 out=tmp2;run;

    /* drop _NAME_ variable, now you can paste to DDE */
    data &out.(drop=_name_);
        set tmp2;
        rename COL1 = int
               COL2 = ad1
               COL3 = ad2;
    run;

    proc datasets lib=work;
        delete tmp1 tmp2;
    run;quit;

%mend dat_trans;


/* Macro 2: DDE Module */
%macro dde_tbl(dat_dir, file_dir, xls_dir, filename, dat, sub, grade, vlist, r1, c1, r2, c2);

    /* dat_dir - sas data directory,  
       file_dir- Formatted excel file directory ,
       xls_dir - EXCEL APP location (EXCEL.EXE) , 
       filename- Formatted excel file name, 
       dat     - output SAS dataset (&sub.&grade.),
       vlist   - Output variable list (int ad1 ad2), 
       subj    - Subject/domain , 
       grade   - Grade 
       r1      - start row number
       c1      - start column number
       r2      - end row number
       c2      - end column number */
	
	%macro dummy();%mend dummy;

    /* Turn on EXCEL program and open excel doc */
    data _null_ ;
        x "'&xls_dir.'
           ""&file_dir.\&filename..xlsx""";

    /* DDE: define tab(&sub._&grade.) and range (r2c3:r19c5)*/
    filename out1 dde "excel|&sub._&grade.!r&r1.c&c1.:r&r2.c&c2.";

    /* Output data (check the variable list-&vlist, separated by " " */
    data &dat ;set &dat;
        file out1;
        put &vlist.; /*int ad1 ad2;*/
    run;

    /* Close out DDE */
    data _null_;
        file cmds;
        put '[close(0)]';
        put '[quit()]';
    run;

%mend dde_tbl;

/* Macro 3: DDE to paste summary statistics to table */
%macro paste_tbl(dat_dir, file_dir, xls_dir, filename, datname, sub, grade, vlist, r1, c1, r2, c2);

    /* dat_dir - sas data directory, 
       file_dir- Formatted excel file directory ,
       xls_dir - EXCEL APP location (EXCEL.EXE) , 
       filename- Formatted excel file name,  
       datname - part of sas data name (list68_sumstats_com_level_rd1.sas7bdat),      
       vlist   - Output variable list (int ad1 ad2), 
       sub     - Subject/domain , 
       grade   - Grade 
       r1      - start row number
       c1      - start column number
       r2      - end row number
       c2      - end column number */
	
	%macro dummy();%mend dummy;

    /* Key words for data directory */
    %if &sub = L %then %let sub1 = LIST;
    %if &sub = R %then %let sub1 = READ;
    %if &sub = S %then %let sub1 = SPEAK;

    /* ---- Special directory manipulation !!!! ---- */
    %let dat_dir1 = &dat_dir.\&sub1. &grade.\SAS;
    /* ---- Special directory manipulation !!!! ----*/

    /* set up library */
    libname x "&dat_dir1.";

    /* Read-in 3 Round data*/
    data round1;set x.&sub1.&grade._&datname._rd1;run;
    data round2;set x.&sub1.&grade._&datname._rd2;run;
    data round3;set x.&sub1.&grade._&datname._rd3;run;


    /* create data for round 1-round 3*/
    %dat_trans(round1,new1);
    %dat_trans(round2,new2);
    %dat_trans(round2,new3);

    /* combine 3 round data*/
    data &sub.&grade.;
        set new1 new2 new3;
    run;

    /* Apply DDE_tbl macro */
    %dde_tbl(&dat_dir1, &file_dir, &xls_dir, &filename, &sub.&grade., &sub., &grade., &vlist, &r1., &c1., &r2., &c2.);
   
    /* remove temp files */
    proc datasets lib=work;
        delete out1 round1 round2 round3 
               new1 new2 new3 &sub.&grade.;
    run;quit;

%mend paste_tbl;

/* Macro 4: Apply to multiple subjects and grade bands */
%macro multi_table(dat_dir, file_dir, xls_dir, filename, datname, subj, grade, vlist, r1, c1, r2, c2);

    /* dat_dir - sas data directory, 
       file_dir- Formatted excel file directory ,
       xls_dir - EXCEL APP location (EXCEL.EXE) , 
       filename- Formatted excel file name,  
       datname - part of sas data name (list68_sumstats_com_level_rd1.sas7bdat),      
       vlist   - Output variable list (int ad1 ad2), 
       sub     - Subject/domain , 
       grade   - Grade 
       r1      - start row number
       c1      - start column number
       r2      - end row number
       c2      - end column number */
	
	%macro dummy();%mend dummy;

    data _null_; 
        call symputx('nsub',count("&subj",'|')+1);
    run;

    /**subject loop;*/
    %do i = 1 %to &nsub; 
       %let s1 = %scan(&subj,&i,'|');
       %let Gra = %scan(&grade,&i,'|');

       data _null_; 
           call symputx('ngrade',count("&Gra",'/')+1);
       run;  
       %do j = 1 %to &ngrade;

            %let gr =  %scan(&Gra,&j,'/');
            %paste_tbl(&dat_dir, &file_dir, &xls_dir, &filename, &datname, &s1.,  &gr., &vlist., &r1., &c1., &r2., &c2.);
       %end; 
    %end;

%mend multi_table;

/** ----  Macro Section END ---- **/

/** -----------------      STEP 3: Examples    ------------------   **/
/** --- Example 1: Single subject + grade --- **/
%let sub      = L;
%let grade    = 23;
%let dat_dir  = Q:\PRS\ACCOUNTS\TX\TELPAS\2018\Standard Setting\ForRA;
%let datname  = sumstats_com_level;
%let xls_dir  = C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE;
%let file_dir = C:\Users\uzhanou\Documents\Standard Setting\TELPAS\tech report; 
%let filename = Recommended_Cut_Score_Summary_Statistics;
%let vlist    = int ad1 ad2;
/*(r2c3:r19c5)*/
%let r1       = 2 ;
%let c1       = 3 ;
%let r2       = 19;
%let c2       = 5 ;

/* Apply macros 3*/
%paste_tbl(&dat_dir, &file_dir, &xls_dir, &filename, &datname, &sub, &grade, &vlist, &r1., &c1., &r2., &c2. );


/** --- Example 2: Multi- subject + grade --- **/
/* set up subject and grade */
%let subj  = R|L|S;               /* R-reading, L-listening, S-Speaking*/
%let grade = 2/3/45/67/89/1012|   /* Grade,gradeband for each subject/domain separated by "|" */
             23/45/68/912|
             23/45/68/912;

%let dat_dir  = Q:\PRS\ACCOUNTS\TX\TELPAS\2018\Standard Setting\ForRA; 
%let datname  = sumstats_com_level;               
%let xls_dir  = C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE;
%let file_dir = C:\Users\uzhanou\Documents\Standard Setting\TELPAS\tech report; 
%let filename = Recommended_Cut_Score_Summary_Statistics;
%let vlist    = int ad1 ad2;

/*(r2c3:r19c5)*/
%let r1       = 2 ;
%let c1       = 3 ;
%let r2       = 19;
%let c2       = 5 ;

/* Apply the final macro */
%multi_table(&dat_dir, &file_dir, &xls_dir, &filename, &datname, &subj, &grade, &vlist, &r1, &c1, &r2, &c2);

/** ------------------  Easter Egg Section ----------------------  **/
/* EGG STEP 1: Turn on EXCEL */
options noxwait noxsync;
x '"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" ';

/* Sleep for 5 seconds to give Excel time to come up */
data _null_;
    x=sleep(5);
run;

/* EGG STEP 2: Use SAS to Run VBA macro on an Excel Macro-Enabled Workbook */
filename cmds dde 'excel|system';
data _null_;
    file cmds;

    /* Open the excel file test.xlsm which contains the VBA macro */
    put '[open("C:\Users\uzhanou\Documents\2018 Conference\Internal\enrichment\DDE+VBA\example\test.xlsm")]';

    /* Run copy macro in the test.xlsm to duplicate formatted tabs */
    put '[run("test.xlsm!copy")]';
run;
 
/** -----------   END of Easter Egg Section  --------------------- **/

/*****************************/
/****         EOF         ****/
/*****************************/



