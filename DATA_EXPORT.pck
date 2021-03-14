create or replace package DATA_EXPORT is

  -- Author  : LEOH
  -- Created : 15/09/2015 4:14:54 PM
  -- Purpose : Export data into files.
  
  /**
   * UNI_COURSE_QUOTE, extract codeii quote and rank for currently year courses. 
   *
   * p_org: organizations C - Curtin, U - UWA, E - ECU, M - Murdoch
   * p_round: round 1 and 2
   *
   * File path: /work/tiscout on tiscprod (CSV format).
   * File is transferred to ARTS for EServices, at: /usr/local/datatfr/<uni>
   *
   * crontab job - /export/home/kobol/sql/unicoursequoteb in tiscprod
  **/
  procedure uni_course_quote (p_org IN varchar2, p_round IN INT,  p_admissionyear IN INT := 0);
  
  /**
   * Notre Dame TR301Q annual request for WACE course and ATAR data for students who have applied to them
   * File path: /work/tiscout on tiscprod (CSV format).
   * Require manually transferring file to ARTS for EServices, at: /usr/local/datatfr/ndu
  **/
  procedure ndu_appnum_results (p_processyear IN INT := 0);
  
  /**
   * Notre Dame current Year 12 student matching
   * File path: /work/tiscout on tiscprod (CSV format).
   * Require manually transferring file to ARTS for EServices, at: /usr/local/datatfr/ndu
  **/
  procedure ndu_appnum_checked (p_processyear IN INT := 0);
  
  /**
   * Export UU015 (UWA Planning report) into Excel format (.xlsx).
   * File path: /work/tiscout on tiscprod.
   * File is transferred to ARTS for EServices, at: /usr/local/datatfr/uwa
   *
   * crontab job - /export/home/kobol/sql/uu015b in tiscprod
  **/
  procedure uwa_uu015_export (p_processyear IN INT := 0);
  
    /**
   * Export UX002 (Curtin or UWA report) into csv format (.csv).
   * call tiscadmin.COURSELIST procedure
   *
   * crontab job - /export/home/kobol/sql/uc001tj in tiscprod
  **/
  procedure ux002_export (p_inst in varchar2);
  
  /**
   * Export Course code into Excel format (.xlsx).
   * File path: /work/tmp on tiscprod.
   * File is transferred to ARTS for EServices, at: /usr/local/datatfr/download
   *
   * crontab job - /export/home/kobol/on_line/ap109b in tiscprod
  **/
  procedure course_code_export (p_admissionyear IN INT := 0);
  
  /**
   * Export Arts Inistitution code into Excel format (.xlsx).
   * File path: /work/tmp on tiscprod.
   * File is transferred to ARTS for EServices, at: /usr/local/datatfr/download
   *
   * crontab job - /export/home/kobol/sql/tisccodesexportb in tiscprod
  **/
  procedure arts_institution_code_export (p_processyear IN INT := 0);
  
  
  /**
   * Export Country code into Excel format (.xlsx).
   * File path: /work/tmp on tiscprod.
   * File is transferred to ARTS for EServices, at: /usr/local/datatfr/download
   *
   * crontab job - /export/home/kobol/sql/tisccodesexportb in tiscprod
  **/
  procedure country_code_export (p_processyear IN INT := 0);
  
  
  /**
   * Export Lang code into Excel format (.xlsx).
   * File path: /work/tmp on tiscprod.
   * File is transferred to ARTS for EServices, at: /usr/local/datatfr/download
   *
   * crontab job - /export/home/kobol/sql/tisccodesexportb in tiscprod
  **/
  procedure lang_code_export (p_processyear IN INT := 0);
  
  
  /**
   * Export School code into Excel format (.xlsx).
   * File path: /work/tmp on tiscprod.
   * File is transferred to ARTS for EServices, at: /usr/local/datatfr/download
   *
   * crontab job - /export/home/kobol/sql/tisccodesexportb in tiscprod
  **/
  procedure school_code_export (p_processyear IN INT := 0);
  
  
  /**
   * Export Subject code into Excel format (.xlsx).
   * File path: /work/tmp on tiscprod.
   * File is transferred to ARTS for EServices, at: /usr/local/datatfr/download
   *
   * crontab job - /export/home/kobol/sql/tisccodesexportb in tiscprod
  **/
  procedure subj_code_export (p_processyear IN INT := 0);
  
    /**
   * Export Basis of admission code into Excel format (.xlsx).
   * File path: /work/tmp on tiscprod.
   * File is transferred to ARTS for EServices, at: /usr/local/datatfr/download
   *
   * crontab job - /export/home/kobol/sql/tisccodesexportb in tiscprod
  **/
  procedure basis_admission_code_export (p_processyear IN INT := 0);
  
  /**
   * Export ECU Preference Report into Excel format (.xlsx).
   * File path: /work/tmp on tiscprod.
   * File is transferred to ARTS for EServices, at: /usr/local/datatfr/ecu/prefs
   *
   * crontab job - /export/home/kobol/sql/ecuprefs1b in tiscprod
  **/
  procedure ecu_prefs_report;
  
  
  /**********************************************************************/
  
  /**
   * Note in use as generated file size is too big due to XML format.
   * Kept as a working example to utilise XML format.
   */
  procedure uwa_uu015_export_xml (p_processyear IN INT := 0);
  
end DATA_EXPORT;
/
create or replace package body DATA_EXPORT is
  v_processyear number;
  v_admissionyear number;
  r number := 0;
  
  procedure uni_course_quote (p_org IN varchar2, p_round IN INT,  p_admissionyear IN INT := 0) IS
    v_file utl_file.file_type;
    v_org VARCHAR2(10);
    
    BEGIN
        IF (p_admissionyear > 0) THEN
            v_admissionyear := p_admissionyear;  
        ELSE
            v_admissionyear := TISCCOMMON.getAdmissionYearUg;
        END IF;
        
        -- check organization
        IF (UPPER(p_org) = 'C') THEN
           v_org := 'cur';
        ELSIF (UPPER(p_org) = 'E') THEN
           v_org := 'ecu';
        ELSIF (UPPER(p_org) = 'M') THEN
           v_org := 'mur'; 
        ELSIF (UPPER(p_org) = 'U') THEN
           v_org := 'uwa';
        ELSE
          dbms_output.put_line('Please input correct organization.');
          RETURN;
        END IF;
          
        v_file :=utl_file.fopen('OUT_TISC','quota-listing-' || v_org || '-r' || to_char(p_round) ||'.csv','w');
        
        -- check round
        IF (p_round = 1) THEN
           utl_file.put_line(v_file,'INCO_A,INCO_NAME,SHORT_NAME,Q1_RANK,Q1_QUOTA,LEVEL_IND,STATUS,DEET_FOS,SEC_FOE,COURSE_TYPE,ADMISSIONYEAR');
        ELSIF (p_round = 2) THEN
           utl_file.put_line(v_file,'INCO_A,INCO_NAME,SHORT_NAME,Q1_RANK,Q1_QUOTA,Q2_RANK,Q2_QUOTA,LEVEL_IND,STATUS,DEET_FOS,SEC_FOE,COURSE_TYPE,ADMISSIONYEAR');
        ELSE
           dbms_output.put_line('Pleaser input correct round.');
           RETURN;
        END IF;
           
        FOR codeii_row IN (SELECT inco_a, inco_name, short_name, replace(q1_rank, '0000', null) as q1_rank, 
                                  replace(q1_quota, '0000', null) as q1_quota, replace(q2_rank, '0000', null) as q2_rank,
                                  replace(q2_quota, '0000', null) as q2_quota, level_ind, status, deet_fos,
                                  replace(sec_foe, '000000', null) as sec_foe, course_type, admissionyear 
                                         FROM codeii ci WHERE inco_a like UPPER(p_org) || '%' and admissionyear=v_admissionyear order by inco_a)            
        LOOP
              
            IF (p_round = 1) THEN
               utl_file.put_line(v_file, '"'|| codeii_row.inco_a ||'","'|| codeii_row.inco_name ||'","'|| codeii_row.short_name ||'","'|| codeii_row.q1_rank ||'","'|| codeii_row.q1_quota ||'","'|| 
                                     codeii_row.level_ind ||'","'|| codeii_row.status ||'","'|| codeii_row.deet_fos ||'","'|| codeii_row.sec_foe ||'","'|| codeii_row.course_type ||'","'|| codeii_row.admissionyear || '"');                           
            ELSIF (p_round = 2) THEN
               utl_file.put_line(v_file, '"'|| codeii_row.inco_a ||'","'|| codeii_row.inco_name ||'","'|| codeii_row.short_name ||'","'|| codeii_row.q1_rank ||'","'|| codeii_row.q1_quota ||'","'|| codeii_row.q2_rank ||'","'|| codeii_row.q2_quota 
                                     ||'","'|| codeii_row.level_ind ||'","'|| codeii_row.status ||'","'|| codeii_row.deet_fos ||'","'|| codeii_row.sec_foe ||'","'|| codeii_row.course_type ||'","'|| codeii_row.admissionyear || '"');
            END IF;   
        END LOOP;
        
        UTL_FILE.FCLOSE(v_file);
    END;
  
  procedure ndu_appnum_results (p_processyear IN INT := 0) IS
    v_file utl_file.file_type;
    
    BEGIN
        IF (p_processyear > 0) THEN
            v_processyear := p_processyear;  
        ELSE
            v_processyear := TISCCOMMON.getProcessYear;
        END IF;
        
        v_file :=utl_file.fopen('OUT_TISC', 'ndu-wa-y12-appnum-results.csv', 'w');
        utl_file.put_line(v_file,'clientnumber,firstname,lastname,dob,student_atar,course_code,course_name,calendaryear,scaled_score');
        
        FOR ndu_match_row IN (select cl.clientnumber from client cl, student st, ext_appnum_csv ea where cl.idn = st.clientid and cl.processyear = v_processyear and cl.clientnumber = ea.ext1_clientnumber)
        LOOP
          FOR record_row IN (select st.clientnumber, cl.name_first, cl.name_last, cl.dob, p.student_atar, sa.subjcode||sa.examind as course_code, sa.subjectdescription as course_name, sa.calendaryear, 
                                    to_number(nvl(decode(sa.FinalScaledMark,'NEC','0',sa.FinalScaledMark),'0'))/100 as scaled_score from client cl join student st on cl.idn = st.clientid join tiscadmin.pep p on st.idn = p.studentid
                                    join subjectsall sa on st.clientnumber = sa.clientnumber and st.processyear = sa.clientprocessyear where cl.processyear = v_processyear and (to_number(nvl(decode(sa.FinalScaledMark,'NEC','0',sa.FinalScaledMark),'0'))/100 > 0)
                                    and st.clientnumber = ndu_match_row.clientnumber)
          LOOP
              utl_file.put_line(v_file,record_row.clientnumber||','||record_row.name_first||','||record_row.name_last||','||record_row.dob||','||
                      record_row.student_atar||','||record_row.course_code||','||REPLACE(record_row.course_name, ',', '')||','||record_row.calendaryear||','||record_row.scaled_score);                                                            
          END LOOP;
        END LOOP;
        
        UTL_FILE.FCLOSE(v_file); 
    END;
        
  procedure ndu_appnum_checked (p_processyear IN INT := 0) IS
    v_date         date;
    v_error_rec    varchar2(1);
    v_match        varchar2(1);

    --client rec
    type ctyp is record (
         clientnumber client.clientnumber%type,
         name_first   client.name_first%type,
         name_last    client.name_last%type,
         dob          client.dob%type
    );

    clrec1 ctyp;
    clrec2 ctyp;

    cursor app_cur is select ea.ext1_clientnumber, ea.ext1_name_first, ea.ext1_name_last, ea.ext1_dob from ext_appnum_csv ea order by ea.ext1_clientnumber;
    
    app_rec app_cur%rowtype;

    v_file utl_file.file_type;

    BEGIN
        IF (p_processyear > 0) THEN
            v_processyear := p_processyear;  
        ELSE
            v_processyear := TISCCOMMON.getProcessYear;
        END IF;
        
        v_file :=utl_file.fopen('OUT_TISC','ndu-wa-y12-appnum-checked.csv','w');

        select sysdate into v_date from dual;
        --utl_file.put_line(v_file,'run date,'||v_date||',AppnumReqMatch');
        utl_file.put_line(v_file,'matchflag,clientnumber,firstname,lastname,dob,TISCclientnumber,TISCfirstname,TISClastname,TISCdob');
                    
        OPEN app_cur;
        LOOP
            FETCH app_cur INTO app_rec;
            EXIT WHEN app_cur%NOTFOUND;
            
            IF (app_rec.ext1_clientnumber != 'clientnumber' and app_rec.ext1_dob != 'dob')  THEN               
              clrec1 := clrec2;                         -- initialise to null
              v_error_rec:='N';
              
              -- read by clientnumber
              begin
                 select cl.clientnumber,cl.name_first,cl.name_last,cl.dob into clrec1 from client cl join student st on cl.idn = st.clientid 
                     where cl.processyear = v_processyear and cl.clientnumber = app_rec.ext1_clientnumber;
                 exception
                 when no_data_found then
                     v_error_rec := 'Y';
              end;
              
              if v_error_rec= 'N' then                  -- clientnumber matched
                  -- fully match
                  begin
                      select cl.clientnumber,cl.name_first,cl.name_last,cl.dob into clrec1 from client cl join student st on cl.idn = st.clientid 
                          where cl.processyear = v_processyear and cl.clientnumber = app_rec.ext1_clientnumber and upper(trim(cl.name_first)) = upper(trim(app_rec.ext1_name_first))
                              and upper(trim(cl.name_last)) = upper(trim(app_rec.ext1_name_last)) and to_char(cl.dob,'dd/mm/yyyy') = trim(app_rec.ext1_dob);
                      exception
                      when no_data_found then
                          v_error_rec := 'Y';
                  end;
                  if v_error_rec= 'N' then                  -- clientname and dob matched
                    v_match := 'Y';
                  else
                    v_match := 'P';
                  end if;
              else
                  v_error_rec:='N';
                  
                  -- read by clientname and dob
                  begin
                      select cl.clientnumber,cl.name_first,cl.name_last,cl.dob into clrec1 from client cl join student st on cl.idn = st.clientid 
                          where cl.processyear = v_processyear and upper(trim(cl.name_first)) = upper(trim(app_rec.ext1_name_first))
                              and upper(trim(cl.name_last)) = upper(trim(app_rec.ext1_name_last)) and to_char(cl.dob,'dd/mm/yyyy') = trim(app_rec.ext1_dob);
                      exception
                      when no_data_found then
                          v_error_rec := 'Y';
                  end;
                            
                  if v_error_rec= 'N' then                  -- clientname and dob matched
                    v_match := 'P';
                  else
                    v_match := 'N';
                  end if;
              end if;
              
              utl_file.put_line(v_file,v_match||','||app_rec.ext1_clientnumber||','||app_rec.ext1_name_first||','||app_rec.ext1_name_last||','||
                    app_rec.ext1_dob||','||clrec1.clientnumber||','||clrec1.name_first||','||clrec1.name_last||','||
                    to_char(clrec1.dob,'dd/mm/yyyy'));  
            END IF;        
        END LOOP;
        UTL_FILE.FCLOSE(v_file);       
    END;

  procedure uwa_uu015_export (p_processyear IN INT := 0) IS
    BEGIN
        IF (p_processyear > 0) THEN
            v_processyear := p_processyear;  
        ELSE
            v_processyear := TISCCOMMON.getProcessYear;
        END IF;
        
        SYSTEM.EXCEL_GENERATOR.clear_workbook;
  
  
        --------------------- worksheet - Y12students ---------------------- 
        SYSTEM.EXCEL_GENERATOR.new_sheet('Y12students');
        
        r := 1;
        SYSTEM.EXCEL_GENERATOR.cell(1, r, 'PROCESSYEAR');
        SYSTEM.EXCEL_GENERATOR.cell(2, r, 'CLIENTNUMBER');
        SYSTEM.EXCEL_GENERATOR.cell(3, r, 'IND_TYPE');
        SYSTEM.EXCEL_GENERATOR.cell(4, r, 'GENDER');
        SYSTEM.EXCEL_GENERATOR.cell(5, r, 'DOB');
        SYSTEM.EXCEL_GENERATOR.cell(6, r, 'PROVIDERSCHOOL');
        --SYSTEM.EXCEL_GENERATOR.freeze_rows( 1 );
        
        FOR rec IN  (SELECT cl.processyear, cl.clientnumber, cl.ind_type, cl.gender, cl.dob, st.providerschool FROM CLIENT cl 
                                  JOIN STUDENT st ON cl.idn = st.clientid AND cl.processyear = v_processyear
                                       WHERE cl.ind_type in ('1','2') AND nvl(st.status_cc,'!') != '2' ORDER BY cl.clientnumber) LOOP
                                       
                  r := r+1;
                  SYSTEM.EXCEL_GENERATOR.cell(1, r, rec.processyear);
                  SYSTEM.EXCEL_GENERATOR.cell(2, r, rec.clientnumber);
                  SYSTEM.EXCEL_GENERATOR.cell(3, r, rec.ind_type);
                  SYSTEM.EXCEL_GENERATOR.cell(4, r, rec.gender);
                  SYSTEM.EXCEL_GENERATOR.cell(5, r, TO_CHAR(rec.dob, 'dd/mm/yyyy'));
                  SYSTEM.EXCEL_GENERATOR.cell(6, r, rec.providerschool);
                  
        END LOOP;
        DBMS_OUTPUT.PUT_LINE('total row number of Y12students: ' || r);
        
        --------------------- worksheet - CUR TEA ---------------------- 
        SYSTEM.EXCEL_GENERATOR.new_sheet('CUR TEA');
        
        r := 1;
        SYSTEM.EXCEL_GENERATOR.cell(1, r, 'CLIENTNUMBER');
        SYSTEM.EXCEL_GENERATOR.cell(2, r, 'BEST_PASTYR');
        SYSTEM.EXCEL_GENERATOR.cell(3, r, 'TEACAT');
        SYSTEM.EXCEL_GENERATOR.cell(4, r, 'INSTITUTIONID');
        SYSTEM.EXCEL_GENERATOR.cell(5, r, 'TEATYPE');
        SYSTEM.EXCEL_GENERATOR.cell(6, r, 'AGGR');
        SYSTEM.EXCEL_GENERATOR.cell(7, r, 'TER');
        
        FOR rec IN  (SELECT cl.clientnumber, pb.best_pastyr, tb.teacat, tb.institutionid, t.teatype, t.aggr, t.ter FROM CLIENT cl 
                            JOIN STUDENT st ON cl.idn = st.clientid AND cl.processyear = v_processyear
                            JOIN pep_base pb ON st.idn = pb.studentid AND pb.institutionid = '1'
                            JOIN tea_base tb ON st.idn = tb.studentid AND tb.institutionid = '1'
                            JOIN tea t ON tb.best_teaid = t.idn
                                 WHERE cl.ind_type in ('1','2') AND nvl(st.status_cc,'!') != '2' ORDER BY cl.clientnumber) LOOP
                                 
            r := r+1;
            SYSTEM.EXCEL_GENERATOR.cell(1, r, rec.clientnumber);
            SYSTEM.EXCEL_GENERATOR.cell(2, r, rec.best_pastyr);
            SYSTEM.EXCEL_GENERATOR.cell(3, r, rec.teacat);
            SYSTEM.EXCEL_GENERATOR.cell(4, r, rec.institutionid);
            SYSTEM.EXCEL_GENERATOR.cell(5, r, rec.teatype);
            SYSTEM.EXCEL_GENERATOR.cell(6, r, rec.aggr);
            SYSTEM.EXCEL_GENERATOR.cell(7, r, rec.ter);
            
        END LOOP;
        DBMS_OUTPUT.PUT_LINE('total row number of CUR TEA: ' || r);
        
        
        --------------------- worksheet - UWA TEA ---------------------- 
        SYSTEM.EXCEL_GENERATOR.new_sheet('UWA TEA');
        
        r := 1;
        SYSTEM.EXCEL_GENERATOR.cell(1, r, 'CLIENTNUMBER');
        SYSTEM.EXCEL_GENERATOR.cell(2, r, 'BEST_PASTYR');
        SYSTEM.EXCEL_GENERATOR.cell(3, r, 'TEACAT');
        SYSTEM.EXCEL_GENERATOR.cell(4, r, 'INSTITUTIONID');
        SYSTEM.EXCEL_GENERATOR.cell(5, r, 'TEATYPE');
        SYSTEM.EXCEL_GENERATOR.cell(6, r, 'AGGR');
        SYSTEM.EXCEL_GENERATOR.cell(7, r, 'TER');
        
        FOR rec IN  (SELECT cl.clientnumber, pb.best_pastyr, tb.teacat, tb.institutionid, t.teatype, t.aggr, t.ter FROM CLIENT cl 
                            JOIN STUDENT st ON cl.idn = st.clientid AND cl.processyear = v_processyear
                            JOIN pep_base pb ON st.idn = pb.studentid AND pb.institutionid = '4'
                            JOIN tea_base tb ON st.idn = tb.studentid AND tb.institutionid = '4'
                            JOIN tea t ON tb.best_teaid = t.idn
                                 WHERE cl.ind_type in ('1','2') AND nvl(st.status_cc,'!') != '2' ORDER BY cl.clientnumber) LOOP
                                 
            r := r+1;
            SYSTEM.EXCEL_GENERATOR.cell(1, r, rec.clientnumber);
            SYSTEM.EXCEL_GENERATOR.cell(2, r, rec.best_pastyr);
            SYSTEM.EXCEL_GENERATOR.cell(3, r, rec.teacat);
            SYSTEM.EXCEL_GENERATOR.cell(4, r, rec.institutionid);
            SYSTEM.EXCEL_GENERATOR.cell(5, r, rec.teatype);
            SYSTEM.EXCEL_GENERATOR.cell(6, r, rec.aggr);
            SYSTEM.EXCEL_GENERATOR.cell(7, r, rec.ter);
            
        END LOOP;
        DBMS_OUTPUT.PUT_LINE('total row number of UWA TEA: ' || r);
        
        
        --------------------- worksheet - Orig Prefs ---------------------- 
        SYSTEM.EXCEL_GENERATOR.new_sheet('Orig Prefs');
        r := 1;
        SYSTEM.EXCEL_GENERATOR.cell(1, r, 'ADMISSIONYEAR');
        SYSTEM.EXCEL_GENERATOR.cell(2, r, 'CLIENTNUMBER');
        SYSTEM.EXCEL_GENERATOR.cell(3, r, 'SETNUM');
        SYSTEM.EXCEL_GENERATOR.cell(4, r, 'PREFNUM');
        SYSTEM.EXCEL_GENERATOR.cell(5, r, 'PREF');
        SYSTEM.EXCEL_GENERATOR.cell(6, r, 'SETINSDATE');
        
        FOR rec IN  (SELECT a.admissionyear, cl.clientnumber, p.setnum, p.prefnum, p.pref, p.setinsdate FROM CLIENT cl 
                                  JOIN application a ON cl.idn = a.clientid AND cl.processyear = v_processyear
                                  JOIN preferences p ON a.idn = p.applicationid
                                  JOIN STUDENT st ON cl.idn = st.clientid AND nvl(st.status_cc,'!') != '2' 
                                       WHERE cl.ind_type = '2' AND p.setnum = 0 ORDER BY cl.clientnumber, p.prefnum) LOOP
                                       
                  r := r+1;
                  SYSTEM.EXCEL_GENERATOR.cell(1, r, rec.admissionyear);
                  SYSTEM.EXCEL_GENERATOR.cell(2, r, rec.clientnumber);
                  SYSTEM.EXCEL_GENERATOR.cell(3, r, rec.setnum);
                  SYSTEM.EXCEL_GENERATOR.cell(4, r, rec.prefnum);
                  SYSTEM.EXCEL_GENERATOR.cell(5, r, rec.pref);
                  SYSTEM.EXCEL_GENERATOR.cell(6, r, TO_CHAR(rec.setinsdate, 'dd/mm/yyyy HH24:MI:SS'));
                  
        END LOOP;
        DBMS_OUTPUT.PUT_LINE('total row number of Orig Prefs: ' || r);
        
        
        --------------------- worksheet - Curr Prefs ----------------------
        SYSTEM.EXCEL_GENERATOR.new_sheet('Curr Prefs');
        r := 1;
        SYSTEM.EXCEL_GENERATOR.cell(1, r, 'ADMISSIONYEAR');
        SYSTEM.EXCEL_GENERATOR.cell(2, r, 'CLIENTNUMBER');
        SYSTEM.EXCEL_GENERATOR.cell(3, r, 'SETNUM');
        SYSTEM.EXCEL_GENERATOR.cell(4, r, 'ISCURRENT');
        SYSTEM.EXCEL_GENERATOR.cell(5, r, 'PREFNUM');
        SYSTEM.EXCEL_GENERATOR.cell(6, r, 'PREF');
        SYSTEM.EXCEL_GENERATOR.cell(7, r, 'SETINSDATE');
        
        FOR rec IN  (SELECT a.admissionyear, cl.clientnumber, p.setnum, p.iscurrent, p.prefnum, p.pref, p.setinsdate FROM CLIENT cl 
                            JOIN application a ON cl.idn = a.clientid AND cl.processyear = v_processyear
                            JOIN preferences p ON a.idn = p.applicationid
                            JOIN STUDENT st ON cl.idn = st.clientid AND nvl(st.status_cc,'!') != '2' 
                                 WHERE cl.ind_type = '2' AND p.iscurrent = 'Y' ORDER BY cl.clientnumber, p.prefnum) LOOP
                                 
            r := r+1;
            SYSTEM.EXCEL_GENERATOR.cell(1, r, rec.admissionyear);
            SYSTEM.EXCEL_GENERATOR.cell(2, r, rec.clientnumber);
            SYSTEM.EXCEL_GENERATOR.cell(3, r, rec.setnum);
            SYSTEM.EXCEL_GENERATOR.cell(4, r, rec.iscurrent);
            SYSTEM.EXCEL_GENERATOR.cell(5, r, rec.prefnum);
            SYSTEM.EXCEL_GENERATOR.cell(6, r, rec.pref);
            SYSTEM.EXCEL_GENERATOR.cell(7, r, TO_CHAR(rec.setinsdate, 'dd/mm/yyyy HH24:MI:SS'));     
        END LOOP;
        DBMS_OUTPUT.PUT_LINE('total row number of Curr Prefs: ' || r);
        
        
        --------------------flush data into file, and close file ------------------------
        SYSTEM.EXCEL_GENERATOR.save( 'OUT_TISC', 'uu015.xlsx' );
        
    END;
  
  procedure ux002_export (p_inst in varchar2) IS 
    BEGIN
        v_processyear := TISCCOMMON.getProcessYear;
        COURSELIST.printcourse(v_processyear, p_inst);
    END;
   
  procedure course_code_export (p_admissionyear IN INT := 0) IS
    BEGIN
        IF (p_admissionyear > 0) THEN
            v_admissionyear := p_admissionyear;  
        ELSE
            v_admissionyear := TISCCOMMON.getAdmissionYearUg;
        END IF;
        
        SYSTEM.EXCEL_GENERATOR.clear_workbook;
  
  
        --------------------- worksheet - TISC Course Code ---------------------- 
        SYSTEM.EXCEL_GENERATOR.new_sheet('TISC Course Code');

        r := 1;
        SYSTEM.EXCEL_GENERATOR.cell(1, r, 'Inco A');
        SYSTEM.EXCEL_GENERATOR.cell(2, r, 'Short Name');
        SYSTEM.EXCEL_GENERATOR.cell(3, r, 'Inco Name');
        SYSTEM.EXCEL_GENERATOR.cell(4, r, 'Deet Fos');
        SYSTEM.EXCEL_GENERATOR.cell(5, r, 'Mos F');
        SYSTEM.EXCEL_GENERATOR.cell(6, r, 'Mos P');
        SYSTEM.EXCEL_GENERATOR.cell(7, r, 'Mos E');
        SYSTEM.EXCEL_GENERATOR.cell(8, r, 'Status');
        SYSTEM.EXCEL_GENERATOR.cell(9, r, 'Course Type');
        SYSTEM.EXCEL_GENERATOR.cell(10, r, 'Admissionyear');
        --SYSTEM.EXCEL_GENERATOR.freeze_rows( 1 );
        
        FOR rec IN  (SELECT inco_a, short_name, inco_name, deet_fos, mos_f, mos_p, mos_e, status, course_type, admissionyear FROM codeii 
                                  WHERE admissionyear = v_admissionyear and status in ('Y','N') ORDER BY inco_a) LOOP
                                       
                  r := r+1;
                  SYSTEM.EXCEL_GENERATOR.cell(1, r, rec.inco_a);
                  SYSTEM.EXCEL_GENERATOR.cell(2, r, rec.short_name);
                  SYSTEM.EXCEL_GENERATOR.cell(3, r, rec.inco_name);
                  SYSTEM.EXCEL_GENERATOR.cell(4, r, rec.deet_fos);
                  SYSTEM.EXCEL_GENERATOR.cell(5, r, rec.mos_f);
                  SYSTEM.EXCEL_GENERATOR.cell(6, r, rec.mos_p);
                  SYSTEM.EXCEL_GENERATOR.cell(7, r, rec.mos_e);
                  SYSTEM.EXCEL_GENERATOR.cell(8, r, rec.status);
                  SYSTEM.EXCEL_GENERATOR.cell(9, r, rec.course_type);
                  SYSTEM.EXCEL_GENERATOR.cell(10, r, rec.admissionyear);
                  
        END LOOP;
        DBMS_OUTPUT.PUT_LINE('total row number of course code: ' || r);      
    
        --------------------flush data into file, and close file ------------------------
        SYSTEM.EXCEL_GENERATOR.save( 'TISCCODES', 'CourseCode.xlsx' );
        
    END;
  
  
  
  procedure arts_institution_code_export (p_processyear IN INT := 0) IS
    BEGIN
        IF (p_processyear > 0) THEN
            v_processyear := p_processyear;  
        ELSE
            v_processyear := TISCCOMMON.getProcessYear;
        END IF;
        
        SYSTEM.EXCEL_GENERATOR.clear_workbook;
  
  
        --------------------- worksheet - qfscode ---------------------- 
        SYSTEM.EXCEL_GENERATOR.new_sheet('qfscode');

        r := 1;
        SYSTEM.EXCEL_GENERATOR.cell(1, r, 'Qc Type');
        SYSTEM.EXCEL_GENERATOR.cell(2, r, 'Deet1');
        SYSTEM.EXCEL_GENERATOR.cell(3, r, 'State');
        SYSTEM.EXCEL_GENERATOR.cell(4, r, 'Name5');
        SYSTEM.EXCEL_GENERATOR.cell(5, r, 'Name50');
        SYSTEM.EXCEL_GENERATOR.cell(6, r, 'Status');
        SYSTEM.EXCEL_GENERATOR.cell(7, r, 'Processyear');
        --SYSTEM.EXCEL_GENERATOR.freeze_rows( 1 );
        
        FOR rec IN  (SELECT qc_type, deet1, state, name5, name50, status FROM qfscode ORDER BY name50) LOOP
                                       
                  r := r+1;
                  SYSTEM.EXCEL_GENERATOR.cell(1, r, rec.qc_type);
                  SYSTEM.EXCEL_GENERATOR.cell(2, r, rec.deet1);
                  SYSTEM.EXCEL_GENERATOR.cell(3, r, rec.state);
                  SYSTEM.EXCEL_GENERATOR.cell(4, r, rec.name5);
                  SYSTEM.EXCEL_GENERATOR.cell(5, r, rec.name50);
                  SYSTEM.EXCEL_GENERATOR.cell(6, r, rec.status);
                  SYSTEM.EXCEL_GENERATOR.cell(7, r, v_processyear);
                  
        END LOOP;
        DBMS_OUTPUT.PUT_LINE('total row number of arts inistitution code: ' || r);      
    
        --------------------flush data into file, and close file ------------------------
        SYSTEM.EXCEL_GENERATOR.save( 'TISCCODES', 'ARTSInstitutionCode.xlsx' );
        
    END;
    
  procedure country_code_export (p_processyear IN INT := 0) IS
    BEGIN
        IF (p_processyear > 0) THEN
            v_processyear := p_processyear;  
        ELSE
            v_processyear := TISCCOMMON.getProcessYear;
        END IF;
        
        SYSTEM.EXCEL_GENERATOR.clear_workbook;
  
  
        --------------------- worksheet - miscodes ---------------------- 
        SYSTEM.EXCEL_GENERATOR.new_sheet('Miscodes');

        r := 1;
        SYSTEM.EXCEL_GENERATOR.cell(1, r, 'scode');
        SYSTEM.EXCEL_GENERATOR.cell(2, r, 'Name50');
        SYSTEM.EXCEL_GENERATOR.cell(3, r, 'Status');
        SYSTEM.EXCEL_GENERATOR.cell(4, r, 'Fromyear');
        SYSTEM.EXCEL_GENERATOR.cell(5, r, 'Toyear');
        SYSTEM.EXCEL_GENERATOR.cell(6, r, 'Processyear');
        --SYSTEM.EXCEL_GENERATOR.freeze_rows( 1 );
        
        FOR rec IN  (SELECT substr(code,2,4) scode, name50, status, fromyear, toyear FROM miscodes 
            WHERE substr(code,1,1) = 'C' AND toyear > (v_processyear - 1) ORDER BY name50) LOOP
                                       
                  r := r+1;
                  SYSTEM.EXCEL_GENERATOR.cell(1, r, rec.scode);
                  SYSTEM.EXCEL_GENERATOR.cell(2, r, rec.name50);
                  SYSTEM.EXCEL_GENERATOR.cell(3, r, rec.status);
                  SYSTEM.EXCEL_GENERATOR.cell(4, r, rec.fromyear);
                  SYSTEM.EXCEL_GENERATOR.cell(5, r, rec.toyear);
                  SYSTEM.EXCEL_GENERATOR.cell(6, r, v_processyear);
                  
        END LOOP;
        DBMS_OUTPUT.PUT_LINE('total row number of country code: ' || r);      
    
        --------------------flush data into file, and close file ------------------------
        SYSTEM.EXCEL_GENERATOR.save( 'TISCCODES', 'CountryCode.xlsx' );
        
    END;  
    
    
  procedure lang_code_export (p_processyear IN INT := 0) IS
    BEGIN
        IF (p_processyear > 0) THEN
            v_processyear := p_processyear;  
        ELSE
            v_processyear := TISCCOMMON.getProcessYear;
        END IF;
        
        SYSTEM.EXCEL_GENERATOR.clear_workbook;
  
  
        --------------------- worksheet - miscodes ---------------------- 
        SYSTEM.EXCEL_GENERATOR.new_sheet('Miscodes');

        r := 1;
        SYSTEM.EXCEL_GENERATOR.cell(1, r, 'scode');
        SYSTEM.EXCEL_GENERATOR.cell(2, r, 'Name50');
        SYSTEM.EXCEL_GENERATOR.cell(3, r, 'Status');
        SYSTEM.EXCEL_GENERATOR.cell(4, r, 'Fromyear');
        SYSTEM.EXCEL_GENERATOR.cell(5, r, 'Toyear');
        SYSTEM.EXCEL_GENERATOR.cell(6, r, 'Processyear');
        --SYSTEM.EXCEL_GENERATOR.freeze_rows( 1 );
        
        FOR rec IN  (SELECT substr(code,2,4) scode, name50, status, fromyear, toyear FROM miscodes 
            WHERE substr(code,1,1) = 'L' AND toyear > (v_processyear - 1) ORDER BY name50) LOOP
                                       
                  r := r+1;
                  SYSTEM.EXCEL_GENERATOR.cell(1, r, rec.scode);
                  SYSTEM.EXCEL_GENERATOR.cell(2, r, rec.name50);
                  SYSTEM.EXCEL_GENERATOR.cell(3, r, rec.status);
                  SYSTEM.EXCEL_GENERATOR.cell(4, r, rec.fromyear);
                  SYSTEM.EXCEL_GENERATOR.cell(5, r, rec.toyear);
                  SYSTEM.EXCEL_GENERATOR.cell(6, r, v_processyear);
                  
        END LOOP;
        DBMS_OUTPUT.PUT_LINE('total row number of lang code: ' || r);      
    
        --------------------flush data into file, and close file ------------------------
        SYSTEM.EXCEL_GENERATOR.save( 'TISCCODES', 'LangCode.xlsx' );
        
    END;  
    
    
  procedure school_code_export (p_processyear IN INT := 0) IS
    BEGIN
        IF (p_processyear > 0) THEN
            v_processyear := p_processyear;  
        ELSE
            v_processyear := TISCCOMMON.getProcessYear;
        END IF;
        
        SYSTEM.EXCEL_GENERATOR.clear_workbook;
  
  
        --------------------- worksheet - School Code ---------------------- 
        SYSTEM.EXCEL_GENERATOR.new_sheet('School Code');

        r := 1;
        SYSTEM.EXCEL_GENERATOR.cell(1, r, 'School');
        SYSTEM.EXCEL_GENERATOR.cell(2, r, 'Principal');
        SYSTEM.EXCEL_GENERATOR.cell(3, r, 'School Name');
        SYSTEM.EXCEL_GENERATOR.cell(4, r, 'Addr1');
        SYSTEM.EXCEL_GENERATOR.cell(5, r, 'Addr2');
        SYSTEM.EXCEL_GENERATOR.cell(6, r, 'Suburb');
        SYSTEM.EXCEL_GENERATOR.cell(7, r, 'Post');
        SYSTEM.EXCEL_GENERATOR.cell(8, r, 'Gov Ind');
        SYSTEM.EXCEL_GENERATOR.cell(9, r, 'Processyear');
        --SYSTEM.EXCEL_GENERATOR.freeze_rows( 1 );
        
        FOR rec IN  (SELECT s.school, s.principal, s.sch_name, s.addr1, s.addr2, suburb, post, gov_ind FROM school s 
          WHERE iscurrent = 'Y' ORDER BY s.school) LOOP
                                       
                  r := r+1;
                  SYSTEM.EXCEL_GENERATOR.cell(1, r, rec.school);
                  SYSTEM.EXCEL_GENERATOR.cell(2, r, rec.principal);
                  SYSTEM.EXCEL_GENERATOR.cell(3, r, rec.sch_name);
                  SYSTEM.EXCEL_GENERATOR.cell(4, r, rec.addr1);
                  SYSTEM.EXCEL_GENERATOR.cell(5, r, rec.addr2);
                  SYSTEM.EXCEL_GENERATOR.cell(6, r, rec.suburb);
                  SYSTEM.EXCEL_GENERATOR.cell(7, r, rec.post);
                  SYSTEM.EXCEL_GENERATOR.cell(8, r, rec.gov_ind);
                  SYSTEM.EXCEL_GENERATOR.cell(9, r, v_processyear);
                  
        END LOOP;
        DBMS_OUTPUT.PUT_LINE('total row number of school code: ' || r);      
    
        --------------------flush data into file, and close file ------------------------
        SYSTEM.EXCEL_GENERATOR.save( 'TISCCODES', 'SchoolCode.xlsx' );
        
    END;  
    
    
  procedure subj_code_export (p_processyear IN INT := 0) IS
    BEGIN
        IF (p_processyear > 0) THEN
            v_processyear := p_processyear;  
        ELSE
            v_processyear := TISCCOMMON.getProcessYear;
        END IF;
        
        SYSTEM.EXCEL_GENERATOR.clear_workbook;
  
  
        --------------------- worksheet - School Code ---------------------- 
        SYSTEM.EXCEL_GENERATOR.new_sheet('Subj Code');

        r := 1;
        SYSTEM.EXCEL_GENERATOR.cell(1, r, 'Subj Num');
        SYSTEM.EXCEL_GENERATOR.cell(2, r, 'Year Id');
        SYSTEM.EXCEL_GENERATOR.cell(3, r, 'Short Name');
        SYSTEM.EXCEL_GENERATOR.cell(4, r, 'Tes Ind');
        SYSTEM.EXCEL_GENERATOR.cell(5, r, 'Long Name');
        SYSTEM.EXCEL_GENERATOR.cell(6, r, 'Processyear');
        --SYSTEM.EXCEL_GENERATOR.freeze_rows( 1 );
        
        FOR rec IN  (SELECT c.subj_num, c.year_id, c.short_name, c.tes_ind, c.long_name, c.processyear FROM codsubj c 
           WHERE c.processyear = v_processyear ORDER BY c.short_name) LOOP
                                       
                  r := r+1;
                  SYSTEM.EXCEL_GENERATOR.cell(1, r, rec.subj_num);
                  SYSTEM.EXCEL_GENERATOR.cell(2, r, rec.year_id);
                  SYSTEM.EXCEL_GENERATOR.cell(3, r, rec.short_name);
                  SYSTEM.EXCEL_GENERATOR.cell(4, r, rec.tes_ind);
                  SYSTEM.EXCEL_GENERATOR.cell(5, r, rec.long_name);
                  SYSTEM.EXCEL_GENERATOR.cell(6, r, rec.processyear);
                  
        END LOOP;
        DBMS_OUTPUT.PUT_LINE('total row number of subj code: ' || r);      
    
        --------------------flush data into file, and close file ------------------------
        SYSTEM.EXCEL_GENERATOR.save( 'TISCCODES', 'SubjCode.xlsx' );
        
    END;
  
  
  procedure basis_admission_code_export (p_processyear IN INT := 0) IS
    BEGIN
        IF (p_processyear > 0) THEN
            v_processyear := p_processyear;  
        ELSE
            v_processyear := TISCCOMMON.getProcessYear;
        END IF;
        
        SYSTEM.EXCEL_GENERATOR.clear_workbook;
  
  
        --------------------- worksheet - School Code ---------------------- 
        SYSTEM.EXCEL_GENERATOR.new_sheet('Basis of Admission Code');

        r := 1;
        SYSTEM.EXCEL_GENERATOR.cell(1, r, 'InstitutionID');
        SYSTEM.EXCEL_GENERATOR.cell(2, r, 'Admissioncategory');
        SYSTEM.EXCEL_GENERATOR.cell(3, r, 'Description');
        SYSTEM.EXCEL_GENERATOR.cell(4, r, 'Yearfrom');
        SYSTEM.EXCEL_GENERATOR.cell(5, r, 'Yearto');
        SYSTEM.EXCEL_GENERATOR.cell(6, r, 'Status');
        SYSTEM.EXCEL_GENERATOR.cell(7, r, 'Deewr Basisofadmission');
        SYSTEM.EXCEL_GENERATOR.cell(8, r, 'Processyear');
        --SYSTEM.EXCEL_GENERATOR.freeze_rows( 1 );
        
        FOR rec IN (SELECT b.admissioncategory, b.description, b.yearfrom, i.instshortname, b.yearto, b.status, b.deewr_basisofadmission 
                                  FROM institution_basisofadmission b, institution i where b.institutionid = i.idn) LOOP
                                       
                  r := r+1;
                  SYSTEM.EXCEL_GENERATOR.cell(1, r, rec.instshortname);
                  SYSTEM.EXCEL_GENERATOR.cell(2, r, rec.admissioncategory);
                  SYSTEM.EXCEL_GENERATOR.cell(3, r, rec.description);
                  SYSTEM.EXCEL_GENERATOR.cell(4, r, rec.yearfrom);
                  SYSTEM.EXCEL_GENERATOR.cell(5, r, rec.yearto);
                  SYSTEM.EXCEL_GENERATOR.cell(6, r, rec.status);
                  SYSTEM.EXCEL_GENERATOR.cell(7, r, rec.deewr_basisofadmission);
                  SYSTEM.EXCEL_GENERATOR.cell(8, r, v_processyear);
                  
        END LOOP;
        DBMS_OUTPUT.PUT_LINE('total row number of basis admission code: ' || r);      
    
        --------------------flush data into file, and close file ------------------------
        SYSTEM.EXCEL_GENERATOR.save( 'TISCCODES', 'BasisOfAdmissionCode.xlsx' );
        
    END;
    
  /* export basis admission code into csv format
  procedure basis_admission_code_export_csv (p_processyear IN INT := 0) IS
    v_file utl_file.file_type;
    
    BEGIN
        IF (p_processyear > 0) THEN
            v_processyear := p_processyear;  
        ELSE
            v_processyear := TISCCOMMON.getProcessYear;
        END IF;
        
        v_file := utl_file.fopen('TISCCODES', 'BasisOfAdmissionCode.csv', 'W');
        utl_file.put_line(v_file,'ADMISSIONCATEGORY,DESCRIPTION,FROMYEAR,INSTITUTION,TOYEAR,STATUS,DEEWR_BASISOFADMISSION,PROCESSYEAR');
        
        FOR record_row IN (select b.admissioncategory, b.description, b.fromyear, b.institution, b.toyear, b.status, b.deewr_basisofadmission 
                                  from basisofadmission b)
        LOOP
              utl_file.put_line(v_file,record_row.admissioncategory||','||REPLACE(record_row.description, ',', '')||','||record_row.fromyear||','||record_row.institution||','||
                      record_row.toyear||','||record_row.status||','||record_row.deewr_basisofadmission||','||v_processyear);                                                            
        END LOOP;
        utl_file.fclose(v_file);  
    END;
  */


  procedure ecu_prefs_report IS
    v_from  VARCHAR2(30);
    v_to    VARCHAR2(30);          
  
    BEGIN
        SYSTEM.EXCEL_GENERATOR.clear_workbook;
  
  
        --------------------- worksheet - School Code ---------------------- 
        SYSTEM.EXCEL_GENERATOR.new_sheet('ECU Preferences Report');

        r := 1;
        
        select 'From '||sc.ecu_pref_date1 into v_from from system_config sc;
        select ' To '||(sysdate - 1) into v_to from dual;
        SYSTEM.EXCEL_GENERATOR.cell(1, r, v_from || v_to);
        
        r := r+1;
        SYSTEM.EXCEL_GENERATOR.cell(1, r, 'Processyear');
        SYSTEM.EXCEL_GENERATOR.cell(2, r, 'Clientnumber');
        SYSTEM.EXCEL_GENERATOR.cell(3, r, 'Title');
        SYSTEM.EXCEL_GENERATOR.cell(4, r, 'Last name');
        SYSTEM.EXCEL_GENERATOR.cell(5, r, 'First name');
        SYSTEM.EXCEL_GENERATOR.cell(6, r, 'Second name');
        SYSTEM.EXCEL_GENERATOR.cell(7, r, 'DOB');
        SYSTEM.EXCEL_GENERATOR.cell(8, r, 'Street1');
        SYSTEM.EXCEL_GENERATOR.cell(9, r, 'Street2');
        SYSTEM.EXCEL_GENERATOR.cell(10, r, 'Suburb');
        SYSTEM.EXCEL_GENERATOR.cell(11, r, 'Country');
        SYSTEM.EXCEL_GENERATOR.cell(12, r, 'Post');
        SYSTEM.EXCEL_GENERATOR.cell(13, r, 'State');
        SYSTEM.EXCEL_GENERATOR.cell(14, r, 'Phone home');
        SYSTEM.EXCEL_GENERATOR.cell(15, r, 'Phone work');
        SYSTEM.EXCEL_GENERATOR.cell(16, r, 'Phone Mobile');
        SYSTEM.EXCEL_GENERATOR.cell(17, r, 'Admissionyear');
        SYSTEM.EXCEL_GENERATOR.cell(18, r, 'Pref num');
        SYSTEM.EXCEL_GENERATOR.cell(19, r, 'Pref');
        SYSTEM.EXCEL_GENERATOR.cell(20, r, 'Mos');
        SYSTEM.EXCEL_GENERATOR.cell(21, r, 'Cat');
        SYSTEM.EXCEL_GENERATOR.cell(22, r, 'Ter');
        SYSTEM.EXCEL_GENERATOR.cell(23, r, 'Pref insert date');
        SYSTEM.EXCEL_GENERATOR.cell(24, r, 'Set insert date');
        SYSTEM.EXCEL_GENERATOR.cell(25, r, 'Email');
        SYSTEM.EXCEL_GENERATOR.cell(26, r, 'Comments');
        --SYSTEM.EXCEL_GENERATOR.freeze_rows( 1 );
        
        FOR ecu_pref_row IN (select cl.PROCESSYEAR, cl.CLIENTNUMBER, cl.TITLE, cl.NAME_LAST, cl.NAME_FIRST, cl.NAME_SEC, 
                                    cl.dob, ad.STREET1, ad.STREET2, ad.SUBURB, ad.COUNTRY, ad.POST, ad.STATE, 
                                    cl.PHONEHOME, cl.PHONEWORK, cl.PHONEMOBILE, ap.ADMISSIONYEAR, 
                                    p.PREFNUM, p.PREF, p.MOS, p.CAT, p.TER, p.PREFINSDATE, p.setinsdate, cl.email, ec.comments
                              FROM TISCADMIN.CLIENT cl join TISCADMIN.APPLICATION ap on cl.IDN = ap.CLIENTID and ap.IND_WD IS NULL
                                   left join tiscadmin.lastcommentecu ec on cl.idn = ec.clientid
                                   left join TISCADMIN.ADDRESS ad on cl.IDN = ad.CLIENTID and ad.addr_type = 'N' and ad.ISCURRENT = 'Y'
                                   join TISCADMIN.PREFERENCES p on ap.IDN = p.APPLICATIONID
                              WHERE cl.processyear = TISCCOMMON.getProcessYear and p.ISCURRENT = 'Y' and p.PREF in ('EAMUC', 'EACPV', 'EAMJC', 'EAMCC', 'EACMC', 'EAMAC', 'EAAMC', 'EAPAP') 
                                    and ((to_char(p.prefinsdate,'yyyymmdd') >= (select to_char(sc.ecu_pref_date1,'yyyymmdd') from tiscadmin.system_config sc)
                                        or to_char(p.setinsdate,'yyyymmdd') >= (select to_char(sc.ecu_pref_date1,'yyyymmdd') from tiscadmin.system_config sc))
                                    and (to_char(p.prefinsdate,'yyyymmdd') < (select to_char(sysdate,'yyyymmdd') from dual)
                                    and to_char(p.setinsdate,'yyyymmdd') < (select to_char(sysdate,'yyyymmdd') from dual)))
                              ORDER BY cl.PROCESSYEAR ASC, p.prefinsdate, p.setinsdate, p.PREF ASC, p.prefnum asc        
                              )  LOOP
                     
                  r := r+1;
                  SYSTEM.EXCEL_GENERATOR.cell(1, r, ecu_pref_row.PROCESSYEAR);
                  SYSTEM.EXCEL_GENERATOR.cell(2, r, ecu_pref_row.clientnumber);
                  SYSTEM.EXCEL_GENERATOR.cell(3, r, ecu_pref_row.title);
                  SYSTEM.EXCEL_GENERATOR.cell(4, r, ecu_pref_row.name_last);
                  SYSTEM.EXCEL_GENERATOR.cell(5, r, ecu_pref_row.name_first);
                  SYSTEM.EXCEL_GENERATOR.cell(6, r, ecu_pref_row.name_sec);
                  SYSTEM.EXCEL_GENERATOR.cell(7, r, ecu_pref_row.dob);
                  SYSTEM.EXCEL_GENERATOR.cell(8, r, ecu_pref_row.street1);
                  SYSTEM.EXCEL_GENERATOR.cell(9, r, ecu_pref_row.street2);
                  SYSTEM.EXCEL_GENERATOR.cell(10, r, ecu_pref_row.suburb);
                  SYSTEM.EXCEL_GENERATOR.cell(11, r, ecu_pref_row.country);
                  SYSTEM.EXCEL_GENERATOR.cell(12, r, ecu_pref_row.post);
                  SYSTEM.EXCEL_GENERATOR.cell(13, r, ecu_pref_row.state);
                  SYSTEM.EXCEL_GENERATOR.cell(14, r, ecu_pref_row.phonehome);
                  SYSTEM.EXCEL_GENERATOR.cell(15, r, ecu_pref_row.phonework);
                  SYSTEM.EXCEL_GENERATOR.cell(16, r, ecu_pref_row.phonemobile);
                  SYSTEM.EXCEL_GENERATOR.cell(17, r, ecu_pref_row.admissionyear);
                  SYSTEM.EXCEL_GENERATOR.cell(18, r, ecu_pref_row.prefnum);
                  SYSTEM.EXCEL_GENERATOR.cell(19, r, ecu_pref_row.pref);
                  SYSTEM.EXCEL_GENERATOR.cell(20, r, ecu_pref_row.mos);
                  SYSTEM.EXCEL_GENERATOR.cell(21, r, ecu_pref_row.cat);
                  SYSTEM.EXCEL_GENERATOR.cell(22, r, ecu_pref_row.ter);
                  SYSTEM.EXCEL_GENERATOR.cell(23, r, ecu_pref_row.prefinsdate);
                  SYSTEM.EXCEL_GENERATOR.cell(24, r, ecu_pref_row.setinsdate);
                  SYSTEM.EXCEL_GENERATOR.cell(25, r, ecu_pref_row.email);
                  SYSTEM.EXCEL_GENERATOR.cell(26, r, ecu_pref_row.comments);
                  
        END LOOP;
        DBMS_OUTPUT.PUT_LINE('total row number of ecu prefs report: ' || r);      
    
        --------------------flush data into file, and close file ------------------------
        SYSTEM.EXCEL_GENERATOR.save( 'ECUPREFS', 'ecuprefs1.' || TO_CHAR (SYSDATE, 'ddmmyy') || '.xlsx');
        
    END;
    
  /* export ecu preferences report into csv format      
  procedure ecu_prefs_report_csv IS
    v_file  utl_file.file_type;
    v_from  VARCHAR2(30);
    v_to    VARCHAR2(30);
    
    BEGIN
        v_file := utl_file.fopen('ECUPREFS', 'ecuprefs1.' || TO_CHAR (SYSDATE, 'ddmmyy') || '.csv', 'W');
        
        select 'From '||sc.ecu_pref_date1 into v_from from system_config sc;
        utl_file.put_line(v_file, v_from);
        
        select 'To '||(sysdate - 1) into v_to from dual;
        utl_file.put_line(v_file, v_to);
        
        FOR ecu_pref_row IN (select cl.PROCESSYEAR, cl.CLIENTNUMBER, cl.TITLE, cl.NAME_LAST, cl.NAME_FIRST, cl.NAME_SEC, 
                                    cl.dob, ad.STREET1, ad.STREET2, ad.SUBURB, ad.COUNTRY, ad.POST, ad.STATE, 
                                    cl.PHONEHOME, cl.PHONEWORK, cl.PHONEMOBILE, ap.ADMISSIONYEAR, 
                                    p.PREFNUM, p.PREF, p.MOS, p.CAT, p.TER, p.PREFINSDATE, p.setinsdate, cl.email, ec.comments
                              FROM TISCADMIN.CLIENT cl join TISCADMIN.APPLICATION ap on cl.IDN = ap.CLIENTID and ap.IND_WD IS NULL
                                   left join tiscadmin.lastcommentecu ec on cl.idn = ec.clientid
                                   left join TISCADMIN.ADDRESS ad on cl.IDN = ad.CLIENTID and ad.addr_type = 'N' and ad.ISCURRENT = 'Y'
                                   join TISCADMIN.PREFERENCES p on ap.IDN = p.APPLICATIONID
                              WHERE cl.processyear = TISCCommon.getProcessYear and p.ISCURRENT = 'Y' and p.PREF in ('EAMUC', 'EACPV', 'EAMJC', 'EAMCC', 'EACMC', 'EAMAC', 'EAAMC', 'EAPAP') 
                                    and ((to_char(p.prefinsdate,'yyyymmdd') >= (select to_char(sc.ecu_pref_date1,'yyyymmdd') from tiscadmin.system_config sc)
                                        or to_char(p.setinsdate,'yyyymmdd') >= (select to_char(sc.ecu_pref_date1,'yyyymmdd') from tiscadmin.system_config sc))
                                    and (to_char(p.prefinsdate,'yyyymmdd') < (select to_char(sysdate,'yyyymmdd') from dual)
                                    and to_char(p.setinsdate,'yyyymmdd') < (select to_char(sysdate,'yyyymmdd') from dual)))
                              ORDER BY cl.PROCESSYEAR ASC, p.prefinsdate, p.setinsdate, p.PREF ASC, p.prefnum asc        
                              )
        LOOP
            utl_file.put_line(v_file, '"'||ecu_pref_row.PROCESSYEAR||'","'||ecu_pref_row.CLIENTNUMBER||'","'||ecu_pref_row.TITLE||'","'||ecu_pref_row.NAME_LAST||'","'||ecu_pref_row.NAME_FIRST||'","'||ecu_pref_row.NAME_SEC
                                         ||'","'||ecu_pref_row.dob||'","'||ecu_pref_row.STREET1||'","'||ecu_pref_row.STREET2||'","'||ecu_pref_row.SUBURB||'","'||ecu_pref_row.COUNTRY||'","'||ecu_pref_row.POST||'","'||ecu_pref_row.STATE||'","'||
                                         ecu_pref_row.PHONEHOME||'","'||ecu_pref_row.PHONEWORK||'","'||ecu_pref_row.PHONEMOBILE||'","'||ecu_pref_row.ADMISSIONYEAR||'","'||
                                         ecu_pref_row.PREFNUM||'","'||ecu_pref_row.PREF||'","'||ecu_pref_row.MOS||'","'||ecu_pref_row.CAT||'","'||ecu_pref_row.TER||'","'||ecu_pref_row.PREFINSDATE||'","'||ecu_pref_row.setinsdate||'","'||ecu_pref_row.email
                                         ||'","'||TISCCommon.replace_linebreak(ecu_pref_row.comments,' ')||'"');                    
        END LOOP;
        
        --UPDATE tiscadmin.system_config sc set sc.ecu_pref_date1 = (select sysdate from dual);
        --COMMIT;  
        
        utl_file.fclose(v_file);  
    END;
   */
   
      
  ---------------------------------------------------------------------------------------------------  
  procedure uwa_uu015_export_xml (p_processyear IN INT := 0) IS
    BEGIN
        IF (p_processyear > 0) THEN
            v_processyear := p_processyear;  
        ELSE
            v_processyear := TISCCOMMON.getProcessYear;
        END IF;
        
        SYSTEM.EXCEL_GENERATOR_XML.create_excel('OUT_TISC', 'uu015.xls');
        
        SYSTEM.EXCEL_GENERATOR_XML.create_style( 'h' , 'Arial', 'black',10, true);
        SYSTEM.EXCEL_GENERATOR_XML.create_style( 'c' , 'Arial', 'black',10, false);
        
        
        --------------------- worksheet - Y12students ----------------------  
        SYSTEM.EXCEL_GENERATOR_XML.create_worksheet( 'Y12students');
        
        r := 1;
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 1, 'Y12students', 'PROCESSYEAR', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 2, 'Y12students', 'CLIENTNUMBER', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 3, 'Y12students', 'IND_TYPE', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 4, 'Y12students', 'GENDER', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 5, 'Y12students', 'DOB', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 6, 'Y12students', 'PROVIDERSCHOOL', 'h' );
        
        FOR rec IN  (SELECT cl.processyear, cl.clientnumber, cl.ind_type, cl.gender, cl.dob, st.providerschool FROM CLIENT cl 
                            JOIN STUDENT st ON cl.idn = st.clientid AND cl.processyear = v_processyear
                                 WHERE cl.ind_type in ('1','2') AND nvl(st.status_cc,'!') != '2' ORDER BY cl.clientnumber) LOOP
                                 
            r := r+1;
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_num( r, 1, 'Y12students', rec.processyear, 'c'  );
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_num( r, 2, 'Y12students', rec.clientnumber, 'c' );
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_num( r, 3, 'Y12students', rec.ind_type, 'c');
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 4, 'Y12students', rec.gender, 'c');
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 5, 'Y12students', TO_CHAR(rec.dob, 'dd/mm/yyyy'), 'c');
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_num( r, 6, 'Y12students', rec.providerschool, 'c');
            
        END LOOP;
        DBMS_OUTPUT.PUT_LINE('total row number of Y12students: ' || r);
        
        --------------------- worksheet - CUR TEA ---------------------- 
        SYSTEM.EXCEL_GENERATOR_XML.create_worksheet( 'CUR TEA');
        
        r := 1;
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 1, 'CUR TEA', 'CLIENTNUMBER', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 2, 'CUR TEA', 'BEST_PASTYR', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 3, 'CUR TEA', 'TEACAT', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 4, 'CUR TEA', 'INSTITUTIONID', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 5, 'CUR TEA', 'TEATYPE', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 6, 'CUR TEA', 'AGGR', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 7, 'CUR TEA', 'TER', 'h' );
        
        FOR rec IN  (SELECT cl.clientnumber, pb.best_pastyr, tb.teacat, tb.institutionid, t.teatype, t.aggr, t.ter FROM CLIENT cl 
                            JOIN STUDENT st ON cl.idn = st.clientid AND cl.processyear = v_processyear
                            JOIN pep_base pb ON st.idn = pb.studentid AND pb.institutionid = '1'
                            JOIN tea_base tb ON st.idn = tb.studentid AND tb.institutionid = '1'
                            JOIN tea t ON tb.best_teaid = t.idn
                                 WHERE cl.ind_type in ('1','2') AND nvl(st.status_cc,'!') != '2' ORDER BY cl.clientnumber) LOOP
                                 
            r := r+1;
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_num( r, 1, 'CUR TEA', rec.clientnumber, 'c'  );
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_num( r, 2, 'CUR TEA', rec.best_pastyr, 'c' );
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 3, 'CUR TEA', rec.teacat, 'c');
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_num( r, 4, 'CUR TEA', rec.institutionid, 'c');
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 5, 'CUR TEA', rec.teatype, 'c');
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_num( r, 6, 'CUR TEA', rec.aggr, 'c');
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_num( r, 7, 'CUR TEA', rec.ter, 'c');
            
        END LOOP;
        DBMS_OUTPUT.PUT_LINE('total row number of CUR TEA: ' || r);
        
        --------------------- worksheet - UWA TEA ---------------------- 
        SYSTEM.EXCEL_GENERATOR_XML.create_worksheet( 'UWA TEA');
        
        r := 1;
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 1, 'UWA TEA', 'CLIENTNUMBER', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 2, 'UWA TEA', 'BEST_PASTYR', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 3, 'UWA TEA', 'TEACAT', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 4, 'UWA TEA', 'INSTITUTIONID', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 5, 'UWA TEA', 'TEATYPE', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 6, 'UWA TEA', 'AGGR', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 7, 'UWA TEA', 'TER', 'h' );
        
        FOR rec IN  (SELECT cl.clientnumber, pb.best_pastyr, tb.teacat, tb.institutionid, t.teatype, t.aggr, t.ter FROM CLIENT cl 
                            JOIN STUDENT st ON cl.idn = st.clientid AND cl.processyear = v_processyear
                            JOIN pep_base pb ON st.idn = pb.studentid AND pb.institutionid = '4'
                            JOIN tea_base tb ON st.idn = tb.studentid AND tb.institutionid = '4'
                            JOIN tea t ON tb.best_teaid = t.idn
                                 WHERE cl.ind_type in ('1','2') AND nvl(st.status_cc,'!') != '2' ORDER BY cl.clientnumber) LOOP
                                 
            r := r+1;
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_num( r, 1, 'UWA TEA', rec.clientnumber, 'c'  );
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_num( r, 2, 'UWA TEA', rec.best_pastyr, 'c' );
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 3, 'UWA TEA', rec.teacat, 'c');
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_num( r, 4, 'UWA TEA', rec.institutionid, 'c');
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 5, 'UWA TEA', rec.teatype, 'c');
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_num( r, 6, 'UWA TEA', rec.aggr, 'c');
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_num( r, 7, 'UWA TEA', rec.ter, 'c');
            
        END LOOP;
        DBMS_OUTPUT.PUT_LINE('total row number of UWA TEA: ' || r);
        
        --------------------- worksheet - Orig Prefs ---------------------- 
        SYSTEM.EXCEL_GENERATOR_XML.create_worksheet( 'Orig Prefs');
        
        r := 1;
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 1, 'Orig Prefs', 'ADMISSIONYEAR', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 2, 'Orig Prefs', 'CLIENTNUMBER', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 3, 'Orig Prefs', 'SETNUM', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 4, 'Orig Prefs', 'PREFNUM', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 5, 'Orig Prefs', 'PREF', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 6, 'Orig Prefs', 'SETINSDATE', 'h' );
        
        FOR rec IN  (SELECT a.admissionyear, cl.clientnumber, p.setnum, p.prefnum, p.pref, p.setinsdate FROM CLIENT cl 
                            JOIN application a ON cl.idn = a.clientid AND cl.processyear = v_processyear
                            JOIN preferences p ON a.idn = p.applicationid
                            JOIN STUDENT st ON cl.idn = st.clientid AND nvl(st.status_cc,'!') != '2' 
                                 WHERE cl.ind_type = '2' AND p.setnum = 0 ORDER BY cl.clientnumber, p.prefnum) LOOP
                                 
            r := r+1;
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_num( r, 1, 'Orig Prefs', rec.admissionyear, 'c'  );
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_num( r, 2, 'Orig Prefs', rec.clientnumber, 'c' );
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_num( r, 3, 'Orig Prefs', rec.setnum, 'c');
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_num( r, 4, 'Orig Prefs', rec.prefnum, 'c');
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 5, 'Orig Prefs', rec.pref, 'c');
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 6, 'Orig Prefs', TO_CHAR(rec.setinsdate, 'dd/mm/yyyy HH24:MI:SS'), 'c');
            
        END LOOP;
        DBMS_OUTPUT.PUT_LINE('total row number of Orig Prefs: ' || r);
        
        
        --------------------- worksheet - Curr Prefs ---------------------- 
        SYSTEM.EXCEL_GENERATOR_XML.create_worksheet( 'Curr Prefs');
        
        r := 1;
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 1, 'Curr Prefs', 'ADMISSIONYEAR', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 2, 'Curr Prefs', 'CLIENTNUMBER', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 3, 'Curr Prefs', 'SETNUM', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 4, 'Curr Prefs', 'ISCURRENT', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 5, 'Curr Prefs', 'PREFNUM', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 6, 'Curr Prefs', 'PREF', 'h' );
        SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 7, 'Curr Prefs', 'SETINSDATE', 'h' );
        
        FOR rec IN  (SELECT a.admissionyear, cl.clientnumber, p.setnum, p.iscurrent, p.prefnum, p.pref, p.setinsdate FROM CLIENT cl 
                            JOIN application a ON cl.idn = a.clientid AND cl.processyear = v_processyear
                            JOIN preferences p ON a.idn = p.applicationid
                            JOIN STUDENT st ON cl.idn = st.clientid AND nvl(st.status_cc,'!') != '2' 
                                 WHERE cl.ind_type = '2' AND p.iscurrent = 'Y' ORDER BY cl.clientnumber, p.prefnum) LOOP
                                 
            r := r+1;
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_num( r, 1, 'Curr Prefs', rec.admissionyear, 'c'  );
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_num( r, 2, 'Curr Prefs', rec.clientnumber, 'c' );
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_num( r, 3, 'Curr Prefs', rec.setnum, 'c');
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 4, 'Curr Prefs', rec.iscurrent, 'c');
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_num( r, 5, 'Curr Prefs', rec.prefnum, 'c');
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 6, 'Curr Prefs', rec.pref, 'c');
            SYSTEM.EXCEL_GENERATOR_XML.write_cell_char( r, 7, 'Curr Prefs', TO_CHAR(rec.setinsdate, 'dd/mm/yyyy HH24:MI:SS'), 'c');
            
        END LOOP;
        DBMS_OUTPUT.PUT_LINE('total row number of Curr Prefs: ' || r);
        r := 0;
        
        --------------------flush data into file, and close file ------------------------
        SYSTEM.EXCEL_GENERATOR_XML.close_file;            
    END;
    
end DATA_EXPORT;
/
