create or replace FUNCTION         J_ISR_UDF_VAL_F(v_ISR_NO varchar2,
                                                    v_Field_text varchar2,
                                                    v_agency_code varchar2, 
                                                    v_service_code varchar2, 
                                                    v_pfilter1 varchar2,
                                                    v_svc_type varchar2)

RETURN varchar2
IS

/*  Please leave this and the next three lines as is
    Standard Header for all JEA FIS Programs
 ******************************************************************************
 ******************************************************************************
 File:  J_ISR_UDF_VAL_F.sql
 Title: Retrieve single value from JEACUST customer agency ISR tables
        based on field label, agency_code, service_code and pfilter1.
 Invocation: Called from reports as needed.

 Description: 

 Schema Owner: CAPR
 Parameters:    

 Frequency: 

 Output: N/A
 Copyright (c) 1997 JEA, All rights reserved.
                    No part of the contents of this program may be
                    reproduced in any form or by any means without the
                    written permission of JEA
 Source Safe Params
 $Header:                                $
 $Author:                                $
 $Log   :                                $
 $Modtime:                               $
 $Revision :                             $
  -------------------------- Revision History --------------------------
  Programmer  TR #  Date         Description
  ==========  ====  ===========  ===========================================
T.Whitehead         03/21/2017  First write for UDF field value retrieval
J.Acosta            12/29/2017  Added rownum < 2 filters to all calls to jeacust.j_scr2udf_def table to prevent dup's from causing TOO_MANY_ROWS exceptions
T.Whitehead         01/11/2018  Add REHAB, S CAPITAL, W CAPITAL to the Mod_datetime case statement
T.Whitehead         01/12/2018  Corrected logic flow to obtain values for SOCC filter1 for METER Jobs.
                                Corrected logic flow of W-WW and ELEC to remove if stmt checking the service
                                code as meter job.  Configuration change prior to pristine so that
                                these jobs will continue to be Meter Agency even though worked by 
                                the W-WW or ELEC agency.

 *********** Please leave this and the next line as is ************************
 ******************************************************************************/
 -- the tables that share p_filter1 codes don't use same labels/field_text.  
 --Even though there are multiple scr_tags, SQL is build to only obtain the
 --first retrieved row of the field_name, scr_tag, table_name.  
 
v_UDF_VALUE varchar2(200);
get_table_name varchar2(50);
his_table_name varchar2(50);
get_field_val varchar2(200);
get_field_val_date TIMESTAMP with time zone;
get_desc varchar2(100);
rec_count number(9):= 0;
v_table_name varchar2(30);
v_field_name varchar2(50);
v_field_type varchar2(20);

v_scr_tag2 varchar2(50);
v_sql varchar2(1000);
v_sql_prep varchar2(1000);
VALID_AGENCY CHAR(1);
v_tab_field varchar2(50);
v_scr_tag_items varchar2(500);
v_vtb_name varchar2(20);
CURSOR c_scr_tag is
  select distinct scr_tag_template scr_tag from jeacust.j_filter1_template
  where p_filter1 = v_pfilter1;

      begin 

   VALID_AGENCY := 'F';
   v_tab_field := 'a';
   -- v2_scr_tag := v_scr_tag;
    if v_agency_code = 'METER' then 
          VALID_AGENCY := 'T';
          if v_field_text  = 'MOD_DATETIME' then
          -- many of the reports used the J_Completion_step mod_date and time to determine a status of closed.
          -- The J_MTR-ISR* tables replaced the j_completion step.  Will be using the mod_datetime from the
          --Completion container to process the date..  That container is determined by p_filter1.  For this
          --function the container is referred to as v_scr_tag2 and is stored as form_name in the
          -- j_MTR-ISR* tables.
              case when v_pfilter1 = 'MTR-EA' then
                       v_table_name := 'J_MTR_ISR1';
                       v_field_name := 'MOD_DATETIME';
                       v_field_type := 'Date';
                       v_scr_tag2 := 'EA_APPT1';
                    when v_pfilter1 in ('MTR-CD','MTR-CC','MTR-TP','MTR-RS','MTR-RCD','MTR-PP') then
                       v_table_name := 'J_MTR_ISR1';
                       v_field_name := 'MOD_DATETIME';
                       v_field_type := 'Date';
                       v_scr_tag2 := 'MTR_ELECTRIC1';
                    when v_pfilter1 in ('MTR-MR') then
                       v_table_name := 'J_MTR_ISR1';
                       v_field_name := 'MOD_DATETIME';
                       v_field_type := 'Date';
                       v_scr_tag2 := 'MTR_MR';                       
                    WHEN v_pfilter1 = 'MTR-CE' then
                       v_table_name := 'J_MTR_ISR10';
                       v_field_name := 'MOD_DATETIME';
                       v_field_type := 'Date';
                       v_scr_tag2 := 'MTR_COMMELECNODE2';
                    WHEN v_pfilter1 = 'SOCC'  and v_svc_type = 'E' then
                       v_table_name := 'J_MTR_ISR1';
                       v_field_name := 'MOD_DATETIME';
                       v_field_type := 'Date';
                       v_scr_tag2 := 'MTR_POLE1';                       
                    WHEN  v_pfilter1 = 'SOCC' AND v_svc_type = 'W' then
                        v_table_name := 'J_MTR_ISR6';
                         v_field_name := 'MOD_DATETIME';
                         v_field_type := 'Date';
                         v_scr_tag2 := 'WATER COMPLETION';    
                    WHEN  v_pfilter1 in ('MTR-WM','MTR-CW','MTR-RW','WATER-TAP','SEWERMAINT','REHAB','S CAPITAL','W CAPITAL') then
                        v_table_name := 'J_MTR_ISR6';
                         v_field_name := 'MOD_DATETIME';
                         v_field_type := 'Date';
                         v_scr_tag2 := 'WATER COMPLETION';                         
                    else
                       v_udf_value := '01/01/1900 00:00:01';
                       v_tab_field := 'ERROR';
              end case;
                        
          else --else part of if v_field_text = 'MOD_DATETIME'
              -- SOCC water and electric will be worked through METER because they are coming through CCnB.
              --Following if statement is how/where the ISR will be processed.  Since SOCC can be
              --either water electric, needed to used service type to determine the scr_tag.
              IF V_pfilter1 = 'SOCC' then
                  if v_svc_type = 'E' then
                        v_scr_tag_items := '(''MTR_POLE1'')';
                    else --IF v_svc_type = 'W' then
                       v_scr_tag_items := '(''WATER COMPLETION'',''WATER GENERAL INFO'',''WATER RECLAIM TEST'', ''WATER TEST RESULTS'', ''WATER READ-NEW METER'')';
                  end if;
              ELSE  --All other p_filter1 use the following loop to obtain possible scr_tags
                    v_scr_tag_items := '(';
                    for s in c_scr_tag loop
                      rec_count := rec_count +1;
                      if rec_count > 1 then
                        v_scr_tag_items := v_scr_tag_items || ',';
                      end if;
                      v_scr_tag_items := v_scr_tag_items ||''''|| s.scr_tag||'''';
                    end loop;
                    v_scr_tag_items := v_scr_tag_items || ')';
               END IF;
              --the following is ued to get the field name for the service_code, agency_code and report_text_field combination.
              --Once this is executed, there should only be one field in this combination that gives the table_name, field_name (UDF...) 
              --and field_type and specific scr_tag.  This is used then to retrieve the
              --data value in the v_Sql build below.
            v_sql_prep :=  'SELECT  distinct table_name, field_name, s_field_type,scr_tag  ' 
                ||'  from jeacust.j_scr2udf_def ' 
                || '  where scr_tag in '|| v_scr_tag_items 
                --|| ' and service_code = ''' || v_service_code 
                || ' and agency_code = ''' || v_agency_code
                || ''' and report_field_text =  ''' ||v_field_text
                || ''' and rownum < 2';

            BEGIN
              --If the following produces an exception, then it will skip the logic
              --to obtain the udf field value.  In this event 'ERROR' will be 
              --returned to the calling SQL.
                EXECUTE IMMEDIATE V_SQL_PREP INTO v_table_name, v_field_name, v_field_type,v_scr_tag2;
                EXCEPTION
                       WHEN NO_DATA_FOUND THEN 
                            v_tab_field := 'ERROR'; 
                       WHEN TOO_MANY_ROWS THEN 
                            v_tab_field := 'ERROR';       
                       WHEN OTHERS THEN
                            -- v_tab_field := 'Other error ' || SQLCODE;
                            v_tab_field := 'ERROR';
       
            END;
          end if;
        
      -- elsif is second part (looking for agency_code of 'W-WW') of if v_agency_code = 'METER
    elsif (v_agency_code = 'W-WW' or v_agency_code = 'SWR' ) then
            VALID_AGENCY := 'T';
            --  the meter configuration for the following job codes was changed just prior to pristine.  They will be processed in the METER agency 
            -- Filter1 codes for these job types are mapped in the j_filter1_template table.
            /*
            if v_service_code in ('DYE TEST','SWRLOC', 'COMPSWR','M-CON SP','TERMINAT','INCREASE','INCREASR','DECREASE','DECREASR','M-CONINS') THEN
                --this means it is a meter job BUT in a 'W-WW' agency 
                  if v_field_text  = 'MOD_DATETIME' then
                        v_table_name := 'J_MTR_ISR6';
                       v_field_name := 'MOD_DATETIME';
                       v_field_type := 'Date';
                       v_scr_tag2 := 'WATER COMPLETION'; 
                       
                    else    
                        v_scr_tag_items := '(''WATER COMPLETION'',''WATER GENERAL INFO'',''WATER RECLAIM TEST'', ''WATER TEST RESULTS'', ''WATER READ-NEW METER'')';
                        
                    end if;
                     v_sql_prep :=  'SELECT  distinct table_name, field_name, s_field_type,scr_tag  ' 
                      ||'  from jeacust.j_scr2udf_def ' 
                      || '  where scr_tag in '|| v_scr_tag_items 
                      --|| ' and service_code = ''' || v_service_code 
                      || ''' and agency_code = ''' || v_agency_code
                      || ''' and report_field_text =  ''' ||v_field_text
                      || ''' and rownum < 2'; 
                   
                      Begin
                        --If the following produces an exception, then it will skip the logic
                        --to obtain the udf field value.  In this event 'ERROR' will be 
                        --returned to the calling SQL.
                          EXECUTE IMMEDIATE V_SQL_PREP INTO v_table_name, v_field_name, v_field_type,v_scr_tag2;
                          EXCEPTION
                                 WHEN NO_DATA_FOUND THEN 
                                      v_tab_field := 'ERROR'; 
                                 WHEN TOO_MANY_ROWS THEN 
                                      v_tab_field := 'ERROR';       
                                 WHEN OTHERS THEN
                                      -- v_tab_field := 'Other error ' || SQLCODE;
                                      v_tab_field := 'ERROR';
                 
                      end;     
                 
        else  --SERVICE TYPE FOR W-WW non-meter) will be used to indicate whether SOCC or Completion Data should be used.
        */
                v_sql_prep :=  'SELECT   table_name, field_name, s_field_type,scr_tag  ' 
                ||'  from jeacust.j_scr2udf_def ' 
                || '  where scr_tag LIKE ''%'|| v_svc_type || '%'''  
                || ' and service_code = ''' || v_service_code 
                || ''' and agency_code = ''' || v_agency_code
                || ''' and report_field_text =  ''' ||v_field_text
                || ''' and rownum < 2';
                 Begin
                  --If the following produces an exception, then it will skip the logic
                  --to obtain the udf field value.  In this event 'ERROR' will be 
                  --returned to the calling SQL.
                    EXECUTE IMMEDIATE V_SQL_PREP INTO v_table_name, v_field_name, v_field_type,v_scr_tag2;
                    EXCEPTION
                           WHEN NO_DATA_FOUND THEN 
                                v_tab_field := 'ERROR'; 
                           WHEN TOO_MANY_ROWS THEN 
                                v_tab_field := 'ERROR';       
                           WHEN OTHERS THEN
                                -- v_tab_field := 'Other error ' || SQLCODE;
                                v_tab_field := 'ERROR';
           
                end;
            -- v_udf_value := '01/01/1900 00:00:02';
           --  v_tab_field := 'ERROR';        
     --   end if; --end if service-code (in METER job types for W-WW)
    elsif v_agency_code = 'ELEC' then
            VALID_AGENCY := 'T';
            
            --many of the elec jobs are processed in the a single table/scr_tag and the report query
            --does not yet use this function.  If report query needs to start using function, may need
            --to update this section of code should there be any custom ELEC configuration
            --to be accomodated.
            
            --Since the following service codes will be worked through meter agency,
            --removed the code to capture data based on ELEC agency.
            
           if v_service_code in ('TDLPM','POLE-DIS','POLE-REC','POLE-SAF','TDL ONLY', 'PCRM') THEN
                --this means it is a meter job BUT in a 'ELEC agency 
                   IF v_field_text  = 'MOD_DATETIME' then
                       v_table_name := 'J_MTR_ISR1';
                       v_field_name := 'MOD_DATETIME';
                       v_field_type := 'Date';
                       v_scr_tag2 := 'MTR_ELECTRIC1';  
                       
                    ELSE    
                        v_scr_tag_items := '(''MTR_ELECTRIC1'')';
                    END IF;    
            end if;       
               v_sql_prep :=  'SELECT   table_name, field_name, s_field_type,scr_tag  ' 
                ||'  from jeacust.j_scr2udf_def ' 
                || '  where scr_tag in '|| v_scr_tag_items 
                || ' and service_code = ''' || v_service_code 
                || ''' and agency_code = ''' || v_agency_code
                || ''' and report_field_text =  ''' ||v_field_text
                || ''' and rownum < 2';
            BEGIN
              --If the following produces an exception, then it will skip the logic
              --to obtain the udf field value.  In this event 'ERROR' will be 
              --returned to the calling SQL.
                EXECUTE IMMEDIATE V_SQL_PREP INTO v_table_name, v_field_name, v_field_type,v_scr_tag2;
                EXCEPTION
                       WHEN NO_DATA_FOUND THEN 
                            v_tab_field := 'ERROR'; 
                       WHEN TOO_MANY_ROWS THEN 
                            v_tab_field := 'ERROR';       
                       WHEN OTHERS THEN
                            -- v_tab_field := 'Other error ' || SQLCODE;
                            v_tab_field := 'ERROR';
       
            END;    
             
      --  else
      --       v_udf_value := '01/01/1900 00:00:03';
     --        v_tab_field := 'ERROR';                
     --   end if; --end if service-code (in ELEC job types)

      else  --  third part of if v_agency_code = 'METER'
      --Agency not found will be returned to the calling sql.
      --This should not occur since the agency_code is coming from ISR table.
          VALID_AGENCY := 'F';
          v_udf_value := 'AGENCY NOT FOUND';
      end if;  --end if  v_agency_code_ 'METER
      
      
      -- Retrieve field value for table_name, scr_tag, and field_name.  THis will take into account that ISR data may have been
      -- disposed and moved to History.
      IF VALID_AGENCY = 'T' and nvl(v_tab_field,' ') <> 'ERROR'  THEN
        get_table_name := 'JEACUST.'||v_table_name;
        his_table_name := 'JEACUST.HIS_'||v_table_name;
        v_sql :=  'SELECT ' ||  v_field_name  
                  ||'  from ' || get_table_name 
                  || '  where isr_no = ''' || V_Isr_no 
                  || ''' and FORM_NAME = ''' || v_scr_tag2
                  || ''' and  node_sequence = (select max(node_sequence) from ' || get_table_name || ' where isr_no = ''' || V_Isr_no ||''''||')' 
                  || ' UNION '
                  || ' SELECT ' ||  v_field_name  
                  || '  from ' || his_table_name 
                  || '  where isr_no = ''' || V_Isr_no 
                  || ''' and FORM_NAME = ''' || v_scr_tag2
                  || ''' and  node_sequence = (select max(node_sequence) from ' || his_table_name || ' where isr_no = ''' || V_Isr_no ||''''||')' ;

         if v_field_type = 'Date' then 
                 BEGIN --executes the sql built above. This logic is for a Date (Mod_datetime) value.  In order to use a date comparison, the calling sql
                 --will need to convert the text date into date.
      
                      EXECUTE IMMEDIATE v_sql into get_field_val_date;
      
                    v_udf_value := to_char(get_field_val_date at local,'MM/DD/YYYY HH24:MI:SS');
      
            --returned date value will indicate which error is received.
                   EXCEPTION
                         WHEN NO_DATA_FOUND THEN 
                        v_udf_value := '01/01/1900 00:00:04';
                         WHEN TOO_MANY_ROWS THEN 
                              v_udf_value := '01/01/1900 00:00:05';
                         WHEN OTHERS THEN
                               v_udf_value := '01/01/1900 00:00:06';
                    END;       
          else    
             
            BEGIN -- Executes the v_sql uild abovve for non-date values.
      
                 EXECUTE IMMEDIATE v_sql  into get_field_val;
              --Some fields on the screen containers are check boxes.  The DB stores these values as 0 (False) or 1(True) and are in a Number field.
              --In this event these values will be converted as follows.  Any other specific wording that is needed on the reports will need to occur
              --in the calling sql.
                 if  v_field_type = 'CHECK' then
                         if get_field_val = 1 then v_udf_value := 'True';
                         elsif get_field_val = 0 then v_udf_value := 'False';
                         else v_udf_value := 'False';
                         end if;
                 else 
                    v_udf_value := get_field_val; 
                 end if;

                   EXCEPTION
                     WHEN NO_DATA_FOUND THEN 
                          v_udf_value := '';
                     WHEN TOO_MANY_ROWS THEN 
                          v_udf_value := '';
                     WHEN OTHERS THEN
                          -- v_udf_value := 'GET VAL ' || SQLCODE;
                        v_udf_value := '';
                 END;
         end if; --end of  if v_field_type = 'Date'
            v_vtb_name := '';
            /*
            dbms_output.put_line('about to get Standard Comments');
            dbms_output.put_line('get_field_val' || get_field_val);
            dbms_output.put_line('v_field_text' || v_field_text);      
            dbms_output.put_line('v_udf_value' || v_udf_value);            
             */
             
             --for meter agency, the configuration uses a Standard Comment
             --that uses a code in the field instead of a full description. 
             --The description has to be obtained from the VTB table depending
             --on the the type of job completed.
        if v_Field_text like 'Standard Comments%' AND get_field_val is not null then
              IF v_scr_tag2 = 'EA_APPT1' THEN
                       V_VTB_NAME := 'STDCMT_EA';
              ELSif v_scr_tag2 = 'MTR_ELECTRIC1' then
                    CASE v_Field_text
                      when 'Standard Comments 1' then
                        v_vtb_name := 'STD_CMT1';
                      when 'Standard Comments 2' then
                        v_vtb_name := 'STD_CMT2' ;
                      when 'Standard Comments 3' then
                        v_vtb_name := 'STD_CMT3';
                      else
                        v_vtb_name := '';
                    end case;
              ELSIF v_scr_tag2 = 'MTR_MR' THEN
                  CASE v_Field_text
                      when 'Standard Comments 1' then
                        v_vtb_name := 'MTR_MR_CMN';
                      when 'Standard Comments 2' then
                        v_vtb_name := 'MTR_CMN2' ;
                      when 'Standard Comments 3' then
                        v_vtb_name := 'MTR_MR_CMN3';
                      else
                        v_vtb_name := '';
                    end case;      
              ELSIF v_scr_tag2 = 'WATER COMPLETION' THEN
                       V_VTB_NAME := 'STDW_CMT1';     
              ELSIF v_scr_tag2 = 'MTR_POLE1' THEN
                       V_VTB_NAME := 'STD_SOCC';   
              ELSIF v_scr_tag2 = 'MTR_COMMELECNODE2' THEN
                  CASE v_Field_text
                      when 'Standard Comments 1' then
                        v_vtb_name := 'STD_CMS1';
                      when 'Standard Comments 2' then
                        v_vtb_name := 'STD_CMS2' ;
                      else
                        v_vtb_name := '';
                    end case; 
              ELSIF v_scr_tag2 = 'MTR_HYDRANT1' THEN
                     V_VTB_NAME := 'MTR_HYDRANT';                   
              END IF;
                dbms_output.put_line('v_vtb_name1' || v_vtb_name);
                dbms_output.put_line('get_field_val1' ||get_field_val);  
        end if;
        if v_vtb_name is not null then
           -- dbms_output.put_line('about to get vtb_name');        
                  BEGIN
                      select distinct description into get_desc
                      from VTB_CODE
                      where Upper(VTB_NAME) = Upper(v_vtb_name)
                      and vtb_code = get_field_val;
                      get_field_val := get_desc;
                      dbms_output.put_line('get_field_val: ' || get_field_val);
                      dbms_output.put_line('get_desc: ' || get_desc);                      
                      EXCEPTION
                         WHEN NO_DATA_FOUND THEN
                         dbms_output.put_line('vtb_code not found');
                              get_desc := get_field_val;
                         WHEN TOO_MANY_ROWS THEN 
                              get_desc := get_field_val;
                         WHEN OTHERS THEN
                              get_desc := get_field_val;
                  END;
                  v_udf_value := get_desc;
        end if;  
      ELSE --not valid agency
        if nvl(v_field_name,' ') <> 'MOD_DATETIME' then      
          v_udf_value := '';
        end if;
      END IF; -- end if valid_agency
      return v_udf_value;
    end J_ISR_UDF_VAL_F;