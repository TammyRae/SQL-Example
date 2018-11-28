/*
	CUFOR144:Commercial Water Field Activities Completed
	Report Title: Completed Commercial Water FA's

*/

			/*  PragmaCad6.5 Upgrade                                                                                                     T.Whitehead 04/2016

  As part of the Upgrade, all reports were converted from CR App, Report Generator and Oracle Reports to 
  BI Publisher.
  Accomodate A. Table changes for ISR: Removal of UDF fields to ISR_UDF table which requires
                                an update to select (table alias for ISR_UDF), the FROM clause to include ISR_UDF and the WHERE clause 
                                  to includeISR.ISR_NO = ISR_UDF.ISR_NO.
                          B. All _date and _time pairs: Combined into _datetime as timestamp with timezone at 'UTC'.
                                If a _datetime is selected for display it is cast at local date.  If _datetime is being compared to 
                                SYSDATE or supplied input date, the date is cast to local.  Internal compares (table to table) within FMS
                                are not cast to local.  Datetime compares to OMS and CCnB will require a cast to local.
                          C. in Scheduled reports such as FMS/EXcel and Oracle Reports, the run date defaulted to SYSDATE-1 or is
                               Hardcoded.  This was changed to allow for an input date but if the input date is null not supplied, use a default date
                               as was coded in Original Program/XML.
                          D. FMSRPT_CAPR11_L and FMSRPT_OMPR11_l were changed to use new link.  CAD reports will 
                              use BIOMSRPT_OMPR11_L as dblink to OMS.  OMS reports will use BICADRPT_CAPR11_L as dblink to CAD.
                          E.  Add History.  Additional Table changes that allows for CAD archiving.  Requires that the SQL be duplicated
                               and changed to reflect HIS_ tables.  Will include OMS reports where necessary.
                               
  CAD reports are to exclusively run in the BI_CAD_RPT user.   OMS reports are to exclusively run in the BI_OMS_RPT user.
                          
                               
*/
/*
  -------------------------- Revision History --------------------------
  Programmer  TR #  Date         Description
  ==========  ====  ===========  ===========================================
T.Whitehead         04/14/2017   Function call update for Customized meter tables.

 *********** Please leave this and the next line as is ************************
 ******************************************************************************/
 
select

cur_cust_name,
fa_id,
to_char(to_date(decode(fa_schd_date,null,'01011900',
      trim(to_char(to_number(substr(fa_schd_date,1,instr(fa_schd_date, '/') - 1)),'00')) ||
      trim(to_char(to_number(substr(fa_schd_date,instr(fa_schd_date, '/')+1,instr(fa_schd_date,'/',1,2) - (instr(fa_schd_date,'/')+1))),'00')) ||
      trim(substr(fa_schd_date,instr(fa_schd_date,'/',1,2)+1,4))),'MMDDYYYY'),'MM/DD/YY') fa_schd_date, 
address,
zip_code,
sp_type_code,
clerk_id,
terr_code,
decode(trim(status_srv_on_arrival)||trim(status_srv_off_arrival),'TrueFalse','ON','FalseTrue','OFF','') status_srv_on_arrival,
decode(trim(status_srv_on_depart)||trim(status_srv_off_depart),'TrueFalse','ON','FalseTrue','OFF','') status_srv_on_depart,
fa_type_code,
work_done,
standard_comments1, 
 standard_comments2, 
standard_comments3,            
im_fa_comments,  
instructions,
nm_sp_water_tap_size, 
nm_sp_water_serv_size,
nm_sp_sewer_tap_size,
nm_meter_loc_cd,
nm_sp_water_tap_loc,

sp_water_tap_size,
sp_water_serv_size,
sp_sewer_tap_size,
meter_loc_cd,
sp_water_tap_loc,

nm_meter_number,
nm_meter_manu,  
nm_meter_size,         
nm_meter_type, 
nm_meter_loc_detail,
 nm_current_reading1,
nm_current_reading1_uom, 
nm_current_reading1_digits_l,
new_rf_id1,
nm_current_reading1_tou,
new_rid_id1,
nm_current_reading2,
nm_current_reading2_uom, 
nm_current_reading2_digits_l,
new_rf_id2,
nm_current_reading2_tou,
new_rid_id2,
nm_current_reading3,          
nm_current_reading3_uom, 
nm_current_reading3_digits_l,
 nm_current_reading3_rf_id,	
nm_current_reading3_tou,

om_meter_number,
meter_manu,
meter_size,
 meter_digits_left,
meter_type,
meter_loc_details,
om_current_reading1,
om_current_reading1_uom,
om_current_reading1_digits_l,
old_rfw_1, 
om_current_reading1_tou,
om_current_reading1_rid,
old_rid_1, 
om_current_reading2,
om_current_reading2_uom,
om_current_reading2_digits_l,
old_rfw_2,
om_current_reading2_tou,
om_current_reading2_rid,
old_rid_2,  

new_test_comments,
prev_test_comments,
case when test_passed = 'True' then 'Passed'
     when test_failed = 'True' then 'Failed'
     when no_test = 'True' then 'No Test'
     else 'unknown'
end  pass_fail,
	-- PRE TEST ---------------
new_high_flow, --test_result1
 new_medium_flow,
new_low_flow,
new_average_flow,
new_gpm, --test_result5
new_psi,

	-- FINAL TEST --------------
  /*
new_reading_before_1,
new_reading_before_2, --test_result8
new_reading_before_3,
 new_reading_after_1,
new_reading_after_2,
new_reading_after_3, --test_result12
*/
prev_test_date,
	-- PREVIUOS TEST -----------
prev_high_flow,
prev_medium_flow,
prev_low_flow,
prev_average_flow,
prev_gpm,
prev_psi,
im_fa_status_date,
tech_assigned, 
isr_no
from (

(
SELECT DISTINCT 
	a.compl_name cur_cust_name,
	u.udf24 fa_id,
 	JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'FA Schedule Date',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') FA_SCHD_DATE,
	a.compl_address address,
	a.zip_code,
	u.udf15 sp_type_code,
JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Clerk ID',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') Clerk_id,      
	--u.udf17 clerk_id,
	null terr_code,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Service on Departure ON',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') status_srv_on_depart,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Service on Arrival ON',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') STATUS_SRV_ON_arrival,
        
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Service on Departure OFF',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') status_srv_OFF_depart,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Service on Arrival OFF',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') STATUS_SRV_OFF_arrival,        
	a.init_service_code fa_type_code,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Work Done',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') work_done,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Standard Comments 1',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') standard_comments1, 
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Standard Comments 2',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') standard_comments2, 
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Standard Comments 3',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') standard_comments3,            
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Completion Remarks',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') im_fa_comments,  
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Clerk Instructions',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') instructions,

  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Water Tap Size',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_sp_water_tap_size, 
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Service Size',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_sp_water_serv_size,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Sewer Tap Size',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_sp_sewer_tap_size,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Meter Loc Code',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_meter_loc_cd,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Water Tap Location',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_sp_water_tap_loc,

	a3.sp_water_tap_size,
	a3.sp_water_serv_size,
	a3.sp_sewer_tap_size,
	a.p_loc_type AS meter_loc_cd,
	a3.sp_water_tap_loc,

  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'New Meter Number',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_meter_number,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'New Meter Make',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_meter_manu,  
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'New Meter Size',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_meter_size,         
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'New Meter Model Type',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '')  nm_meter_type, 
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Meter Loc Details',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_meter_loc_detail,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'New Read A',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_current_reading1,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'New UOM A',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_current_reading1_uom, 
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'New Digits A',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_current_reading1_digits_l,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'New RF A',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') new_rf_id1,
null nm_current_reading1_tou,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'RID Installed A',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') new_rid_id1,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'New Read B',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_current_reading2,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'New UOM B',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_current_reading2_uom, 
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'New Digits B',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_current_reading2_digits_l,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'New RF B',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') new_rf_id2,
null nm_current_reading2_tou,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'RID Installed B',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') new_rid_id2,
	null nm_current_reading3,          
	null nm_current_reading3_uom, 
	null nm_current_reading3_digits_l,
	null nm_current_reading3_rf_id,	
	null nm_current_reading3_tou,

--	u.udf3 om_meter_number,

JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Old Meter Number',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') om_meter_number,
	a3.manufacturer_code meter_manu,
	a3.meter_size,
	(SELECT jea_register.digits_left FROM jea_register WHERE isr_no = a.isr_no AND ROWNUM < 2) AS meter_digits_left,
	a3.type meter_type,
	a3.location_details meter_loc_details,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Read Today A',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') om_current_reading1,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'UOM A',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') om_current_reading1_uom,
	(SELECT digits_left FROM jea_register WHERE a.isr_no = isr_no AND read_seq = '1' AND ROWNUM < 2) AS om_current_reading1_digits_l,
	(SELECT y.badge_number FROM jea_item y WHERE y.isr_no = a.isr_no AND y.type_code = 'RFW' AND ROWNUM < 2) old_rfw_1, 
	(SELECT time_of_use FROM jea_register WHERE a.isr_no = isr_no AND read_seq = '1' AND ROWNUM < 2) AS om_current_reading1_tou,
	(SELECT register_id FROM jea_register WHERE a.isr_no = isr_no AND read_seq = '1' AND ROWNUM < 2) AS om_current_reading1_rid,
	(SELECT y.badge_number FROM jea_item y WHERE y.isr_no = a.isr_no AND y.type_code = 'RID' AND y.comments not LIKE 'B-BIG%' AND ROWNUM < 2) old_rid_1, 
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Read Today B',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '')om_current_reading2,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'UOM B',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') om_current_reading2_uom,
	(SELECT digits_left FROM jea_register WHERE a.isr_no = isr_no AND read_seq = '2' AND ROWNUM < 2) AS om_current_reading2_digits_l,
	(SELECT y.badge_number FROM jea_item y WHERE y.isr_no = a.isr_no AND y.type_code = 'RFW' AND ROWNUM < 2 
        AND y.badge_number <> (SELECT y.badge_number FROM jea_item y WHERE y.isr_no = a.isr_no AND y.type_code = 'RFW' AND ROWNUM < 2)) old_rfw_2,
	(SELECT time_of_use FROM jea_register WHERE a.isr_no = isr_no AND read_seq = '2' AND ROWNUM < 2) AS om_current_reading2_tou,
	(SELECT register_id FROM jea_register WHERE a.isr_no = isr_no AND read_seq = '2' AND ROWNUM < 2) AS om_current_reading2_rid,
	(SELECT y.badge_number FROM jea_item y WHERE y.isr_no = a.isr_no AND y.type_code = 'RID' AND y.comments LIKE 'B-BIG%' AND ROWNUM < 2) old_rid_2,  

  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Test Comments',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') new_test_comments,
	e.test_comments prev_test_comments,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Test Passed',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') test_passed,
JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Test Failed',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') test_failed,
JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'No Test',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') no_test,        
	-- PRE TEST ---------------
JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'High Flow 1',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') new_high_flow, --test_result1
JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Medium Flow 1',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') new_medium_flow,
JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Low Flow 1',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') new_low_flow,
JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Average 1',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') new_average_flow,
JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'GPM 1',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') new_gpm, --test_result5
JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'PSI 1',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') new_psi,

	-- FINAL TEST --------------
  /*
	b.test_result7 new_reading_before_1,
	b.test_result8 new_reading_before_2, --test_result8
	b.test_result9 new_reading_before_3,
	b.test_result10 new_reading_after_1,
	b.test_result11 new_reading_after_2,
	b.test_result12 new_reading_after_3, --test_result12
*/
	TO_CHAR(TO_DATE(TO_NUMBER(DECODE(e.test_date,' ', '19000101',e.test_date)),'YYYYMMDD'), 'MM/DD/YYYY') prev_test_date,
	-- PREVIUOS TEST -----------
	e.high_flow prev_high_flow,
	e.medium_flow prev_medium_flow,
	e.low_flow prev_low_flow,
	e.average_flow prev_average_flow,
	e.gpm prev_gpm,
	e.psi prev_psi,

JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'MOD_DATETIME',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') as im_fa_status_date,
	a.handling_unit tech_assigned, 
  a.isr_no

FROM isr a, 
     jeacust.jea_meter a3,          
     jeacust.jea_watermetertesthistory e,
     isr_udf u

WHERE a.isr_no = a3.isr_no (+)
	AND a.isr_no = e.isr_no (+)
  AND a.isr_no = u.isr_no (+)

	AND a.agency_code = 'METER'
	AND (a.p_status = 'CL' OR a.s_mode = 'L')
	AND a.p_filter1 = 'MTR-CW'
    and a.init_service_code not like '\_\_%' escape '\'
    

    and   nvl(JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Work Done',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') ,' ') <> 'RETURN TO DISPATCH'
  AND
      trunc(TO_DATE(JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'MOD_DATETIME',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        ''),'MM/DD/YYYY HH24:MI:SS')) BETWEEN 
            DECODE(:P_BEGIN_DATE,NULL,TRUNC(SYSDATE),TO_DATE(:P_BEGIN_DATE,'MM/DD/YYYY'))
  AND        
             DECODE(:P_END_DATE,NULL,trunc(sysdate),TO_DATE(:P_END_DATE,'MM/DD/YYYY'))


	AND (TRIM(e.test_date) IS NULL OR e.test_date = (SELECT MAX(y.test_date) FROM jea_watermetertesthistory y WHERE y.isr_no = e.isr_no AND ROWNUM < 2))



)

UNION --HISTORY
(
SELECT DISTINCT 
	a.compl_name cur_cust_name,
	u.udf24 fa_id,
 	JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'FA Schedule Date',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') FA_SCHD_DATE,
	a.compl_address address,
	a.zip_code,
	u.udf15 sp_type_code,
JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Clerk ID',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') Clerk_id,      
	--u.udf17 clerk_id,
	null terr_code,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Service on Departure ON',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') status_srv_on_depart,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Service on Arrival ON',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') STATUS_SRV_ON_arrival,
 JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Service on Departure OFF',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') status_srv_OFF_depart,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Service on Arrival OFF',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') STATUS_SRV_OFF_arrival,         
	a.init_service_code fa_type_code,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Work Phone',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') work_done,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Standard Comments 1',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') standard_comments1, 
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Standard Comments 2',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') standard_comments2, 
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Standard Comments 3',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') standard_comments3,            
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Completion Remarks',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') im_fa_comments,  
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Clerk Instructions',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') instructions,

  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Water Tap Size',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_sp_water_tap_size, 
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Service Size',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_sp_water_serv_size,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Sewer Tap Size',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_sp_sewer_tap_size,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Meter Loc Code',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_meter_loc_cd,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Water Tap Location',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_sp_water_tap_loc,

	a3.sp_water_tap_size,
	a3.sp_water_serv_size,
	a3.sp_sewer_tap_size,
	a.p_loc_type AS meter_loc_cd,
	a3.sp_water_tap_loc,

  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'New Meter Number',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_meter_number,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'New Meter Make',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_meter_manu,  
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'New Meter Size',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_meter_size,         
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'New Meter Model Type',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '')  nm_meter_type, 
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Meter Loc Details',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_meter_loc_detail,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'New Read A',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_current_reading1,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'New UOM A',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_current_reading1_uom, 
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'New Digits A',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_current_reading1_digits_l,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'New RF A',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') new_rf_id1,
null nm_current_reading1_tou,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'RID Installed A',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') new_rid_id1,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'New Read B',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_current_reading2,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'New UOM B',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_current_reading2_uom, 
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'New Digits B',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') nm_current_reading2_digits_l,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'New RF B',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') new_rf_id2,
null nm_current_reading2_tou,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'RID Installed B',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') new_rid_id2,
	null nm_current_reading3,          
	null nm_current_reading3_uom, 
	null nm_current_reading3_digits_l,
	null nm_current_reading3_rf_id,	
	null nm_current_reading3_tou,

--	u.udf3 om_meter_number,
      	 	JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Old Meter Number',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') om_meter_number,
	a3.manufacturer_code meter_manu,
	a3.meter_size,
	(SELECT his_jea_register.digits_left FROM his_jea_register WHERE isr_no = a.isr_no AND ROWNUM < 2) AS meter_digits_left,
	a3.type meter_type,
	a3.location_details meter_loc_details,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Read Today A',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') om_current_reading1,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'UOM A',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') om_current_reading1_uom,
	(SELECT digits_left FROM his_jea_register WHERE a.isr_no = isr_no AND read_seq = '1' AND ROWNUM < 2) AS om_current_reading1_digits_l,
	(SELECT y.badge_number FROM his_jea_item y WHERE y.isr_no = a.isr_no AND y.type_code = 'RFW' AND ROWNUM < 2) old_rfw_1, 
	(SELECT time_of_use FROM his_jea_register WHERE a.isr_no = isr_no AND read_seq = '1' AND ROWNUM < 2) AS om_current_reading1_tou,
	(SELECT register_id FROM his_jea_register WHERE a.isr_no = isr_no AND read_seq = '1' AND ROWNUM < 2) AS om_current_reading1_rid,
	(SELECT y.badge_number FROM his_jea_item y WHERE y.isr_no = a.isr_no AND y.type_code = 'RID' AND y.comments not LIKE 'B-BIG%' AND ROWNUM < 2) old_rid_1, 
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Read Today B',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '')om_current_reading2,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'UOM B',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') om_current_reading2_uom,
	(SELECT digits_left FROM his_jea_register WHERE a.isr_no = isr_no AND read_seq = '2' AND ROWNUM < 2) AS om_current_reading2_digits_l,
	(SELECT y.badge_number FROM his_jea_item y WHERE y.isr_no = a.isr_no AND y.type_code = 'RFW' AND ROWNUM < 2 
        AND y.badge_number <> (SELECT y.badge_number FROM his_jea_item y WHERE y.isr_no = a.isr_no AND y.type_code = 'RFW' AND ROWNUM < 2)) old_rfw_2,
	(SELECT time_of_use FROM his_jea_register WHERE a.isr_no = isr_no AND read_seq = '2' AND ROWNUM < 2) AS om_current_reading2_tou,
	(SELECT register_id FROM his_jea_register WHERE a.isr_no = isr_no AND read_seq = '2' AND ROWNUM < 2) AS om_current_reading2_rid,
	(SELECT y.badge_number FROM his_jea_item y WHERE y.isr_no = a.isr_no AND y.type_code = 'RID' AND y.comments LIKE 'B-BIG%' AND ROWNUM < 2) old_rid_2,  

  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Test Comments',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') new_test_comments,
	e.test_comments prev_test_comments,
  JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Test Passed',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') test_passed,
JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Test Failed',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') test_failed,
JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'No Test',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') no_test,        
	-- PRE TEST ---------------
JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'High Flow 1',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') new_high_flow, --test_result1
JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Medium Flow 1',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') new_medium_flow,
JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Low Flow 1',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') new_low_flow,
JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Average 1',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') new_average_flow,
JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'GPM 1',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') new_gpm, --test_result5
JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'PSI 1',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') new_psi,

	-- FINAL TEST --------------
  /*
	b.test_result7 new_reading_before_1,
	b.test_result8 new_reading_before_2, --test_result8
	b.test_result9 new_reading_before_3,
	b.test_result10 new_reading_after_1,
	b.test_result11 new_reading_after_2,
	b.test_result12 new_reading_after_3, --test_result12
*/
	TO_CHAR(TO_DATE(TO_NUMBER(DECODE(e.test_date,' ', '19000101',e.test_date)),'YYYYMMDD'), 'MM/DD/YYYY') prev_test_date,
	-- PREVIUOS TEST -----------
	e.high_flow prev_high_flow,
	e.medium_flow prev_medium_flow,
	e.low_flow prev_low_flow,
	e.average_flow prev_average_flow,
	e.gpm prev_gpm,
	e.psi prev_psi,

JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'MOD_DATETIME',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') as im_fa_status_date, 
	a.handling_unit tech_assigned, 
  a.isr_no

FROM his_isr a, 
     jeacust.his_jea_meter a3,          
     jeacust.his_jea_watermetertesthistory e,
     his_isr_udf u

WHERE a.isr_no = a3.isr_no (+)
	AND a.isr_no = e.isr_no (+)
  AND a.isr_no = u.isr_no (+)

	AND a.agency_code = 'METER'
	AND (a.p_status = 'CL' OR a.s_mode = 'L')
	AND a.p_filter1 = 'MTR-CW'
    and a.init_service_code not like '\_\_%' escape '\'

    and   nvl(JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'Work Done',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        '') ,' ') <> 'RETURN TO DISPATCH'
  AND
      trunc(TO_DATE(JEACUST.J_ISR_UDF_VAL_F(a.ISR_NO,
        'MOD_DATETIME',
        a.AGENCY_CODE, 
        a.INIT_SERVICE_CODE, 
        a.P_FILTER1,
        ''),'MM/DD/YYYY HH24:MI:SS')) BETWEEN 
            DECODE(:P_BEGIN_DATE,NULL,TRUNC(SYSDATE),TO_DATE(:P_BEGIN_DATE,'MM/DD/YYYY'))
  AND        
             DECODE(:P_END_DATE,NULL,trunc(sysdate),TO_DATE(:P_END_DATE,'MM/DD/YYYY'))


	AND (TRIM(e.test_date) IS NULL OR e.test_date = (SELECT MAX(y.test_date) FROM his_jea_watermetertesthistory y WHERE y.isr_no = e.isr_no AND ROWNUM < 2))



)
)
ORDER BY    

            im_fa_status_date,
              sp_type_code, 
              fa_id