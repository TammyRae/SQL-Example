/* script is used to defined the filter1 mapping of meter jobs to the 
correct scr_tag (screen data is viewed on) in order to retrieve the data
for reports */

/* T.Whitehead 1/4/2018 added S CAPITAL and W CAPITAL for water taps, mcons*/
/* T.Whitehead 1/11/2018 Added REHAB for Water Taps, mcons*/
--truncate table jeacust.j_filter1_template;

drop table jeacust.j_filter1_template;
create table jeacust.j_filter1_template
(dispatch_group varchar2(15),
 description varchar2(100),
 p_filter1 varchar2(20),
scr_tag_template varchar2(50)
);
GRANT Select on jeacust.j_filter1_template to BI_CAD_RPT;
GRANT Select on jeacust.j_filter1_template to DEVELOPER_INQ;
GRANT Select on jeacust.j_filter1_template to USERCAD;
GRANT Select on jeacust.j_filter1_template to FMS_WEBUSER;

insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values 
('CE','Commercial Services - Electric','MTR-CE','MTR_COMMELEC');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('CE','Commercial Services - Electric','MTR-CE','MTR_COMMELECNODE');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('CE','Commercial Services - Electric','MTR-CE','MTR_COMMELECNODE2');

insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('SE','SOCC - ELECTRIC','SOCC','MTR_POLE1');

insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('S CAPITAL','Water Taps - M-CONS','S CAPITAL','WATER GENERAL INFO');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('S CAPITAL','Water Taps - M-CONS','S CAPITAL','WATER READ-NEW METER');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('S CAPITAL','Water Taps - M-CONS','S CAPITAL','WATER COMPLETION');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('S CAPITAL','Water Taps - M-CONS','S CAPITAL','WATER RECLAIM TEST');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('S CAPITAL','Water Taps - M-CONS','S CAPITAL','WATER TEST RESULTS');

insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('W CAPITAL','Water Taps - M-CONS','W CAPITAL','WATER GENERAL INFO');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('W CAPITAL','Water Taps - M-CONS','W CAPITAL','WATER READ-NEW METER');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('W CAPITAL','Water Taps - M-CONS','W CAPITAL','WATER COMPLETION');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('W CAPITAL','Water Taps - M-CONS','W CAPITAL','WATER RECLAIM TEST');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('W CAPITAL','Water Taps - M-CONS','W CAPITAL','WATER TEST RESULTS');

insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('CW','Commercial Services - Water','MTR-CW','WATER GENERAL INFO');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('CW','Commercial Services - Water','MTR-CW','WATER READ-NEW METER');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('CW','Commercial Services - Water','MTR-CW','WATER COMPLETION');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('CW','Commercial Services - Water','MTR-CW','WATER RECLAIM TEST');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('CW','Commercial Services - Water','MTR-CW','WATER TEST RESULTS');

insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('WM','Residential Services - Water Maintenance','MTR-WM','WATER GENERAL INFO');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('WM','Residential Services - Water Maintenance','MTR-WM','WATER READ-NEW METER');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('WM','Residential Services - Water Maintenance','MTR-WM','WATER COMPLETION');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('WM','Residential Services - Water Maintenance','MTR-WM','WATER RECLAIM TEST');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('WM','Residential Services - Water Maintenance','MTR-WM','WATER TEST RESULTS');

insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('RW','Reclaim Water','MTR-RW','WATER GENERAL INFO');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('RW','Reclaim Water','MTR-RW','WATER READ-NEW METER');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('RW','Reclaim Water','MTR-RW','WATER COMPLETION');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('RW','Reclaim Water','MTR-RW','WATER RECLAIM TEST');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('RW','Reclaim Water','MTR-RW','WATER TEST RESULTS');

insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('SEWERMAINT','Sewer Maintenance','WATER-TAP','WATER GENERAL INFO');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('SEWERMAINT','Sewer Maintenance','WATER-TAP','WATER READ-NEW METER');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('SEWERMAINT','Sewer Maintenance','WATER-TAP','WATER COMPLETION');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('SEWERMAINT','Sewer Maintenance','WATER-TAP','WATER RECLAIM TEST');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('SEWERMAINT','Sewer Maintenance','WATER-TAP','WATER TEST RESULTS');


insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('WT','Water Tapping','WATER-TAP','WATER GENERAL INFO');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('WT','Water Tapping','WATER-TAP','WATER READ-NEW METER');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('WT','Water Tapping','WATER-TAP','WATER COMPLETION');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('WT','Water Tapping','WATER-TAP','WATER RECLAIM TEST');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('WT','Water Tapping','WATER-TAP','WATER TEST RESULTS');

insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('RSW','Revenue Services Water','MTR-RSW','WATER GENERAL INFO');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('RSW','Revenue Services Water','MTR-RSW','WATER READ-NEW METER');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('RSW','Revenue Services Water','MTR-RSW','WATER COMPLETION');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('RSW','Revenue Services Water','MTR-RSW','WATER RECLAIM TEST');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('RSW','Revenue Services Water','MTR-RSW','WATER TEST RESULTS');


insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('SW','SOCC - Water','SOCC','WATER GENERAL INFO');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('SW','SOCC - Water','SOCC','WATER READ-NEW METER');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('SW','SOCC - Water','SOCC','WATER COMPLETION');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('SW','SOCC - Water','SOCC','WATER RECLAIM TEST');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('SW','SOCC - Water','SOCC','WATER TEST RESULTS');

insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('REHAB','SOCC - Water','REHAB','WATER GENERAL INFO');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('REHAB','SOCC - Water','REHAB','WATER READ-NEW METER');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('REHAB','SOCC - Water','REHAB','WATER COMPLETION');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('REHAB','SOCC - Water','REHAB','WATER RECLAIM TEST');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('REHAB','SOCC - Water','REHAB','WATER TEST RESULTS');

insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('SS','SOCC - Sewer','SOCC','WATER GENERAL INFO');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('SS','SOCC - Sewer','SOCC','WATER READ-NEW METER');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('SS','SOCC - Sewer','SOCC','WATER COMPLETION');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('SS','SOCC - Sewer','SOCC','WATER RECLAIM TEST');
insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('SS','SOCC - Sewer','SOCC','WATER TEST RESULTS');

insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('EA','Energy Augit','MTR-EA','EA_APPT1');

insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('CD','Residential Services - Connect/Disconnect','MTR-CD','MTR_ELECTRIC1');


insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('CC','Customer Order Management Services (Service Fulfilment) - Construction','MTR-CC','MTR_ELECTRIC1');

insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('CT','Customer Order Management Services - Permits','MTR-TP','MTR_ELECTRIC1');

insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('RS','Revenue Services (Revenue Protection and Revenue Cycle Services','MTR-RS','MTR_ELECTRIC1');

insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('RCD','TwoWay Meter','MTR-RCD','MTR_ELECTRIC1');

insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('PP','PrePaid Meter','MTR-PP','MTR_ELECTRIC1');

insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('MR','Meter Read','MTR-MR','MTR_MR');

insert into j_filter1_template (dispatch_group, description,p_filter1,scr_tag_template)
Values
('RE','Reclaim Electric','MTR-RE','MTR_ELECTRIC1');
commit;



/
