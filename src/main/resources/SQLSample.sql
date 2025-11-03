DECLARE
      demoDocument     ExcelDocumentType;

      documentArray    ExcelDocumentLine := ExcelDocumentLine();

      clobDocument     CLOB;

       v_file        UTL_FILE.FILE_TYPE;
	   
	   v_from_name         VARCHAR2(1000);
	   v_to_name           VARCHAR2(4000);
	   
	--v_from_name         VARCHAR2(1000) := 'havicorel2@havi.com';
    --v_to_name           VARCHAR2(4000) := 'santosh.minupurey@havi.com;Laxman.Panigrahy@havi.com;Venkateswara.Bapathu@havi.com;sathish.mitta@havi.com';
	
	--v_to_name           VARCHAR2(100) := 'santosh.minupurey@havi.com';
 
    v_subject           VARCHAR2(100); -- := 'Havi Core Europe Daily DITL Status *** VRP TESTING ***';
    v_message_body      clob := '';
    v_message_type      VARCHAR2(100) := 'text/html';
 
    v_smtp_server       VARCHAR2(50)  := 'mailhost.perseco.com';
    n_smtp_server_port  NUMBER        := 25;
    conn                utl_smtp.connection;
 
    TYPE attach_info IS RECORD (
        attach_name     VARCHAR2(500),
        data_type       VARCHAR2(40) DEFAULT 'text/plain',
        attach_content  CLOB DEFAULT ''
    );
    TYPE array_attachments IS TABLE OF attach_info;
    attachments array_attachments := array_attachments();
 
    n_offset            NUMBER;
    n_amount            NUMBER        := 5000;
    v_crlf              VARCHAR2(5)   := CHR(13) || CHR(10);
	l_day				VARCHAR2(30);
	l_comp_cnt			NUMBER;
	l_comp_chk			NUMBER;
	l_day1				VARCHAR2(30);
	l_day2				VARCHAR2(30);
	l_day3				VARCHAR2(30);
	l_day4				VARCHAR2(30);
	l_day5				VARCHAR2(30);
	ddl_query 			varchar2(500);
	ddl_drop		    varchar2(500);
	
	ctxh           DBMS_XMLGEN.ctxHandle;
    queryresult    XMLTYPE;
    xslt_tranfsorm XMLTYPE;
	
	C_BORDERS_XML CONSTANT VARCHAR2(400) := '<Borders>
        <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
        <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
        <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
        <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
       </Borders>';
	  
  BEGIN
  
  ddl_drop := 'drop view hc_pmix_mrng_chk_vw';
  
  EXECUTE IMMEDIATE ddl_drop;
  
  --ddl_query := 'create or replace view hc_pmix_mrng_chk_vw as SELECT * FROM TABLE (pivotrep (''select * from hc_pmix_mrng_chk'')) ORDER BY 2, 1';
    ddl_query := 'create or replace view hc_pmix_mrng_chk_vw as SELECT * FROM TABLE (pivotrep (''select rep,cust_id,to_char(bday,''''yyyy.mm.dd'''') bday,loc_cnt from hc_pmix_mrng_chk'')) ORDER BY 2, 1';
  
  EXECUTE IMMEDIATE ddl_query;
  
  delete from hc_pmix_mrng_chk_rep;

	insert into hc_pmix_mrng_chk_rep
	select * from  hc_pmix_mrng_chk_vw -- where cust_id=7
	order by 2,1;

	commit;
	
	with tt as (
select min(bday) minbday,max(bday) maxbday,max(bday)-min(bday) bday_cnt from hc_pmix_mrng_chk order by bday
)
select minbday day5,minbday+1 day4,minbday+2 day3,minbday+3 day2,maxbday day1 into l_day5,l_day4,l_day3,l_day2,l_day1 from tt;

     demoDocument := ExcelDocumentType();

     -- Open the document
     demoDocument.documentOpen;

     -- Define Styles

     demoDocument.stylesOpen;

     -- Include Default Style
     demoDocument.defaultStyle;

     -- Add Custom Styles

     /* Style for Column Header Row */
     demoDocument.createStyle(p_style_id =>'ColumnHeader',
							   p_cell_color=>'#92d050',
							   p_cell_pattern =>'Solid',
                               p_font     =>'Calibri',
                               --p_ffamily  =>'Roman',
                               p_fsize    =>'11',
                               p_bold     =>'Y',
                               --p_underline =>'Single',
                               p_align_horizontal=>'Center',
                               p_align_vertical=>'Bottom',
							   p_custom_xml => C_BORDERS_XML
							   );
							   
    /* Styles for alternating row colors. */ 
    demoDocument.createStyle(p_style_id=>'NumberStyleBlueCell',
                               p_cell_color=>'Cyan',
                               p_cell_pattern =>'Solid',
                               p_number_format => '####',
                               p_align_horizontal => 'Right');

    demoDocument.createStyle(p_style_id=>'TextStyleBlueCell',
                               p_cell_color=>'Cyan',
                               p_cell_pattern =>'Solid');

    /* Style for numbers */
    demoDocument.createStyle(p_style_id => 'NumberStyle',
                              p_number_format => '####',
                              p_align_horizontal => 'Right',
							  p_custom_xml => C_BORDERS_XML);

   /* Style for Column Sum */
    demoDocument.createStyle(p_style_id => 'ColumnSum',
                              p_number_format => '###,###,###.00',
                              p_align_horizontal => 'Right',
                              p_text_color => 'Blue'); 

   /* Style for Column Sum */
    demoDocument.createStyle(p_style_id => 'RowSum',
                              p_number_format => '###,###,###.00',
                              p_align_horizontal => 'Right',
                              p_text_color => 'Red');  


     -- Close Styles
     demoDocument.stylesClose;
	 
	 
/* *****************************************************
				SUPPLY PLAN STARTS HERE
********************************************************* */	

for cust in (select * from hc_customer where 1=1 and market_name NOT IN ('HR','RU') order by cust_id) loop
	-- Open New Worksheet
  demoDocument.worksheetOpen('Supply_Plan_Check_'||cust.market_name);	
  
     -- Define Columns
     demoDocument.defineColumn(p_index=>'1',p_width=>30); -- Emp Name
     demoDocument.defineColumn(p_index=>'2',p_width=>30); -- Daily Dollar
     demoDocument.defineColumn(p_index=>'3',p_width=>30);
     
    -- Define Header Row
	
   demoDocument.rowOpen;
   
      demoDocument.addCell(p_custom_attr => 'ss:MergeAcross="2"',p_style=>'ColumnHeader',p_data=>'Supply Plan Check_'||cust.market_name||' '||TO_CHAR(FROM_TZ(CAST(SYSDATE AS TIMESTAMP), 'CST') AT TIME ZONE 'Asia/Calcutta','DD-MON-YYYY'));
   
   demoDocument.rowClose; 
	
   demoDocument.rowOpen;

   demoDocument.addCell(p_style=>'ColumnHeader',p_data=>'Total Number of Locations');
   demoDocument.addCell(p_style=>'ColumnHeader',p_data=>'Number of Loc with Plan Output');
   demoDocument.addCell(p_style=>'ColumnHeader',p_data=>'Number of Loc with Intransits');
   
   demoDocument.rowClose; 

   FOR rec IN (SELECT * FROM supply_plan WHERE market_name=cust.market_name) LOOP
   
   demoDocument.rowOpen;
   
   demoDocument.addCell(p_style=>'NumberStyle',p_data_type=>'Number', p_data=>rec.total_loc);
   demoDocument.addCell(p_style=>'NumberStyle',p_data_type=>'Number', p_data=>rec.loc_with_plan_op);
   demoDocument.addCell(p_style=>'NumberStyle',p_data_type=>'Number', p_data=>rec.loc_with_intransits);   
   demoDocument.rowClose;
   
   END LOOP;
   
  
    -- Define Columns
     --demoDocument.defineColumn(p_index=>'3',p_width=>30); -- Emp Name
     --demoDocument.defineColumn(p_index=>'4',p_width=>30); -- Daily Dollar
     ---demoDocument.defineColumn(p_index=>'3',p_width=>30);
     
    -- Define Header Row
	
	demoDocument.rowOpen;
   demoDocument.addCell(p_data=>' ');   
   demoDocument.rowClose;
   
   demoDocument.rowOpen;
   demoDocument.addCell(p_data=>' ');   
   demoDocument.rowClose;
   
   demoDocument.rowOpen;
   demoDocument.addCell(p_data=>' ');   
   demoDocument.rowClose;
	
   demoDocument.rowOpen;
   
      demoDocument.addCell(p_custom_attr => 'ss:MergeAcross="1"',p_style=>'ColumnHeader',p_data=>'Plan Quality Check_'||cust.market_name||' '||TO_CHAR(FROM_TZ(CAST(SYSDATE AS TIMESTAMP), 'CST') AT TIME ZONE 'Asia/Calcutta','DD-MON-YYYY'));
   
   demoDocument.rowClose; 
	
   demoDocument.rowOpen;

   demoDocument.addCell(p_style=>'ColumnHeader',p_data=>'Description');
   demoDocument.addCell(p_style=>'ColumnHeader',p_data=>'Count');
   --demoDocument.addCell(p_style=>'ColumnHeader',p_data=>'Number of Loc with Intransits');
   
   demoDocument.rowClose; 

   --FOR rec IN (SELECT * FROM supply_plan WHERE market_name=cust.market_name) LOOP
   FOR rec IN (select * from hc_plan_quality WHERE udc_customer=cust.market_name
				order by 1, case when trim(descr)='Vehicleload Yes'	then 1 
                 when trim(descr)='Vehicleload No'	then 2 
                 when trim(descr)='Recship Yes'	then 3 
                 when trim(descr)='Recship No'	then 4 
                 when trim(descr)='Fair Share Recship Yes'	then 5 
                 when trim(descr)='Fair Share Recship No'	then 6 
            end
            ) LOOP
   
   demoDocument.rowOpen;
   
   demoDocument.addCell(p_style=>'NumberStyle', p_data=>rec.descr);
   demoDocument.addCell(p_style=>'NumberStyle',p_data_type=>'Number', p_data=>rec.cnt);
   --demoDocument.addCell(p_style=>'NumberStyle',p_data_type=>'Number', p_data=>rec.loc_with_intransits);   
   demoDocument.rowClose;
   
   END LOOP;
   
  -- Close the Worksheet
  demoDocument.worksheetClose;

  end loop; 
  
 
	 
/* ****************************************************
				PMIX REPORT STARTS HERE
************************************************ */	 


for cust1 in (select * from hc_customer where 1=1 and market_name NOT IN ('HR','RU') order by cust_id) loop  
  
  demoDocument.worksheetOpen('Pmix_Report_'||cust1.market_name);

     -- Define Columns
     demoDocument.defineColumn(p_index=>'1',p_width=>25);
     demoDocument.defineColumn(p_index=>'2',p_width=>30);
     demoDocument.defineColumn(p_index=>'3',p_width=>30);
     demoDocument.defineColumn(p_index=>'4',p_width=>30);
	 demoDocument.defineColumn(p_index=>'5',p_width=>20);
	 demoDocument.defineColumn(p_index=>'6',p_width=>20);
     
    -- Define Header Row
   demoDocument.rowOpen;
   
      demoDocument.addCell(p_custom_attr => 'ss:MergeAcross="3"',p_style=>'ColumnHeader',p_data=>'Pmix_Report_'||cust1.market_name||' '||TO_CHAR(FROM_TZ(CAST(SYSDATE AS TIMESTAMP), 'CST') AT TIME ZONE 'Asia/Calcutta','DD-MON-YYYY'));
   
   demoDocument.rowClose; 		
	
   demoDocument.rowOpen;

   --Define Header Row Data Cells
   demoDocument.addCell(p_style=>'ColumnHeader',p_data=>'Business Day');
   demoDocument.addCell(p_style=>'ColumnHeader',p_data=>'Number of Stores Reported PMIX');
   demoDocument.addCell(p_style=>'ColumnHeader',p_data=>'Number of Stores Missing PMIX');
   demoDocument.addCell(p_style=>'ColumnHeader',p_data=>'Average PMIX Records Per Store');

   demoDocument.rowClose; 


   FOR rec IN (SELECT * FROM pmix_rep WHERE cust_id=cust1.cust_id ORDER BY TO_DATE(bday,'DD-Mon-YYYY')) LOOP
   
   demoDocument.rowOpen;
   demoDocument.addCell(p_style=>'NumberStyle',p_data=>rec.bday);
   demoDocument.addCell(p_style=>'NumberStyle',p_data_type=>'Number', p_data=>rec.stores_rep);
   demoDocument.addCell(p_style=>'NumberStyle',p_data_type=>'Number', p_data=>rec.stores_miss);
   demoDocument.addCell(p_style=>'NumberStyle',p_data_type=>'Number', p_data=>rec.avg_store);   
   demoDocument.rowClose;
   
   END LOOP;
   
   demoDocument.rowOpen;
   demoDocument.addCell(p_data=>' ');   
   demoDocument.rowClose;
   
   demoDocument.rowOpen;
   demoDocument.addCell(p_data=>' ');   
   demoDocument.rowClose;
   
   demoDocument.rowOpen;
   demoDocument.addCell(p_data=>' ');   
   demoDocument.rowClose;
   
   demoDocument.rowOpen;
   
      demoDocument.addCell(p_custom_attr => 'ss:MergeAcross="5"',p_style=>'ColumnHeader',p_data=>'Pmix_Morning_Check_Report_'||cust1.market_name||' '||TO_CHAR(FROM_TZ(CAST(SYSDATE AS TIMESTAMP), 'CST') AT TIME ZONE 'Asia/Calcutta','DD-MON-YYYY'));
   
   demoDocument.rowClose; 		
	
   demoDocument.rowOpen;

   --Define Header Row Data Cells
   demoDocument.addCell(p_style=>'ColumnHeader',p_data=>' ');
   demoDocument.addCell(p_style=>'ColumnHeader',p_data=>l_day5);
   demoDocument.addCell(p_style=>'ColumnHeader',p_data=>l_day4);
   demoDocument.addCell(p_style=>'ColumnHeader',p_data=>l_day3);
   demoDocument.addCell(p_style=>'ColumnHeader',p_data=>l_day2);
   demoDocument.addCell(p_style=>'ColumnHeader',p_data=>l_day1);

   demoDocument.rowClose; 

   FOR rec IN (select * from hc_pmix_mrng_chk_rep where cust_id=cust1.cust_id order by 2,1 ) LOOP
   
   demoDocument.rowOpen;
   demoDocument.addCell(p_style=>'NumberStyle',p_data=>rec.rep);
   demoDocument.addCell(p_style=>'NumberStyle',p_data_type=>'Number', p_data=>rec.day5);
   demoDocument.addCell(p_style=>'NumberStyle',p_data_type=>'Number', p_data=>rec.day4);
   demoDocument.addCell(p_style=>'NumberStyle',p_data_type=>'Number', p_data=>rec.day3);   
   demoDocument.addCell(p_style=>'NumberStyle',p_data_type=>'Number', p_data=>rec.day2);   
   demoDocument.addCell(p_style=>'NumberStyle',p_data_type=>'Number', p_data=>rec.day1);   
   demoDocument.rowClose;
   
   END LOOP;
   
  -- Close the Worksheet
  demoDocument.worksheetClose;
  
end loop;  

  /* ****************************************************
				PMIX MISSING STARTS HERE
************************************************ */	 
  
for cust2 in (select * from hc_customer where 1=1 and market_name NOT IN ('HR','RU') order by cust_id) loop  

-- Open New Worksheet
  demoDocument.worksheetOpen('Pmix_Missing_'||cust2.market_name);

     -- Define Columns
     demoDocument.defineColumn(p_index=>'1',p_width=>15); -- Emp Name
     demoDocument.defineColumn(p_index=>'2',p_width=>30); -- Daily Dollar
    
    -- Define Header Row
   demoDocument.rowOpen;
   
      demoDocument.addCell(p_custom_attr => 'ss:MergeAcross="1"',p_style=>'ColumnHeader',p_data=>'Pmix_Missing_'||cust2.market_name||' '||TO_CHAR(FROM_TZ(CAST(SYSDATE AS TIMESTAMP), 'CST') AT TIME ZONE 'Asia/Calcutta','DD-MON-YYYY'));
   
   demoDocument.rowClose; 	
	
   demoDocument.rowOpen;

   demoDocument.addCell(p_style=>'ColumnHeader',p_data=>'Business Day');
   demoDocument.addCell(p_style=>'ColumnHeader',p_data=>'Restaurant # Missing PMIX');
   
   demoDocument.rowClose; 

   FOR rec IN (SELECT * FROM pmix_miss WHERE cust_id=cust2.cust_id order by to_date(bday,'DD-Mon-YYYY')) LOOP
   
   demoDocument.rowOpen;
   demoDocument.addCell(p_style=>'NumberStyle',p_data=>rec.bday);
   demoDocument.addCell(p_style=>'NumberStyle',p_data_type=>'Number', p_data=>rec.rest#_missing_pmix);
   demoDocument.rowClose;
   
   END LOOP;
   
  -- Close the Worksheet
  demoDocument.worksheetClose;

end loop;
  
  
  /* ********************************
		MASTER DATA LOAD STARTS HERE
  ************************************ */
  
for cust3 in (select * from hc_customer where 1=1 and market_name NOT IN ('HR','RU') order by cust_id) loop    
  -- Open New Worksheet
  demoDocument.worksheetOpen(cust3.market_name||'_MasterDataLoad_StartTime');

     -- Define Columns
     demoDocument.defineColumn(p_index=>'1',p_width=>15); -- Emp Name
     demoDocument.defineColumn(p_index=>'2',p_width=>28); -- Daily Dollar
    
    -- Define Header Row
   demoDocument.rowOpen;

   demoDocument.addCell(p_style=>'ColumnHeader',p_data=>'Business Day');
   demoDocument.addCell(p_style=>'ColumnHeader',p_data=>cust3.market_name||'_MasterDataLoad_StartTime');
   
   demoDocument.rowClose; 

   FOR rec IN (SELECT * FROM master_data WHERE market_name=cust3.market_name) LOOP
   
   demoDocument.rowOpen;
   demoDocument.addCell(p_style=>'NumberStyle',p_data=>rec.start_dt_eu);
   demoDocument.addCell(p_style=>'NumberStyle',p_data=>rec.start_time_eu);
   demoDocument.rowClose;
   
   END LOOP;
   
  -- Close the Worksheet
  demoDocument.worksheetClose;
  
end loop;    
  

  -- Close the document.
  demoDocument.documentClose;

  -- Get CLOB Version

  clobDocument := demoDocument.getDocument;
dbms_output.put_line('****End****');

attachments.extend(1);
   FOR i IN 1..1
    LOOP
        SELECT 'Operational_Exception_Report-' || TO_CHAR((FROM_TZ(CAST(SYSDATE AS TIMESTAMP), 'CST') AT TIME ZONE 'Europe/Brussels'),'DD-MON-YYYY') || '.xls','text/plain',clobDocument
        INTO attachments(i)
        FROM dual;
    END LOOP;
dbms_output.put_line('****INit****');	

	v_subject	:= NULL;
	l_day		:= NULL;
	v_from_name	:= NULL;
	v_to_name	:= NULL;
	l_comp_cnt  := 0;
	l_comp_chk	:= 0;

BEGIN	
	SELECT from_address,to_address INTO v_from_name,v_to_name
	FROM hc_email_recipents
	WHERE 1=1
	AND market='EU'
	AND email_type='DITL_STATUS';	
EXCEPTION
	WHEN no_data_found THEN
		v_from_name		:= 'havicorel2@havi.com';
		v_to_name		:= 'santosh.minupurey@havi.com';
END;

select count(1) into l_comp_cnt from ditl_status WHERE 1=1;

SELECT count(1) into l_comp_chk FROM master_data WHERE 1=1;
			

SELECT TO_CHAR(FROM_TZ(CAST(SYSDATE AS TIMESTAMP), 'CST') AT TIME ZONE 'Asia/Calcutta','Day') INTO l_day FROM dual;

dbms_output.put_line('****l_day****' || l_day || length(trim(l_day)));

IF trim(l_day) = 'Sunday' THEN
	v_subject := 'Havi Core Europe and Market Enabler Weekly DITL Status *** TESTING ***';
	dbms_output.put_line('****Weekend****');	
ELSE
	v_subject := 'Havi Core Europe and Market Enabler Daily DITL Status *** TESTING ***';
	dbms_output.put_line('****Weekday****');	
END IF;

IF l_comp_cnt<>l_comp_chk THEN
v_subject := 'INCOMPLETE :::: '||v_subject;
END IF;


dbms_output.put_line('****v_subject****' || v_subject);
  --attachments(1) := 'test1.xls','text/plain',clobDocument;
 
 
/*v_message_body:='<center><h3><i><font color=#000099>List of Oracle Users and their account statuses </font></i></h3></center><br>'||utl_tcp.CRLF;
  v_message_body:=v_message_body||'<table style="border: solid 0px #cccccc"  cellspacing="0" cellpadding="0"><tr BGCOLOR=#000099>';
  v_message_body:=v_message_body||'<td><b><font color=white>Username</font></td>';
  v_message_body:=v_message_body||'<td><b><font color=white>Account Status</font></td></tr>'||utl_tcp.CRLF;
  FOR I IN (select username, account_status from dba_users)
  LOOP
     v_message_body:=v_message_body||'<tr><td>'||i.username||'</td><td>'||i.account_status||'</td></tr>'||utl_tcp.CRLF;
  END LOOP;
*/

	v_message_body := '<p style="font-family:Calibri; font-size:95%;">'||'Hi All,'||'</p>';
	
	IF trim(l_day) != 'Sunday' THEN
		v_message_body := v_message_body||'<p style="font-family:Calibri; font-size:95%;">'||'Please find the below Daily DITL Status.'||'</p>';
	ELSE
		v_message_body := v_message_body||'<p style="font-family:Calibri; font-size:95%;">'||'Please find the below Weekly DITL Status.'||'</p>';
	END IF;
	
	--v_message_body := v_message_body||'<p style="font-family:Calibri; font-size:95%;">'||'Please find the below Daily DITL Status.'||'</p>';
	
  v_message_body:=v_message_body||'<style type="text/css"> table {width: 100%;} table, th, td {font-family: "Calibri"; font-size:"100%"; border: 1px solid black; border-collapse: collapse; padding: 0px;} </style>';
  --v_message_body:=v_message_body||'<style type="text/css">     table { background: #FEFEFE; font-family: "Calibri"; font-size: 100%; border-collapse: collapse; border: 1px solid black;}     th { background: #00b0f0; }     td { padding: 0px; }  </style>';
  --'table, th, td {border: 1px solid black; border-collapse: collapse;}';
  v_message_body:=v_message_body||'<table style="border: 1px solid black" "width:100%" border-collapse: collapse cellspacing="5" cellpadding="0" ><tr BGCOLOR=#00b0f0>';
  v_message_body:=v_message_body||'<th nowrap><center><b><font color=black>S .No</font></center></th>';
  v_message_body:=v_message_body||'<th nowrap><center><b><font color=black>Market Name</font></center></th>';
  v_message_body:=v_message_body||'<th nowrap><center><b><font color=black>Plan day</font></center></th>';
  v_message_body:=v_message_body||'<th nowrap><center><b><font color=black>Demand day</font></center></th>';
  v_message_body:=v_message_body||'<th nowrap><center><b><font color=black>Status</font></center></th>';
  v_message_body:=v_message_body||'<th nowrap><center><b><font color=black>SLO Time</font></center></th>';
  v_message_body:=v_message_body||'<th nowrap><center><b><font color=black>Plan Output Published<br> (Order file and Exception file)</font></center></th>';
  v_message_body:=v_message_body||'<th nowrap><center><b><font color=black>ME SLO Time</font></center></th>';
  v_message_body:=v_message_body||'<th nowrap><center><b><font color=black>Plan Results<br> Processed in ME (VRP)</font></center></th>';
  v_message_body:=v_message_body||'<th nowrap><center><b><font color=black>ME Status</font></center></th>';
  v_message_body:=v_message_body||'<th nowrap><center><b><font color=black>HIST File</font></center></th>';
  v_message_body:=v_message_body||'<th nowrap><center><b><font color=black>FCST File</font></center></th>';
  v_message_body:=v_message_body||'<th nowrap><center><b><font color=black>Comments</font></center></th></tr>'||utl_tcp.CRLF;
  
  FOR I IN (SELECT sl_no,market_name,plan_date,demand_date,
            CASE 
            WHEN NVL(plan_chk,0)=0 AND TO_DATE(plan_pub,'HH:Mi:SS AM') <= TO_DATE(slo_time,'HH:Mi:SS AM') THEN '<td bgcolor="#ff0000" nowrap><center><font color="#ffffff">'||'RED'||'</font></center></td>'
            WHEN NVL(plan_chk,0)=0 AND TO_DATE(plan_pub,'HH:Mi:SS AM') > TO_DATE(slo_time,'HH:Mi:SS AM') THEN '<td bgcolor="#ff0000"  nowrap><center><font color="#ffffff">'||TO_NUMBER(TO_CHAR(TRUNC((TO_DATE(plan_pub,'HH:MI:SS AM')- TO_DATE(slo_time,'HH:MI:SS AM'))*24*60), '000'))||':'||             
                 LTRIM(TO_CHAR(TRUNC(MOD((TO_DATE(plan_pub,'HH:MI:SS AM')- TO_DATE(slo_time,'HH:MI:SS AM'))*24*60*60, 60)), '00'))||'</font></center></td>'
            WHEN NVL(plan_chk,0)=1 AND TO_DATE(plan_pub,'HH:Mi:SS AM')<= TO_DATE(slo_time,'HH:Mi:SS AM') AND loc_chk='GREEN' AND file_chk IS NOT NULL THEN '<td bgcolor="#ffff00" nowrap><center><font color="#000000">'||'YELLOW'||'</font></center></td>' 
            WHEN NVL(plan_chk,0)=1 AND TO_DATE(plan_pub,'HH:Mi:SS AM')<= TO_DATE(slo_time,'HH:Mi:SS AM') AND loc_chk='GREEN' THEN '<td bgcolor="#00b050" nowrap><center><font color="#ffffff">'||'GREEN'||'</font></center></td>'
            WHEN NVL(plan_chk,0)=1 AND TO_DATE(plan_pub,'HH:Mi:SS AM')<= TO_DATE(slo_time,'HH:Mi:SS AM') AND loc_chk='RED' THEN '<td bgcolor="#ff0000" nowrap><center><font color="#ffffff">'||'RED'||'</font></center></td>' 
            WHEN NVL(plan_chk,0)=1 AND TO_DATE(plan_pub,'HH:Mi:SS AM') > TO_DATE(slo_time,'HH:Mi:SS AM') THEN '<td bgcolor="#ff0000"  nowrap><center><font color="#ffffff">'||TO_NUMBER(TO_CHAR(TRUNC((TO_DATE(plan_pub,'HH:MI:SS AM')- TO_DATE(slo_time,'HH:MI:SS AM'))*24*60), '000'))||':'||             
                 LTRIM(TO_CHAR(TRUNC(MOD((TO_DATE(plan_pub,'HH:MI:SS AM')- TO_DATE(slo_time,'HH:MI:SS AM'))*24*60*60, 60)), '00'))||'</font></center></td>'
            END status,
            --TO_DATE(completion_time,'HH:Mi:SS AM'),TO_DATE(me_slo_time,'HH:Mi:SS AM') ,            
            CASE 
            WHEN TO_DATE(comp_time,'HH:Mi:SS AM') <= TO_DATE(me_slo_time,'HH:Mi:SS AM') --OR market_name='FINLAND'
            THEN '<td bgcolor="#00b050" nowrap><center><font color="#ffffff">'||'GREEN'||'</font></center></td>'
            WHEN TO_DATE(comp_time,'HH:Mi:SS AM') > TO_DATE(me_slo_time,'HH:Mi:SS AM') and comp_time is not null
            THEN '<td bgcolor="#ff0000"  nowrap><center><font color="#ffffff">'||TO_NUMBER(TO_CHAR(TRUNC((TO_DATE(comp_time,'HH:MI:SS AM')- TO_DATE(me_slo_time,'HH:MI:SS AM'))*24*60), '000'))||':'||             
                 LTRIM(TO_CHAR(TRUNC(MOD((TO_DATE(comp_time,'HH:MI:SS AM')- TO_DATE(me_slo_time,'HH:MI:SS AM'))*24*60*60, 60)), '00'))||'</font></center></td>'
			ELSE '<td bgcolor="#ff0000" nowrap><center><font color="#ffffff">'||'RED'||'</font></center></td>'
            END me_status, 
slo_time||DECODE(SUBSTR(TZ_OFFSET('Europe/Brussels'),2,2),'01',' CET',' CEST') slo_time,
me_slo_time||DECODE(SUBSTR(TZ_OFFSET('Europe/Brussels'),2,2),'01',' CET',' CEST') me_slo_time,
plan_publish_time,completion_time,hist_file_time,fcst_file_time,
CASE WHEN NVL(plan_chk,0)=0 THEN 'Supply plan output has not been generated, please check for source files availability.'
     WHEN NVL(plan_chk,0)=1 AND loc_chk='RED' THEN 'Number of Loc with Plan Output is less then 50%.'
     WHEN NVL(plan_chk,0)=1 AND loc_chk='GREEN' AND file_chk IS NOT NULL and file_exists='N' THEN 'No Master data file for '||INITCAP(market_name)||' market today, so used yesterday''s ('||TO_CHAR((TO_DATE(plan_date,'DD-Mon-YYYY')-1),'MM/DD')||') master data.'
     WHEN NVL(plan_chk,0)=1 AND loc_chk='GREEN' AND file_chk IS NOT NULL and zero_partial_flag='Y' THEN 'Zero Byte/Partial Master data file received for '||INITCAP(market_name)||' market today, Kindly review plan results.'
     --WHEN market_name='BELGIUM' THEN 'Currently receiving only Master Data files for BELGIUM market'
     ELSE NULL END comments,
NVL(plan_chk,0) plan_chk,loc_chk,file_chk
FROM(
SELECT row_number() OVER (ORDER BY CASE WHEN d.market_name='SE' THEN 3
                          WHEN d.market_name='AT' THEN 4
                          WHEN d.market_name='RU' THEN 5
                          WHEN d.market_name='DE' THEN 6
                          WHEN d.market_name='IT' THEN 7
                          WHEN d.market_name='PL' THEN 8
                          WHEN d.market_name='ES' THEN 9
                          WHEN d.market_name='NL' THEN 10
                          WHEN d.market_name='BE' THEN 11
                          WHEN d.market_name='CZ' THEN 12
                          WHEN d.market_name='SK' THEN 13
						  WHEN d.market_name='UA' THEN 14
						  WHEN d.market_name='DK' THEN 15
						  WHEN d.market_name='NO' THEN 16
						  WHEN d.market_name='FI' THEN 17
						  WHEN d.market_name='CH' THEN 18
						  WHEN d.market_name='HR' THEN 19
						  WHEN d.market_name='PT' THEN 20
                          ELSE 100
                          END) sl_no,
                    DECODE(d.market_name,'SE','SWEDEN',
                                       'AT','AUSTRIA',
                                       'RU','GEORGIA',
                                       'DE','GERMANY',
                                       'IT','ITALY',
                                       'PL','POLAND',
                                       'ES','SPAIN',
                                       'NL','NETHERLANDS',
                                       'BE','BELGIUM',
                                       'CZ','CZECH REPUBLIC',
                                       'SK','SLOVAKIA',
                                       'UA','UKRAINE',
									   'DK','DENMARK',
									   'NO','NORWAY',
									   'FI','FINLAND',
									   'CH','SWITZERLAND',
									   'HR','CROATIA',
									   'PT','PORTUGAL',
                                       d.market_name) market_name,
       plan_date,
       TO_CHAR(NEXT_DAY(FROM_TZ(CAST(SYSDATE-8 AS TIMESTAMP), 'CST') AT TIME ZONE 'Europe/Brussels','FRI'),'DD-MON-YYYY') demand_date,
       TO_CHAR(TO_DATE(SUBSTR(plan_publish_time,1,INSTR(plan_publish_time,' ',1,2)-1),'HH:Mi:SS AM'),'HH:Mi:SS AM') plan_pub,
         TO_CHAR(TO_DATE(SUBSTR(completion_time,1,INSTR(completion_time,' ',1,2)-1),'HH:Mi:SS AM'),'HH:Mi:SS AM') comp_time,
       CASE WHEN d.market_name='SE' THEN TO_CHAR(TO_DATE('08:45:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
            WHEN d.market_name='AT' THEN TO_CHAR(TO_DATE('06:00:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
            WHEN d.market_name='RU' THEN TO_CHAR(TO_DATE('06:30:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
            WHEN d.market_name='DE' THEN TO_CHAR(TO_DATE('07:00:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
            WHEN d.market_name='IT' THEN TO_CHAR(TO_DATE('08:00:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
            WHEN d.market_name='PL' THEN TO_CHAR(TO_DATE('08:00:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
            WHEN d.market_name='ES' THEN TO_CHAR(TO_DATE('07:30:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
            WHEN d.market_name='NL' THEN TO_CHAR(TO_DATE('06:30:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
            WHEN d.market_name='BE' THEN TO_CHAR(TO_DATE('06:30:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
            WHEN d.market_name='CZ' THEN TO_CHAR(TO_DATE('07:00:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
            WHEN d.market_name='SK' THEN TO_CHAR(TO_DATE('07:00:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
            WHEN d.market_name='UA' THEN TO_CHAR(TO_DATE('08:00:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
			WHEN d.market_name='DK' THEN TO_CHAR(TO_DATE('08:00:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
			WHEN d.market_name='FI' THEN TO_CHAR(TO_DATE('08:00:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
			WHEN d.market_name='NO' THEN TO_CHAR(TO_DATE('08:00:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
			WHEN d.market_name='CH' THEN TO_CHAR(TO_DATE('07:00:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
			WHEN d.market_name='PT' THEN TO_CHAR(TO_DATE('08:00:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
        END SLO_TIME,
        CASE WHEN d.market_name='SE' THEN TO_CHAR(TO_DATE('09:00:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
            WHEN d.market_name='AT' THEN TO_CHAR(TO_DATE('06:30:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
            WHEN d.market_name='RU' THEN TO_CHAR(TO_DATE('07:30:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
            WHEN d.market_name='DE' THEN TO_CHAR(TO_DATE('08:00:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
            WHEN d.market_name='IT' THEN TO_CHAR(TO_DATE('08:30:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
            WHEN d.market_name='PL' THEN TO_CHAR(TO_DATE('08:30:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
            WHEN d.market_name='ES' THEN TO_CHAR(TO_DATE('08:00:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
            WHEN d.market_name='NL' THEN TO_CHAR(TO_DATE('07:00:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
            WHEN d.market_name='BE' THEN TO_CHAR(TO_DATE('07:00:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
            WHEN d.market_name='CZ' THEN TO_CHAR(TO_DATE('07:30:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
            WHEN d.market_name='SK' THEN TO_CHAR(TO_DATE('07:30:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
            WHEN d.market_name='UA' THEN TO_CHAR(TO_DATE('08:30:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
			WHEN d.market_name='DK' THEN TO_CHAR(TO_DATE('08:30:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
			WHEN d.market_name='FI' THEN TO_CHAR(TO_DATE('08:30:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
			WHEN d.market_name='NO' THEN TO_CHAR(TO_DATE('08:30:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
			WHEN d.market_name='CH' THEN TO_CHAR(TO_DATE('07:30:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
			WHEN d.market_name='PT' THEN TO_CHAR(TO_DATE('08:30:00 AM','HH:Mi:SS AM'),'HH:Mi:SS AM')
        END ME_SLO_TIME,
       --CASE 
       plan_publish_time,
       pr.completion_time, 
       hist_file_time,fcst_file_time,s.plan_chk,s1.loc_chk,fc.market file_chk,file_exists,zero_partial_flag
FROM ditl_status d,
(select * from plan_re) pr,
--FROM ditl_status_20171112 d,
(SELECT market_name,
CASE WHEN COUNT(1)>0 THEN 1 ELSE 0 END plan_chk FROM supply_plan 
WHERE 1=1
--AND market_name='SE'
GROUP BY market_name
--UNION
--SELECT 'BE ',1 FROM dual
UNION
SELECT 'CH',1 FROM dual
) s,
(SELECT market_name,
 CASE WHEN loc_with_plan_op<ROUND(total_loc*.50) AND market_name NOT IN ('HR','HU')  THEN 'RED' ELSE 'GREEN' END loc_chk   FROM supply_plan
 --UNION
 --SELECT 'BE','GREEN' FROM dual
 ) s1,
 (WITH file_chk AS
    (
      SELECT cust_id,market,file_exists,case when ZERO_BYTE_FILE='Y' then 'Y'
                 when NOTIFY_PARTIAL_FILE='Y' then 'Y'
                 else 'N'
                 end zero_partial_flag FROM hc_file_chk --WHERE business_day=TRUNC(SYSDATE)
    )
    SELECT DISTINCT NVL(cust_id,0) cust_id ,NVL(market,'ZZ') market,file_exists,zero_partial_flag
            FROM file_chk t1, 
                (SELECT rnum 
                FROM (    
                        SELECT ROWNUM rnum
                        FROM dual
                        Connect BY ROWNUM <= 25 
                        )
                   WHERE rnum>=3) 
        WHERE t1.cust_id(+) = rnum) fc
WHERE 1=1
AND d.market_name=s.market_name(+)
AND d.market_name=s1.market_name(+)
AND d.market_name=fc.market(+)
--AND d.MARKET_NAME<>'DK'
and d.market_name=pr.market(+)
--AND d.MARKET_NAME IN ('CZ','SK')
and d.market_name NOT IN ('HR','RU','HU')
))
  LOOP
     v_message_body:=v_message_body||'<tr><td nowrap><center>'||i.sl_no||'</center></td><td nowrap>'||i.market_name||'</td><td nowrap><center>'||i.plan_date||'</center></td><td nowrap><center>'||i.demand_date||'</center></td><b>'||i.status||'</b><td nowrap><center>'||i.slo_time||'</center></td><td nowrap><center>'||i.plan_publish_time||'</center><td nowrap><center>'||i.me_slo_time||'</center><td nowrap><center>'||i.completion_time||'</center></td><b>'||i.me_status||'</b><td nowrap><center>'||i.hist_file_time||'</center></td><td nowrap><center>'||i.fcst_file_time||'</center></td><td nowrap><center>'||i.comments||'</center></td></tr>'||utl_tcp.CRLF;
  END LOOP;
  

  v_message_body:=v_message_body||'</table></body></html>'||utl_tcp.CRLF;
  
  
  
--  v_message_body := v_message_body||'<p style="font-family:Calibri; font-size:95%;">'||'Note: We have marked SLO Time as * for CZ,SK markets as SLO Times have not been finalized yet.'||'</p>';
  
  --v_message_body := v_message_body||'<p style="font-family:Calibri; font-size:95%;">'||'Please find the attached Operational Exception Report.'||'</p>';
 -- v_message_body := v_message_body||'<p>'||'</p>';
  v_message_body := v_message_body||'<p style="font-family:Calibri;color:black;font-size:95%;">'||'Please find the attached Operational Exception Report.'||'<p>'||'</p>'||'Thanks '||CHR(38)||' Regards,<br>Application Operations and Support'||'</p>';
  --v_message_body := v_message_body||'<style="font-family:Calibri; font-size:95%;">'||'HAVI Core L2 Support';
  --v_message_body := v_message_body||'HAVI Core L2 Support';

  --dbms_output.put_line(v_message_body);
 
  -- Open the SMTP connection ...
    conn := utl_smtp.open_connection(v_smtp_server,n_smtp_server_port);
    dbms_output.put_line('****Init0****');
    utl_smtp.helo(conn, v_smtp_server);
	dbms_output.put_line('****Init1****');
    utl_smtp.mail(conn, v_from_name);
    --utl_smtp.rcpt(conn, v_to_name);
	
	FOR rec_mail IN (
			  SELECT regexp_substr(v_to_name,'[^;]+', 1, LEVEL) mailid
				FROM dual
			  CONNECT BY regexp_substr(v_to_name, '[^;]+', 1, LEVEL) IS NOT NULL
			)
	LOOP
			dbms_output.put_line('Mail sent to : '||rec_mail.mailid);
			UTL_SMTP.rcpt(conn, rec_mail.mailid);
	END LOOP;

 dbms_output.put_line('****1****');
  -- Open data
    utl_smtp.open_data(conn);
 dbms_output.put_line('****2****');
  -- Message info
    utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('To: ' || v_to_name || v_crlf));
    --utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('Date: ' || to_char(sysdate+5/24, 'Dy, DD Mon YYYY hh24:mi:ss') || v_crlf));
    utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('From: ' || v_from_name || v_crlf));
    utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('Subject: ' || v_subject || v_crlf));
    utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('MIME-Version: 1.0' || v_crlf));
    utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('Content-Type: multipart/mixed; boundary="SECBOUND"' || v_crlf || v_crlf));
 dbms_output.put_line('****3****');
  -- Message body
    utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('--SECBOUND' || v_crlf));
    utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('Content-Type: ' || v_message_type || v_crlf || v_crlf));
    utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw(v_message_body || v_crlf));
 dbms_output.put_line('****4****');
  -- Attachment Part
    FOR i IN attachments.FIRST .. attachments.LAST
    LOOP
    -- Attach info
        utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('--SECBOUND' || v_crlf));
        utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('Content-Type: ' || attachments(i).data_type
                            || ' name="'|| attachments(i).attach_name || '"' || v_crlf));
        utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('Content-Disposition: attachment; filename="'
                            || attachments(i).attach_name || '"' || v_crlf || v_crlf));
 
    -- Attach body
        n_offset := 1;
        WHILE n_offset < dbms_lob.getlength(attachments(i).attach_content)
        LOOP
            utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw(dbms_lob.substr(attachments(i).attach_content, n_amount, n_offset)));
            n_offset := n_offset + n_amount;
        END LOOP;
        utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('' || v_crlf));
    END LOOP;
 dbms_output.put_line('****5****');
  -- Last boundry
    utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('--SECBOUND--' || v_crlf));
 dbms_output.put_line('****6****');
  -- Close data
    utl_smtp.close_data(conn);
    utl_smtp.quit(conn);
	
dbms_output.put_line('****End 2****');	
  -- Display the document to browser.
  --demoDocument.displayDocument;

  -- Write document to a file
  -- Assuming UTL file setting are setup in your DB Instance.
  --  
  -- documentArray := demoDocument.getDocumentData;

   -- Use command CREATE DIRECTORY FOO as '<pick a location on your machine>'
   -- to create a directory for the file.

   --v_file := UTL_FILE.fopen('FOO','ExcelObjectTest.xml','W',4000);

   --FOR x IN 1 .. documentArray.COUNT LOOP
  
    -- UTL_FILE.put_line(v_file,documentArray(x));
    
  -- END LOOP;

  -- UTL_FILE.fclose(v_file);  

EXCEPTION
  WHEN OTHERS THEN
      /* For displaying web based error.*/
      htp.p(sqlerrm);
      /* For displaying command line error */
      dbms_output.put_line(sqlerrm);
	  DBMS_OUTPUT.PUT_LINE('In other exception ::: '||DBMS_UTILITY.FORMAT_ERROR_STACK || '@' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE);

 END;
/