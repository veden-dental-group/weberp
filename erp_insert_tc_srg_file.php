<?
  session_start();
  //$pagtitle = "廠務部 &raquo; 期間內未出貨工單"; 
  $erp_db_host1 = "topprod";
  $erp_db_user1 = "vd110";
  $erp_db_pass1 = "vd110";  
  $erp_conn1 = oci_connect($erp_db_user1, $erp_db_pass1, $erp_db_host1,'AL32UTF8'); 

    $s1=  "select distinct(tc_srg001), tc_srg003, sfbud02, to_char(tc_srgcrat,'yy/mm/dd') tc_srgcrat, tc_srguser, tc_srggrup, tc_srgoriu, tc_srgorig  from tc_srg_file, sfb_file " .
          "where tc_srg001>'A311-1208025500' and tc_srg001=sfb01 " .
          "and  tc_srg001 not in (select tc_srg001 from tc_srg_file where tc_srg001>'A311-1208025500' and tc_srg004='20' )  " .
          "and tc_srg003||20 in ( select ecb01||ecb03 from ecb_file) ";
    $erp_sql1 = oci_parse($erp_conn1,$s1 );
    oci_execute($erp_sql1);                                
    while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {  
          $tc_srg001=$row1['TC_SRG001'];
          $tc_srg002=$row1['SFBUD02'];
          $tc_srg003=$row1['TC_SRG003'];
          $tc_srgcrat=$row1['TC_SRGCRAT']; 
          $tc_srguser=$row1['TC_SRGUSER']; 
          $tc_srggrup=$row1['TC_SRGGRUP']; 
          $tc_srgoriu=$row1['TC_SRGORIU']; 
          $tc_srgorig=$row1['TC_SRGORIG']; 
          
          $s2="select ecb06, ecb40 from ecb_file where ecb01='$tc_srg003' and ecb03='20' ";
          $erp_sql2 = oci_parse($erp_conn1,$s2 );
          oci_execute($erp_sql2);                                
          $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);    
          $ecb06=$row2['ECB06'];
          $ecb40=$row2['ECB40'];                                                                   
    
          $s3 = "insert into tc_srg_file values ('$tc_srg001','$tc_srg002','$tc_srg003','20', NULL, '$ecb40', NULL, NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N', " .
                "NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'frank','Y', " .
                "to_date('$tc_srgcrat','yy/mm/dd'), NULL, '$tc_srguser', '$tc_srggrup', NULL, '$tc_srgoriu', '$tc_srgorig', '$ecb06')";

          $erp_sql3 = oci_parse($erp_conn1,$s3 );
          oci_execute($erp_sql3);                                
        
      
    }                      
                
