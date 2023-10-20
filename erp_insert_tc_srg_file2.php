<?
  session_start();
  //$pagtitle = "廠務部 &raquo; 期間內未出貨工單"; 
  $erp_db_host1 = "topprod";
  $erp_db_user1 = "vd110";
  $erp_db_pass1 = "vd110";  
  $erp_conn1 = oci_connect($erp_db_user1, $erp_db_pass1, $erp_db_host1,'AL32UTF8'); 

    $s1=  "select tc_srg001  from tc_srg_file where tc_srg005 is not null and tc_srg001 like 'A311-1208%' group by tc_srg001 having count(*) >1 order by tc_srg001 desc ";
    $erp_sql1 = oci_parse($erp_conn1,$s1 );
    oci_execute($erp_sql1);                                
    while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {  
          $tc_srg001=$row1['TC_SRG001'];   
          
          $s2="select count(*) lc from tc_srg_file where tc_srg001='$tc_srg001' and tc_srg004>'20' and tc_srg007 is not null ";
          $erp_sql2 = oci_parse($erp_conn1,$s2 );
          oci_execute($erp_sql2);                                
          $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);    
          $lc=$row2['LC'];
          if ($lc==0) {
              $s4="select tc_srg004 from tc_srg_file where tc_srg001='$tc_srg001' and tc_srg004>'20' and tc_srg005 is not null  ";
              $erp_sql4 = oci_parse($erp_conn1,$s4 );
              oci_execute($erp_sql4);                                
              $row4 = oci_fetch_array($erp_sql4, OCI_ASSOC); 
              $tc_srg004=$row4['TC_SRG004'];
              
              $s3 = "update tc_srg_file set tc_srg005=NULL where tc_srg001='$tc_srg001' and tc_srg004='$tc_srg004'";
              $erp_sql3 = oci_parse($erp_conn1,$s3 );
              oci_execute($erp_sql3);  
            
          } else {
              $s3 = "update tc_srg_file set tc_srg005=NULL where tc_srg001='$tc_srg001' and tc_srg004='20'";
              $erp_sql3 = oci_parse($erp_conn1,$s3 );
              oci_execute($erp_sql3); 
            
          }         
    }                      
                
