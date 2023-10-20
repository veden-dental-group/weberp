<?php
      
$erp_db_host = "topprod";
$erp_db_user = "vd11";
$erp_db_pass = "vd11";  
$erp_conn = oci_connect($erp_db_user, $erp_db_pass, $erp_db_host,'AL32UTF8'); 


$s1= "select occ01 from occ_file";
$erp_sql1 = oci_parse($erp_conn,$s1 );
oci_execute($erp_sql1);  
while ($row4 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
  $s="insert into tc_occ_file(tc_occ001,tc_occ002, tc_occ003, tc_occacti) values('".$row4['OCC01'] . "','9123','6A4000','Y')";
  $erp_sql = oci_parse($erp_conn,$s );
  oci_execute($erp_sql); 

  
}    
          
          
?>