<?php
      
$erp_db_host = "topprod";
$erp_db_user = "vd11";
$erp_db_pass = "vd11";  
$erp_conn = oci_connect($erp_db_user, $erp_db_pass, $erp_db_host,'AL32UTF8'); 


$s= "select occ01 from occ_file where substr(occ01,1,4)='E121' ";
$erp_sql = oci_parse($erp_conn,$s );
oci_execute($erp_sql);  
while ($row4 = oci_fetch_array($erp_sql, OCI_ASSOC)) {
  $s="insert into tc_cus_file(tc_cus001,tc_cus002, tc_cus003, tc_cus004) values('".$row4['OCC01'] . "','1','1','固定要求1')";
  $erp_sql = oci_parse($erp_conn,$s );
  oci_execute($erp_sql);    
  
  
  
}    
          
          
?>