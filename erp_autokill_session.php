<?

//刪除沒有活動的ERPAP 連結
session_start();   
$erp_db_host1 = "topprod";
$erp_db_user1 = "sys";
$erp_db_pass1 = "sys";  
$erp_conn1 = oci_connect($erp_db_user1, $erp_db_pass1, $erp_db_host1,'AL32UTF8', 2); 

          
$s1= "select sid , serial# from v" . "$" . "session  " .
     "where status='INACTIVE' AND PROGRAM ='httpd.exe' " .
     "and machine='WORKGROUP" . "\\" . "ERPAP' " .
     "order by status, serial# ";   
$erp_sql1 = oci_parse($erp_conn1,$s1 );  
oci_execute($erp_sql1);  
while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
    $sid=$row1['SID'];
    $serial=$row1['SERIAL#'];
    
    $s2="alter system kill session '" . $sid . "," . $serial . "' ";
    $erp_sql2 = oci_parse($erp_conn1,$s2 );  
    oci_execute($erp_sql2);  
}

?>