<?php

  //為方便移值到其他公司 在此定義各公司使用的資料庫名字
    // $outserver: 境外  $inserver:境內
    $_SESSION['outserver']='vd210';
    $_SESSION['inserver']='vd110';
    
    $erp_db_host = "topprod";
    $erp_db_user = "vd110";
    $erp_db_pass = "vd110";  
    $erp_conn = oci_connect($erp_db_user, $erp_db_pass, $erp_db_host,'AL32UTF8');
    
    $erp_db_host1 = "topprod";
    $erp_db_user1 = "vd110";
    $erp_db_pass1 = "vd110";  
    $erp_conn1 = oci_connect($erp_db_user1, $erp_db_pass1, $erp_db_host1,'AL32UTF8'); 
     
    $erp_db_host2 = "topprod";
    $erp_db_user2 = "vd210";
    $erp_db_pass2 = "vd210";  
    $erp_conn2 = oci_connect($erp_db_user2, $erp_db_pass2, $erp_db_host2,'AL32UTF8'); 
    
    
?>
