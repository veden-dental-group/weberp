<?php
      
$erp_db_host = "topprod";
$erp_db_user = "vd11";
$erp_db_pass = "vd11";  
$erp_conn = oci_connect($erp_db_user, $erp_db_pass, $erp_db_host,'AL32UTF8'); 


$s1= "select occ01 from occ_file where substr(occ01,1,4)='E121' ";
$erp_sql1 = oci_parse($erp_conn,$s1 );
oci_execute($erp_sql1);  
while ($row4 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
  $s="insert into tc_cus_file(tc_cus001,tc_cus002, tc_cus003, tc_cus004) values('".$row4['OCC01'] . "','1','1','咬合咬咬到，1張')";
  $erp_sql = oci_parse($erp_conn,$s );
  oci_execute($erp_sql);    
  $s="insert into tc_cus_file(tc_cus001,tc_cus002, tc_cus003, tc_cus004) values('".$row4['OCC01'] . "','1','2','輕咬，2張')";
  $erp_sql = oci_parse($erp_conn,$s );
  oci_execute($erp_sql); 
  $s="insert into tc_cus_file(tc_cus001,tc_cus002, tc_cus003, tc_cus004) values('".$row4['OCC01'] . "','1','3','有附（）咬合器底座，請小心保管，退回')";
  $erp_sql = oci_parse($erp_conn,$s );
  oci_execute($erp_sql);    
  $s="insert into tc_cus_file(tc_cus001,tc_cus002, tc_cus003, tc_cus004) values('".$row4['OCC01'] . "','1','4','請去倉庫領取（）個咬合器底座，出貨前請拆掉')";
  $erp_sql = oci_parse($erp_conn,$s );
  oci_execute($erp_sql); 
  $s="insert into tc_cus_file(tc_cus001,tc_cus002, tc_cus003, tc_cus004) values('".$row4['OCC01'] . "','1','5','接觸點做成點接觸')";
  $erp_sql = oci_parse($erp_conn,$s );
  oci_execute($erp_sql); 
  $s="insert into tc_cus_file(tc_cus001,tc_cus002, tc_cus003, tc_cus004) values('".$row4['OCC01'] . "','1','6','接觸點做成面接觸，從咬合面開始完全接觸')";
  $erp_sql = oci_parse($erp_conn,$s );
  oci_execute($erp_sql); 
  $s="insert into tc_cus_file(tc_cus001,tc_cus002, tc_cus003, tc_cus004) values('".$row4['OCC01'] . "','1','7','雙層冠第二 層，整顆都用玻璃牙覆蓋')";
  $erp_sql = oci_parse($erp_conn,$s );
  oci_execute($erp_sql);
  $s="insert into tc_cus_file(tc_cus001,tc_cus002, tc_cus003, tc_cus004) values('".$row4['OCC01'] . "','1','8','雙層冠第二層做咬金，其他地方請用玻璃牙覆蓋')";
  $erp_sql = oci_parse($erp_conn,$s );
  oci_execute($erp_sql);
  $s="insert into tc_cus_file(tc_cus001,tc_cus002, tc_cus003, tc_cus004) values('".$row4['OCC01'] . "','1','9','雙層冠第二層做金屬冠')";
  $erp_sql = oci_parse($erp_conn,$s );
  oci_execute($erp_sql);
  $s="insert into tc_cus_file(tc_cus001,tc_cus002, tc_cus003, tc_cus004) values('".$row4['OCC01'] . "','1','10','雙層冠第二層請預留焊點')";
  $erp_sql = oci_parse($erp_conn,$s );
  oci_execute($erp_sql);
  $s="insert into tc_cus_file(tc_cus001,tc_cus002, tc_cus003, tc_cus004) values('".$row4['OCC01'] . "','1','11','咬合面染色')";
  $erp_sql = oci_parse($erp_conn,$s );
  oci_execute($erp_sql);
  $s="insert into tc_cus_file(tc_cus001,tc_cus002, tc_cus003, tc_cus004) values('".$row4['OCC01'] . "','1','12','OP前噴沙，用壓力6公斤噴，必須涂金膜油')";
  $erp_sql = oci_parse($erp_conn,$s );
  oci_execute($erp_sql);
  $s="insert into tc_cus_file(tc_cus001,tc_cus002, tc_cus003, tc_cus004) values('".$row4['OCC01'] . "','1','13','激光植PIN，白底黃模')";
  $erp_sql = oci_parse($erp_conn,$s );
  oci_execute($erp_sql);
  $s="insert into tc_cus_file(tc_cus001,tc_cus002, tc_cus003, tc_cus004) values('".$row4['OCC01'] . "','1','14','固定做好后請儘快送活動部門製作')";
  $erp_sql = oci_parse($erp_conn,$s );
  oci_execute($erp_sql);
                         
  
}    
          
          
?>