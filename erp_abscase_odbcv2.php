<?php
include("_data.php");      
$odbc_db_host = "abs";
$odbc_db_user = "veden";
$odbc_db_pass = "veden";  
$odbc_conn = odbc_connect('abs','veden','veden');  

//$odbc_rs= odbc_exec( $odbc_conn, "select * from ABSDataShare" );  
//while (odbc_fetch_row($odbc_rs)) {
//    $id            = odbc_result($odbc_rs, 'ID'); 
//    $casenumber    = odbc_result($odbc_rs, 'CASENUMBER');    
//    $clientcaseid  = odbc_result($odbc_rs, 'CLIENTCASEID');    
//    $casedate      = odbc_result($odbc_rs, 'CASEDATE');    
//    $caseshipdate  = odbc_result($odbc_rs, 'CASESHIPDATE');    
//    $patientname   = odbc_result($odbc_rs, 'PATIENTNAME');    
//    $status        = odbc_result($odbc_rs, 'STATUS');    
//    $dropship      = odbc_result($odbc_rs, 'DROPSHIP');    
//    $vdropship     = odbc_result($odbc_rs, 'VDROPSHIP');    
//    $vstatus       = odbc_result($odbc_rs, 'VSTATUS');    
//    $vdateupdated  = odbc_result($odbc_rs, 'VDATEUPDATED');    
//    $vtrackingnum  = odbc_result($odbc_rs, 'VTRACKINGNUM');     
//} 



//由erp_conn1中取出全部的Order
//vd110中的只有送貨客戶編號
//由tc_srg_file 去讀 sfb_file 再去讀 oea_file 才能取出PLS的全部order
//ecb_file 存有產品/順序/工序代號   ecb01:品代  ecb03:順序  ecb06:工序代號
date_default_timezone_set('Asia/Taipei');  
$s1= "select tc_srg002, ecb06, oea01 " .
     "from (select tc_srg001, tc_srg002, tc_srg003, tc_srg004 from tc_srg_file where tc_srg005 is not null ), " .
     " ecb_file, sfb_file, " .
     " (select oea01 from oea_file WHERE substr(oea04,1,4)='U132' and  not EXISTS (SELECT 1 FROM oga_file  WHERE oea01=oga16)) " .    
     "where tc_srg003=ecb01 and tc_srg004=ecb03 and sfb28 is null and  tc_srg001=sfb01 and sfb22=oea01 ";

$erp_sql1 = oci_parse($erp_conn,$s1 );
oci_execute($erp_sql1);  
while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
  $casenumber=$row1['TC_SRG002']; 
  $vstatus=$row1['ECB06'];
  $vdateupdated=date('Y-m-d H:i:s');
  $vtrackingnum=$row1['OEA01'];
  // 工序不變則不寫
  $queryu = "select vcasenumber from erp_abs_odbc_updatedrecord where vcasenumber='" . $casenumber . "' and vstatus='" . $vstatus . "' limit 1";
  $resultu= mysql_query($queryu) ;
  if (mysql_num_rows($resultu) == 0) {     
    //$odbc_sql= "insert into ABSCase (CASENUMBER, VDROPSHIP, VCASESTATUS, VDATEUPDATED, VTRACKINGNUM)
    //            values('$casenumber','0','$vstatus','$vdateupdated','$vtrackingnum')";
    $odbc_sql="update ABSCase set VDROPSHIP='0', VCASESTATUS='$vstatus', VDATEUPDATED='$vdateupdated', VTRACKINGNUM='$vtrackingnum' " .
              "where CASENUMBER='$casenumber' ";
    
    $isupdated=$odbc_rs = odbc_exec( $odbc_conn, $odbc_sql);  
    
    $queryap   = "insert into erp_abs_odbc_updatedrecord ( vcasenumber, vdropship, vstatus, vdateupdated, vtrackingnum ) values ( 
                            '" . $casenumber   . "',  
                            '0',  
                            '" . $vstatus      . "',  
                            '" . $vdateupdated  . "',      
                            '" . $vtrackingnum . "')";
    $resultap= mysql_query($queryap); 
  }
}




   
          
          
?>