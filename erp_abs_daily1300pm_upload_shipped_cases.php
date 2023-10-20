<?php
include("_data.php");      
$odbc_db_host = "abs";
$odbc_db_user = "veden";
$odbc_db_pass = "veden";  
$odbc_conn1 = odbc_connect('abs','veden','veden');  

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
$edate=date('Y-m-d');
//$edate='2012-02-08';
//$edate='2012-05-02';
$s1= "select ogaud02,oga16 from oga_file where oga02=to_date('$edate','yy/mm/dd') and oga04 like 'U132%' order by ogaud02";

$erp_sql1 = oci_parse($erp_conn1,$s1 );
oci_execute($erp_sql1);  
while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {     
                          
    $casenumber=$row1['OGAUD02']; 
    $vstatus='999';
    $vdateupdated=date('Y-m-d H:i:s');
    $vtrackingnum=$row1['OGA16'];
    $odbc_sql="update ABSCase set VDROPSHIP='0', VCASESTATUS='$vstatus', VDATEUPDATED='$vdateupdated', VTRACKINGNUM='$vtrackingnum' " .
              "where CASENUMBER='$casenumber' and status=10 ";
    $isupdated=$odbc_rs = odbc_exec( $odbc_conn1, $odbc_sql);  
    
    $queryap   = "insert into erp_abs_odbc_updatedrecord ( vcasenumber, vdropship, vstatus, vdateupdated, vtrackingnum ) values ( 
                            '" . $casenumber   . "',  
                            '0',  
                            '" . $vstatus      . "',  
                            '" . $vdateupdated  . "',      
                            '" . $vtrackingnum . "')";
    $resultap= mysql_query($queryap);   
}       
?>