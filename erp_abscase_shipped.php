<?php
include("_data.php");      
$odbc_db_host = "abs";
$odbc_db_user = "veden";
$odbc_db_pass = "veden";  
$odbc_conn = odbc_connect('abs','veden','veden');  

date_default_timezone_set('Asia/Taipei');  

$tdate=date('Y-m-d', strtotime("-1 days")); 
$tdate='12/05/17';
$s1="select ogaud02, ima1002, metalno, metal, imaud04, ima021, tc_dex004, ogb12  from " .
      "( select oga01, ogb03, ogaud02, ta_oea003, ta_oea001, ima1002, ogb12 " .
      "  from oga_file, ogb_file, ima_file, oea_file where oga16=oea01 and ogb04=ima01 and oga01=ogb01 and oga04 like 'U132%' and oga02=to_date('$tdate','yy/mm/dd')) " .
    "left join " .
      "(select tc_dex001, tc_dex002, tc_dex003 metalno, ima1002 metal, imaud04, ima021, tc_dex004 from tc_dex_file, ima_file where tc_dex003=ima01 and imaud10 is not null ) ".    
    "on oga01=tc_dex001 and ogb03=tc_dex002 " .
    "order by ogaud02 ";

$erp_sql1 = oci_parse($erp_conn,$s1 );
oci_execute($erp_sql1);  
while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
  $casenumber   =$row1['OGAUD02']; 
  $metalno      =$row1['METALNO'];  
  $imaud04      =$row1['IMAUD04'];  
  $tc_dex004    =$row1['TC_DEX004'];  
  $vstatus      ='999';            
  $vdateupdated =date('Y-m-d H:i:s');      
  
  //不用檢查是否已出 過貨
  //$queryu = "select vcasenumber from erp_abs_odbc_updatedrecord where vcasenumber='" . $casenumber . "' and vstatus='" . $vstatus . "' limit 1";
  //$resultu= mysql_query($queryu) ;
  //if (mysql_num_rows($resultu) == 0) {                                                                   
    $odbc_sql="update ABSCase set VDROPSHIP='0', VCASESTATUS='$vstatus', VDATEUPDATED='$vdateupdated' " .
              "where CASENUMBER='$casenumber' and status=10 ";       
    $isupdated=$odbc_rs = odbc_exec( $odbc_conn, $odbc_sql);  
    
    if (!is_null($metalno)) {  //有金屬重量才更新ABS table
        $odbc_sql="update ABSCaseProducts set VINVOICEQTY='$tc_dex004' " .
                  "where ABSProductId='$imaud04' AND ABSCaseId_fk = (SELECT Id FROM ABSCase where CASENUMBER='$casenumber')";       
        $isupdated=$odbc_rs = odbc_exec( $odbc_conn, $odbc_sql);  
    } else {
        $tc_dex004=0;
        $metalno='zz';
        $imaud04='zz';
    }
    
    $queryap   = "insert into erp_abs_odbc_updatedrecord ( vcasenumber, vdropship, vstatus, vdateupdated, metalno, imaud04, metalqty, isupdated ) values ( 
                            '" . $casenumber    . "',  
                            '0',  
                            '" . $vstatus       . "',  
                            '" . $vdateupdated  . "',
                            '" . $metalno       . "',  
                            '" . $imaud04       . "',  
                            "  . $tc_dex004     . ",
                            '" . $isupdated     . "' )";
                            
    $resultap= mysql_query($queryap) or die (mysql_error()); 
  //}
}




   
          
          
?>