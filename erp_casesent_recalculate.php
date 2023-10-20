<?php
session_start();   
include("_data.php");  
date_default_timezone_set('Asia/Taipei'); 
//$edate=date('Y-m-d');    
$edate='2012-04-01';   
//先刪除同一日期的所有資料
$query2= "delete from casesent where s01>='$edate' ";     
$result2= mysql_query($query2);

$s2= "select oga02, sfb82, gem02, sfbud02, sfb08, sfb22, sfb221, sfb01, to_char(oea02,'yyyy-mm-dd') oea02,  to_char(ta_oea005,'yyyy-mm-dd') ta_oea005, sfb05, ima02, imaud07, occ01, occ02 " .
         "from oga_file, ogb_file, sfb_file, ima_file, oea_file, gem_file, occ_file " .
         //"where sfb81<=to_date('$edate','yy/mm/dd') and sfb81>=to_date('120201','yy/mm/dd')   " . //某天以前  
         "where oga02>=to_date('$edate','yy/mm/dd') " .     
         "and oga01=ogb01 " .
         "and ogb31=sfb22 " .   
         "and ogb32=sfb221 " .                                           
         "and sfb22=oea01 and oea04=occ01 " .
         "and sfb05=ima01 " .
         "and sfb82=gem01 " .
         "order by sfb82,sfb22,sfb221";

$erp_sql2 = oci_parse($erp_conn1,$s2 );  
oci_execute($erp_sql2);  
                                                   
while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {     
    $queryus = "insert into casesent ( s01,s02,s03,s031,s04,s05,s06,s09,s10,s11,s12,s13,s14,s25,s26 ) values (  
               '" . $edate                      . "',
               '" . safetext($row2['SFBUD02'])  . "',   
               '" . $row2['OGB31']              . "',  
               '" . $row2['OGB32']          . "',  
               '" . $row2['SFB01']          . "',  
               '" . $row2['OEA02']          . "',  
               '" . $row2['TA_OEA005']      . "',  
               '" . $row2['SFB05']          . "',  
               '" . $row2['IMA02']          . "',  
               '" . $row2['OCC01']          . "',  
               '" . safetext($row2['OCC02']). "',  
               '" . $row2['SFB82']          . "',  
               '" . safetext($row2['GEM02']). "',  
               " . $row2['SFB08']           . ",  
               " . $row2['IMAUD07']         . " )";     
    $resultus = mysql_query($queryus);       
    $msg=mysql_error();
    
}  
?>
