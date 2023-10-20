<?php
      
include("_data.php");      
      
$erp_db_host = "topprod";
$erp_db_user = "vd110";
$erp_db_pass = "vd110";  
$erp_conn = oci_connect($erp_db_user, $erp_db_pass, $erp_db_host,'AL32UTF8'); 



$s1= "select bmb03, bmbud02 from bmb_file";
$erp_sql1 = oci_parse($erp_conn,$s1 );
oci_execute($erp_sql1);  
$i=0;
$j=0;
while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {  
  
  $query="select code from color_code where cate='" . $row1['BMBUD02'] ."'";
  $result=mysql_query($query, $conn);
  while ($row=mysql_fetch_array($result)){
      $key=substr($row1['BMB03'],0,5) . $row['code'] . substr($row1['BMB03'],7,3)  ; 
      
      $s2="select ima01 from ima_file where ima01='$key'";
      $erp_sql2 = oci_parse($erp_conn,$s2 );
      oci_execute($erp_sql2); 
      $row2= oci_fetch_array($erp_sql2, OCI_ASSOC);
      $j++;
      if (is_null($row2['IMA01'])) {
        $i++;  
         echo $i . "--" .$row1['BMB03'] . "--" . $row1['BMBUD02'] . "--" . $row['code'] ."--沒有轉換料號: " . $key. "<br>" ;
        //$data= $row1['BMB03'] . "--" . $row1['BMBUD02'] . "--" . $row['code'] ."--沒有轉換料號: " . $key  ; 
        
      }
      
  }                    
}    
echo $j++;           
          
?>