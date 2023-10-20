<?php
                                   
  include("_data.php");
  date_default_timezone_set('Asia/Taipei');    
  
  $bdate = date('Y-m-d');    
    
  $query="select occ02, a811+a821+a831 a8 from erp_dailycount where qtype='U' and iotype='OU' and tdate='$bdate' order by a8 desc limit 10 ";
  $result = mysql_query($query) or die ('393 erp_dailycount read error!!' . mysql_error());  
      //以下用來產生JSON格式的資料
  $prefix = '';
  echo "[\n";
  while ( $row = mysql_fetch_array($result ) ) {
    echo $prefix . " {\n";
    echo '  "customer": "' . $row['occ02'] . '",' . "\n";
    echo '  "value": ' . $row['a8'] . ',' . "\n";   
    echo " }";
    $prefix = ",\n";           
   
  }
  echo "\n]";     
  
  
?>
