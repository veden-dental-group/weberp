<?php    
session_start();        
include("_data.php");
$msg='';    

      $query = "select a from log3 ";                                                                                                  
      $result = mysql_query($query) or die ('37 ctuerror error!!');   
      $result=mysql_query($query);   
      while ($row= mysql_fetch_array($result)) {
        $s=$row['a'];  
        $sfile=$s;                               
            
        $e2=strrpos($s,'\\');  
        $efile='d:\\ctu\\'.substr($s,$e2+1);      
            
        if(!copy($sfile, $efile)) $msg.="$sfile error -- ";     

      }

$msg.='轉檔完畢'; 
msg($msg);

?>
