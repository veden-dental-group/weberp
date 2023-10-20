<?php    
session_start();        
include("_data.php");
$msg='';    

      $query = "select a from log3 ";                                                                                                  
      $result = mysql_query($query) or die ('37 ctuerror error!!');   
      $result=mysql_query($query);   
      while ($row= mysql_fetch_array($result)) {
        $s=$row['a'];  
        if (strpos($s,'register')>0) {
            $s1=strpos($s,'D:\\');
            $e1=strpos($s,'.xls');
            $len1=$e1-$s1+4;
            $sfile=substr($s,$s1,$len1);   
            
            $e2=strrpos($s,'\\');
            $len2=$e1-$e2+3;
            $efile='d:\\ctu\\'.substr($s,$e2+1,$len2);      
             
            if(!copy($sfile, $efile)) $msg.="$sfile error -- ";     
                                    
        }
      
      }

$msg.='轉檔完畢'; 
msg($msg);

?>
