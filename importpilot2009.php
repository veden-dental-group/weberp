<?php
session_start();   
include("_data.php");  
date_default_timezone_set('Asia/Taipei'); 
                                            
$q1= "select * from sheet2009 where K!=0 order by a ";     
$r1= mysql_query($q1);
while ($row2 = mysql_fetch_array($r1)){ 
    $a1=$row2['B'];    
    $b1=$row2['C']; 
    $c1=$row2['C1'];
    $j1=floatval($row2['K']);
    if (strpos($c1,'(')>0) {
        $y=strrpos($c1,'(');
        $code=substr($c1,$y+1);
        $name=substr($code,0,(strlen($code)-1));
        
    }
    if ($a1=='承轉上期') {
        $q2 = "insert into stock2009 ( fullname, code, name, qty ) values (  
               '" . safetext($b1)     . "',   
               '" . $code             . "',  
               '" . $name             . "', 
               '" . $j1               . "')";     
        $resultus = mysql_query($q2) or die(mysql_error());   
    }        
}


?>
