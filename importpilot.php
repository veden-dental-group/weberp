<?php
session_start();   
include("_data.php");  
date_default_timezone_set('Asia/Taipei'); 
    


$q1= "select * from sheet2011 order by a ";     
$r1= mysql_query($q1);
while ($row2 = mysql_fetch_array($r1)){ 
    $a1=$row2['A'];    
    $b1=$row2['B'];
    $c1=$row2['C'];
    $f1=$row2['F'];
    $i1=$row2['I'];
    $j1=$row2['J'];
    if (strpos($b1,'(')>0) {              
        $x=strpos($b1,':');          //$b1=產品名稱:瓷牙粉(11C34070020)
        $namecode=substr($b1,$x+1);  
        //$namecode=substr($b1,0, strlen($namecode)-1); 
        $y=strrpos($namecode,'(');
        $code=substr($namecode,$y+1, strlen($namecode)-1);
        $name=substr($namecode,0,$y-1);
        $name='';   
        $i=0;
    }
    if (strpos('asa'.$f1,'Total:')>0 && $i==0) {
        $q2 = "insert into pilot2011in ( p1,p2,p3,p4,p5,p7,p8,p9 ) values (  
               '" . safetext($namecode)   . "',
               '2011/01/01',   
               '" . $i1         . "',' ',  
               '" . $code       . "',  
               '" . safetext($name)       . "', 
               '" . $a1         . "','1')";     
        $resultus = mysql_query($q2) or die(mysql_error());   
    }

    
    if (strpos('aa'.$c1,'import')>0) {
        $q3 = "insert into pilot2011in ( p1,p2,p3,p4,p5,p7,p8,p9 ) values (  
               '" . safetext($namecode)   . "',
               '" . $b1         . "',   
               '" . $i1         . "',  
               '" . $j1         . "',  
               '" . $code       . "',  
               '" . safetext($name)       . "',  
               '" . $a1         . "','2')";     
        $resultus = mysql_query($q3) or die(mysql_error());      
        $i++;
    }   
}
?>
