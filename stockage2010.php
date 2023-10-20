<?php
  session_start();
  $pagetitle = "財務部 &raquo; 庫存帳齡分析";
  include("_data.php");
  
  $q0="delete from stock2010";
  $re0=mysql_query($q0) or die ('7 stock2010 detete error!!' . mysql_error());    
  commit;
  
  $q0="select code, qty from stock2009";
  $re0= mysql_query($q0) or die ('3 pilot2010in error!!' . mysql_error());     
  while ($r0 = mysql_fetch_array($re0)) {
      $code=$r0['code'];  
      $qty=$r0['qty'];  
      $q1="insert into stock2010 (countdate, code, name, unit, date1, date2, qty1, qty2, qty3, qty4, qty5, qty6, flag1) ".  
            "values ('2010/12/31','$code','','','2010/07/01','2010/01/01',$qty,0,0,0,0,0,'0')";
      $re1=mysql_query($q1) or die ('15 stock2010 added error!!' . mysql_error());  
  }
  commit;
  
  
  //201001-06進貨
  $q1="select b, sum(g) g from sheet201001 group by b" ;  
  $re1= mysql_query($q1) or die ('3 pilot2010in error!!' . mysql_error());     
  while ($r1 = mysql_fetch_array($re1)) {
    $code=$r1['b'];   
    $qty=floatval($r1['g']);   
    
    $q2="select pkey from stock2010 where code='$code'";
    $re2=mysql_query($q2) or die ('14 stock2010 error!!' . mysql_error());    
    if (mysql_num_rows($re2) == 0) {
        $q3="insert into stock2010 (countdate, code, name, unit, date1, date2, qty1, qty2, qty3, qty4, qty5, qty6, flag1) ".  
            "values ('2010/12/31','$code','','','2010/07/01','2010/01/01',0,0,0,0,0,0,'1')";
        $re3=mysql_query($q3) or die ('15 stock2010 added error!!' . mysql_error());      
    }  
    commit;
    
    $q4="update stock2010 set qty1=qty1+$qty where code='$code' limit 1";
    $re4=mysql_query($q4) or die ('23 stock2010 updated error!!' . mysql_error());      

    commit;
  }  
  
  //201007-12進貨
  $q1="select b, sum(g) g from sheet201007 group by b" ;  
  $re1= mysql_query($q1) or die ('3 pilot2010in error!!' . mysql_error());     
  while ($r1 = mysql_fetch_array($re1)) {
    $code=$r1['b'];   
    $qty=floatval($r1['g']);  
    
    $q2="select pkey from stock2010 where code='$code'";
    $re2=mysql_query($q2) or die ('14 stock2010 error!!' . mysql_error());    
    if (mysql_num_rows($re2) == 0) {
        $q3="insert into stock2010 (countdate, code, name, unit, date1, date2, qty1, qty2, qty3, qty4, qty5, qty6, flag1) ".  
            "values ('2010/12/31','$code','','','2010/07/01','2010/01/01',0,0,0,0,0,0,'1')";
        $re3=mysql_query($q3) or die ('15 stock2010 added error!!' . mysql_error());      
    }  
    commit;
    
    $q4="update stock2010 set qty2=qty2+$qty where code='$code' limit 1";
    $re4=mysql_query($q4) or die ('23 stock2010 updated error!!' . mysql_error());      

    commit;
  }  


  //領料
  //201001-06出貨
  $q1="select b, sum(h) h from sheet201001 group by b" ;  
  $re1= mysql_query($q1) or die ('3 pilot2010in error!!' . mysql_error());      
  while ($r1 = mysql_fetch_array($re1)) {
    $code=$r1['b'];    
    $qty=floatval($r1['h']);   
    $q2="select pkey from stock2010 where code='$code'";
    $re2=mysql_query($q2);
    if (mysql_num_rows($re2) == 0) {
        $q3="insert into stock2010 (countdate, code, name, unit, date1, date2, qty1, qty2, qty3, qty4, qty5, qty6,flag1) ".
            "values ('2010/12/31','$code','','','2010/07/01','2010/01/01',0,0,0,0,0,0,'3')";
        $re3=mysql_query($q3) or die ('63 stock2010 added error!!' . mysql_error());      
    }  
    commit;
    $q2="select * from stock2010 where code='$code'";
    $re2=mysql_query($q2);
    $row2=mysql_fetch_array($re2);
    
    $pkey=$row2['pkey'];
    $qty1=$row2['qty1'];
    $qty2=$row2['qty2'];
    $qty3=$row2['qty3'];
    $qty4=$row2['qty4'];
    $qty5=$row2['qty5'];
    $qty6=$row2['qty6'];
    
          if ($qty>=$qty6){
              $qty -= $qty6; 
              $qty6=0;                                      
          } else {
              $qty6-=$qty;
              $qty=0;
          }
          
          if ($qty>=$qty5){              
              $qty -= $qty5; 
              $qty5=0;          
          } else {
              $qty5-=$qty;
              $qty=0;
          }
          
          if ($qty>=$qty4){
              $qty -= $qty4; 
              $qty4=0;          
          } else {
              $qty4-=$qty;
              $qty=0;
          }
          
          if ($qty>=$qty3){  
              $qty -= $qty3; 
              $qty3=0;          
          } else {
              $qty3-=$qty;
              $qty=0;
          }
          
          if ($qty>=$qty2){ 
              $qty -= $qty2; 
              $qty2=0;          
          } else {
              $qty2-=$qty;
              $qty=0;
          }
          
          if ($qty>=$qty1){  
              $qty -= $qty1; 
              $qty1=0;          
          } else {
              $qty1-=$qty;
              $qty=0;
          }
          
          $query6="update stock2010 set qty0=$qty, qty1=$qty1,qty2=$qty2,qty3=$qty3,qty4=$qty4,qty5=$qty5,qty6=$qty6 where pkey='$pkey'"; 
          $result6= mysql_query($query6) or die ('131 Stock2010 update error!!' . mysql_error());  
          commit; 
  }   
  
  
    //領料
  //201007-12出貨
  $q1="select b, sum(h) h from sheet201007 group by b" ;  
  $re1= mysql_query($q1) or die ('3 pilot2010in error!!' . mysql_error());      
  while ($r1 = mysql_fetch_array($re1)) {
    $code=$r1['b'];    
    $qty=floatval($r1['h']);   
    $q2="select pkey from stock2010 where code='$code'";
    $re2=mysql_query($q2);
    if (mysql_num_rows($re2) == 0) {
        $q3="insert into stock2010 (countdate, code, name, unit, date1, date2, qty1, qty2, qty3, qty4, qty5, qty6,flag1) ".
            "values ('2010/12/31','$code','','','2010/07/01','2010/01/01',0,0,0,0,0,0,'3')";
        $re3=mysql_query($q3) or die ('63 stock2010 added error!!' . mysql_error());      
    }  
    commit;
    $q2="select * from stock2010 where code='$code'";
    $re2=mysql_query($q2);
    $row2=mysql_fetch_array($re2);
    
    $pkey=$row2['pkey'];
    $qty1=$row2['qty1'];
    $qty2=$row2['qty2'];
    $qty3=$row2['qty3'];
    $qty4=$row2['qty4'];
    $qty5=$row2['qty5'];
    $qty6=$row2['qty6'];
    
          if ($qty>=$qty6){
              $qty -= $qty6; 
              $qty6=0;                                      
          } else {
              $qty6-=$qty;
              $qty=0;
          }
          
          if ($qty>=$qty5){              
              $qty -= $qty5; 
              $qty5=0;          
          } else {
              $qty5-=$qty;
              $qty=0;
          }
          
          if ($qty>=$qty4){
              $qty -= $qty4; 
              $qty4=0;          
          } else {
              $qty4-=$qty;
              $qty=0;
          }
          
          if ($qty>=$qty3){  
              $qty -= $qty3; 
              $qty3=0;          
          } else {
              $qty3-=$qty;
              $qty=0;
          }
          
          if ($qty>=$qty2){ 
              $qty -= $qty2; 
              $qty2=0;          
          } else {
              $qty2-=$qty;
              $qty=0;
          }
          
          if ($qty>=$qty1){  
              $qty -= $qty1; 
              $qty1=0;          
          } else {
              $qty1-=$qty;
              $qty=0;
          }
          
          $query6="update stock2010 set qty0=$qty, qty1=$qty1,qty2=$qty2,qty3=$qty3,qty4=$qty4,qty5=$qty5,qty6=$qty6 where pkey='$pkey'"; 
          $result6= mysql_query($query6) or die ('131 Stock2010 update error!!' . mysql_error());  
          commit; 
  }   
  

?>
