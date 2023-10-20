<?php
  session_start();
  $pagetitle = "行政部 &raquo; 刷卡記錄匯出";
  include("_data.php");
  include("_erp.php");
 // auth("erp_presentout.php");  
  
  if (is_null($_GET['bdate'])) {
    $bdate =  date('Y-m-d');
  } else {
    $bdate=$_GET['bdate'];
  }                              
  if (is_null($_GET['range'])) {
    $range = "10:00:00";
  } else {
    $range = $_GET['range'];
  }                   
  
  if ($_GET["submit"]=="匯出") { 
        
      if ($range=="10:00:00") {
        $filename=$bdate ."-10.txt";  
        
      } else {
        $filename=$bdate ."all.txt";           
      }
      $btime=$bdate . " 00:00:00";
      $etime=$bdate . " " . $range;
      header("Content-type: application/text");
      header("Content-Disposition: attachment; filename=$filename");
      $s1="select deviceserno, rectime, cardserno from accessdata " .
          "where recstatus='' and deviceserno in ('8600101124','8600101070','8400000760','8400000788','8400000759', " .
          "'8400000285','8610000724','8400000768','8400000784','8400000785','8400000313','8400000771','8400000752', " .
          "'8610000677','8610000676') ".
          "and rectime between '$btime' and '$etime' " .
          "order by deviceserno";     
      
      $odbc_conn = odbc_connect('present','sa','');      
      $odbc_rs= odbc_exec( $odbc_conn, $s1);     
      while (odbc_fetch_row($odbc_rs)) {
        $deviceno = substr(trim(odbc_result($odbc_rs, 'DEVICESERNO')),-3);   
        $rectime= trim(odbc_result($odbc_rs, 'RECTIME'));
        $rectime=substr($rectime,0,4). substr($rectime,5,2).substr($rectime,8,2).substr($rectime,11,2).substr($rectime,14,2);
        $cardno= trim(odbc_result($odbc_rs, 'CARDSERNO'));    
        echo $deviceno.$rectime.'010'.$cardno."\r\n";
      }                
      exit;    
  }  

  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>員工每日刷卡資料匯出</p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
         日期:   
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()">  &nbsp;&nbsp; 
         範圍
        <input name="range" type="radio" value="10:00:00" id="range1" <? if($range=='10:00:00') echo " checked";?>><label for="range1" >早上</label>&nbsp;&nbsp;
        <input name="range" type="radio" value="23:59:59" id="range2" <? if($range=='23:59:59') echo " checked";?>><label for="range2" >全天 </label> &nbsp;&nbsp;&nbsp;&nbsp;     
        <input type="submit" name="submit" id="submit" value="匯出">  &nbsp;&nbsp;   &nbsp;&nbsp;            
      </td></tr>
    </table>
  </div>
</form>
