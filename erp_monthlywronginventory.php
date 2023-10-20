<?php
  session_start();
  $pagetitle = "資材部 &raquo; 每月庫存<0檢查";
  include("_data.php");
  auth("erp_monthlywronginventory.php");  
  
  if (is_null($_GET['thismonth'])) {
    $thismonth = date("Y-m",strtotime("-1 month")) ;     
  } else {      
    $thismonth = $_GET['thismonth'] ;   
  }                    
  
  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>以下為每日庫存<0的清單!! </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
        月份:   
        <input name="thismonth" type="text" id="thismonth" onfocus="WdatePicker({dateFmt:'yyyy-MM'})" value="<?=$thismonth;?>"> &nbsp;&nbsp;  
        <input type="submit" name="submit" id="submit" value="查詢">  &nbsp;&nbsp;     
      </td></tr>
    </table>
  </div>
</form>  
          
<? if (is_null($_GET['submit'])) die ; ?>
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
        <th width="16">&nbsp;</th>
        <th>物料編號</th>  
        <th>物料名稱</th>    
        <th style="text-align:right">月底庫存</th>   
    </tr>
    <?
    
      $yyyy=substr($thismonth,0,4);
      $mm=substr($thismonth,5,2);
  
      $s2="select imk01, ima02, imk09 from imk_file, ima_file where imk09<0 and imk05=$yyyy and imk06=$mm and imk01=ima01 ";
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);  
      $bgkleur = "ffffff";  
      $i=0;
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) { 
          $i++;
      ?>
	        <tr bgcolor="#<?=$bgkleur;?>">
		          <td><?=$i;?></td>
              <td><?=$row2["IMK01"];?></td>  
              <td><?=$row2["IMA02"];?></td>  
              <td><?=$row2["IMK09"];?></td>   
          </tr>
		  <?
			}  
      ?>  
      <tr><td colspan="4">以下無資料</td></tr> 
</table>   