<?php
  session_start();
  $pagetitle = "財務部 &raquo; 上個月有期末 本月無期初";
  include("_data.php");
  auth("erp_casesnonzero.php");  
  
  if (is_null($_GET['yearmm'])) {
    $yearmm = date("Y-m") ;     
  } else {      
    $yearmm = $_GET['yearmm'] ;   
  }  
  
  $eyear=intval(substr($yearmm,0,4))  ;
  $emm=intval(substr($yearmm,5,2));
  if ($emm==1) {
      $bmm=12;
      $byear=$eyear-1;      
  } else {
      $bmm=$emm-1;  
      $byear=$eyear;     
  }
  
  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>以下為上月有期末 本月無期初 且無開帳為0 清單 </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
        結帳年月:   
        <input name="yearmm" type="text" id="yearmm" onfocus="WdatePicker({dateFmt:'yyyy-MM'})" value="<?=$yearmm;?>"> &nbsp;&nbsp;  
        <input type="submit" name="submit" id="submit" value="查詢">  &nbsp;&nbsp; 
      </td></tr>
    </table>
  </div>
</form>  
          
<? if (is_null($_GET['submit'])) die ; ?>
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
        <th width="16">&nbsp;</th>
        <th>工單號</th>   
        <th style="text-align:right">金額</th>        
    </tr>
    <?
      $total=0;
      $s2 = "select ccg01, ccg92 from ccg_file              where ccg02=$byear and ccg03=$bmm and ccg92<>0 ".
            "and ccg01 not in (select ccg01 from  ccg_file  where ccg02=$eyear and ccg03=$emm) ".
            "and ccg01 not in (select ccf01 from ccf_file   where ccf02=$byear and ccf03=$bmm) ";
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);  
      $bgkleur = "ffffff";             
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {     
          $total+=$row2['CCG92'];
    ?>
	        <tr bgcolor="#<?=$bgkleur;?>">
		          <td><img src="i/arrow.gif" width="16" height="16">  
              <td><?=$row2["CCG01"];?></td>   
		          <td style="text-align:right"><?=number_format($row2["CCG92"], 2);?></td>
          </tr>
		  <?
			}
      ?> 
      <tr bgcolor="#<?=$bgkleur;?>">
          <td><img src="i/arrow.gif" width="16" height="16">  
          <td>合計</td>   
          <td style="text-align:right"><?=number_format($total, 2);?></td>
      </tr>
            
</table>    