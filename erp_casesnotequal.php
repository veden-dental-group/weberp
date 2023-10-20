<?php
  session_start();
  $pagetitle = "財務部 &raquo; 投入合計 != 總帳錄入";
  include("_data.php");
 // auth("erp_casesnotequal.php");  
  
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
<p>以下為 投入合計 != 總帳錄入 (一般是上月結案工單在本月報工 請刪除報工資料) 清單 </p>     
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
        <th>品代</th>  
        <th>品名</th>  
        <th>類別</th>
        <th style="text-align:right">成本</th>        
        <th style="text-align:right">工單分攤數</th>
        <th style="text-align:right">單位成本</th>   
    </tr>
    <?
      $total1=0;    
      $s2 = "select cdc041, decode(cdc04, '1', '人工', '2','製費一','3', '製費二','unknowm') cdc04, cdc05 , cdc06, cdc07,  sfb05, ima02 " .
            "from cdc_file, sfb_file, ima_file " .
            "where cdc041=sfb01 and sfb05=ima01 and cdc01=$eyear and cdc02=$emm " .
            "and cdc041 not in ( select cch01 from cch_file where cch02=$eyear and cch03=$emm ) order by cdc041, cdc04 ";
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);  
      $bgkleur = "ffffff";             
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {     
          $total1+=$row2['CDC05']; 
    ?>
	        <tr bgcolor="#<?=$bgkleur;?>">
		          <td><img src="i/arrow.gif" width="16" height="16">  
              <td><?=$row2["CDC041"];?></td>   
              <td><?=$row2["SFB05"];?></td> 
              <td><?=$row2["IMA02"];?></td> 
              <td><?=$row2["CDC04"];?></td>    
		          <td style="text-align:right"><?=number_format($row2["CDC05"], 6);?></td>
              <td style="text-align:right"><?=number_format($row2["CDC06"], 6);?></td>  
              <td style="text-align:right"><?=number_format($row2["CDC07"], 6);?></td>                               
          </tr>
		  <?
			}
      ?> 
      <tr bgcolor="#<?=$bgkleur;?>">
          <td><img src="i/arrow.gif" width="16" height="16">  
          <td colspan="4">合計</td>   
          <td style="text-align:right"><?=number_format($total1, 6);?></td>    
          <td colspan="2">&nbsp;</td>                 
      </tr>
            
</table>    