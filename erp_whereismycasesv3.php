<?php
  session_start();
  $pagetitle = "業務部 &raquo; Where is the CASE";
  include("_data.php");
  include("_erpv3.php");
  auth("erp_whereismycases.php");

$bdate = $_GET['bdate'] ? $_GET['bdate'] : date('Y-m-d', strtotime("-60 days"));

  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>WHERE IS MY CASE!! </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
		<td bgcolor="#eeeeee">
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value="<?=$bdate;?>" onfocus="new WdatePicker()">之後
        </td>
		<td bgcolor="#eeeeee">
         RX #:   
        <input name="rxno" type="text" id="rxno" size="70" value="<?=$_GET['rxno'];?>">    (若有多個 RX#  請用 , 隔開 )           
        <input type="submit" name="submit" id="submit" value="查詢">  &nbsp;&nbsp;   &nbsp;&nbsp;          
        </td>
      </tr>
    </table>
  </div>
</form>
<? if (is_null($_GET['submit'])) die ; ?>   
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">   
      <?
        $rxarray=explode(',', $_GET['rxno']);
        $max=count($rxarray);
        $msg='';
        for($i=0; $i<$max; $i++){
            $rxno=$rxarray[$i]; 
            $msg.=findcasewithrxno($rxno,$erp_conn1,$erp_conn2, $bdate);
        }   
      ?>
      <tr bgcolor="#ffffff">
        <td><?=$msg;?></td>      
      </tr>  
</table>   
