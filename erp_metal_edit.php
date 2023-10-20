<?
  session_start();
  $pagetitle = "資材部 &raquo; 修改鑄造後重量";   
  include("_data.php");
  //auth("erp_metal_edit.php");  
        //若已結完帳 不能修改     

  if ($_POST["action"] == "save") {
	    $queryul = "update erp_metal_add set        
                  bdate       = '" . safetext($_POST["bdate"])      . "',   
                  metal       = '" . safetext($_POST["metal"])      . "',
                  ticketno    = '" . safetext($_POST["ticketno"])   . "', 
                  weight1     = '" . safetext($_POST["weight1"])    . "',
                  weight2     = '" . safetext($_POST["weight2"])    . "',           
		        	    memo         = '" . safetext($_POST["memo"])      . "'
                  where pkey   = '" . safetext($_POST["pkey"])      . "' limit 1";
	    $resultul = mysql_query($queryul) or die ('16 erp_metal_add updated error!! ' . mysql_error());   
	    msg('更新完畢.');
	    forward("erp_metal.php");    
  }

  $queryu = "select * from erp_metal_add where pkey='" . safetext($_GET["pkey"]) . "' limit 1";
  $resultu= mysql_query($queryu) or die ('28 erp_metal_add error!!');
  if (mysql_num_rows($resultu) > 0) {
	  $rowu = mysql_fetch_array($resultu);
  }else{
	  msg("資料不存在!!");
	  forward("erp_metal.php?bdate=". $_GET['bdate']);
  }

  //for xajax
  require ('xajax/xajax_core/xajax.inc.php');
  $xajax = new xajax();
  $xajax->configure('javascript URI', 'xajax/');


  echo '<?xml version="1.0" encoding="UTF-8"?>'; 
  $IsAjax = True;      
  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>請填入以下資料.</p>
<form name="form1" method="post" action="<?=$PHP_SELF;?>">
  <table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table2">      
    <tr> 
      <td class="title1">日期: *</td>
      <td><input name="bdate"   type="text" id="bdate" size="12" maxlength="12" value="<?=$rowu['bdate'];?>" onfocus="new WdatePicker()"></td>
    </tr>
    <tr>
        <td class="title1"><div align="left">金屬: *</td>
        <td><input name="metal" type="text" id="metal" size="5" value="<?=$rowu['metal'];?>"></td>  
    </tr>
    <tr>
      <td class="title1">工單號: </td>
      <td><input name="ticketno" type="text" id="ticketno" size="25" value="<?=$rowu['ticketno'];?>"></td>                                                 
    </tr>
    <tr>    
    <tr>    
      <td class="title1">鑄造後重量: </td>
      <td><input name="weight1" type="text" id="weight1" size="5" style="text-align:right" value="<?=$rowu['weight1'];?>" onkeypress="return numberOnly(event,'weight1')"></td>          
    </tr>
    <tr>    
      <td class="title1">研磨後重量: </td>
      <td><input name="weight2" type="text" id="weight2" size="5"  style="text-align:right" value="<?=$rowu['weight2'];?>" onkeypress="return numberOnly(event,'weight2')"></td>          
    </tr>   
    <tr>
      <td><input type="hidden" name="action" value="save">                                 
          <input type="hidden" name="pkey" value="<?=safetext($_GET["pkey"]);?>">  
          <input type="submit" name="Submit" value="更新"></td>  
    </tr>
  </table>
</form>     