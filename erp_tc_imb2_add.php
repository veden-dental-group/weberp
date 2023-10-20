<?php
  session_start();
  $pagetitle = "資材部 &raquo; 新增料號";
  include("_data.php");
  //auth("erp_tc_imb2_add.php");
  if (is_null($_GET['code'])) {
      $code='';
  } else {
      $code=$_GET['code'];
  }
  
  if (is_null($_GET['name'])) {
      $name='';
  } else {
      $name=$_GET['name'];
  }

  if ($_POST["action"] == "save") {
      $queryus = "insert into erp_tc_imb2 ( code, name, ldate, lotno ) values (  
				          '" . safetext($_POST["code"])       . "',   
                  '" . safetext($_POST["name"])       . "',
                  '" . safetext($_POST["ldate"])      . "',   
                  '" . safetext($_POST["lotno"])      . "' )";     
	  $resultus = mysql_query($queryus) or die ('17 erp_tc_imb2 added error!!.' .mysql_error());
	  forward('erp_tc_imb2.php');
  }
  $showodata = "請輸入以下資料以新增料號.";
  
  //for xajax
  require ('xajax/xajax_core/xajax.inc.php');
  $xajax = new xajax();
  $xajax->configure('javascript URI', 'xajax/');

  Function checkCode($v1,$v2){      
      $objResponse=new xajaxResponse();    
      if ( $v1 != "" ) {  // validate if input 
          $query="select name from erp_tc_imb2 where code = '$v1' and ldate=$v2 limit 1"; 
          $result=mysql_query($query); 
          if (!$result) {
             $objResponse->alert("42 erp_tc_imib2 error!!");   
          } else {
              if ( mysql_num_rows($result)>0) {
                  $row = mysql_fetch_array($result);
                  $objResponse->assign("codestatus","innerHTML", "警告, 料號已存在!!");    
              }else{
                  $objResponse->assign("codestatus","innerHTML", "");   
              }
          }  
      }          
      return $objResponse;
  } 

  $xajax->register(XAJAX_FUNCTION,'checkCode');       
  $xajax->processRequest();

  echo '<?xml version="1.0" encoding="UTF-8"?>'; 
  $IsAjax = True;  
                       
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p><?=$showodata;?> </p>
<form name="form1" method="post" action="<?=$PHP_SELF;?>" >
  <table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">  
    <tr>
      <td bgcolor="#FF66FF" class="witbold">日期: </td>
      <td><input name="ldate" type="text" id="ldate" size="12" maxlength="12" value=<?=date('Y-m-d');?> onfocus="new WdatePicker()"> </td>
    </tr>  
    <tr>
      <td bgcolor="#FF66FF" class="witbold">料號: *</td>
      <td><input name="code" type="text" id="code" size="15" maxlength="15" value="<?=$code;?>"
          onKeyDown="xajax_checkCode(document.getElementById('code').value, document.getElementById('ldate').value)">&nbsp;&nbsp;
          <span style="color:red" id="codestatus"></span>
      </td>
    </tr>         
    <tr>
      <td bgcolor="#FF66FF" class="witbold">名稱: </td>
      <td><input name="name" type="text" id="name" size="50" maxlength="50" value="<?=$name;?>" ></td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF" class="witbold">Lot No.: </td>
      <td><input name="lotno" type="text" id="lotno" size="20" maxlength="20"></td>
    </tr>  
    <tr>
      <td bgcolor="#FF66FF">&nbsp;</td>
      <td><input type="hidden" name="action" value="save">
          <input type="submit" name="Submit" value="送出">          
      </td>
    </tr>
  </table>
</form>
