<?php
  session_start();
  $pagetitle = "系統設定 &raquo; 新增人員";
  include("_data.php");
  //auth("staff_add.php");
  
  $stafftypes = $_POST['stafftype'];
  $stafftype='';
  foreach ($stafftypes as $v) {
    $stafftype .= $v; 
  }
  
  if ($_POST["action"] == "save") {
      $queryus = "insert into staff ( guid, name, stafftype, memo ) values (
                  '" . uuid()                           . "',   
				          '" . safetext($_POST["name"])         . "',  
                  '" . $stafftype    . "',       
				          '" . safetext($_POST["memo"])         . "')";     
	  $resultus = mysql_query($queryus) or die ('17 Staff added error!!.' .mysql_error());
	  forward('staff.php');
  }
  $showodata = "請輸入以下資料以新增人員.";
  
  //for xajax
  require ('xajax/xajax_core/xajax.inc.php');
  $xajax = new xajax();
  $xajax->configure('javascript URI', 'xajax/');

  Function checkName($v){      
      $objResponse=new xajaxResponse();    
      if ( $v != "" ) {  // validate if input 
          $query="select name from staff where name = '" .$v . "' limit 1"; 
          $result=mysql_query($query); 
          if (!$result) {
             $objResponse->alert("42 Staff error!!");   
          } else {
              if ( mysql_num_rows($result)>0) {
                  $row = mysql_fetch_array($result);
                  $objResponse->assign("namestatus","innerHTML", "警告, 名稱已存在!!");    
              }else{
                  $objResponse->assign("namestatus","innerHTML", "");   
              }
          }  
      }          
      return $objResponse;
  } 

  $xajax->register(XAJAX_FUNCTION,'checkName');       
  $xajax->processRequest();

  echo '<?xml version="1.0" encoding="UTF-8"?>'; 
  $IsAjax = True;  
                       
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p><?=$showodata;?> </p>
<form name="form1" method="post" action="<?=$PHP_SELF;?>" onsubmit="return validStaffAdd()">
  <table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
      <td bgcolor="#FF66FF" class="witbold">名稱: *</td>
      <td><input name="name" type="text" id="name" size="50"
          onKeyDown="xajax_checkName(document.getElementById('name').value)">&nbsp;&nbsp;
          <span style="color:red" id="namestatus"></span>
      </td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF" class="witbold">身份別: </td>
      <td>
        <select multiple name="stafftype[]" id="stafftype[]" size="9">
          <option value="1" checked>包腊</option>            
          <option value="2">包埋</option> 
          <option value="3">鑄造</option>   
          <option value="4">切牙</option>   
          <option value="5">領牙</option>   
          <option value="6">鑄造後倉管</option> 
          <option value="7">存牙倉管</option> 
          <option value="8">領牙倉管</option>   
        </select>
      </td>
    </tr>    
    <tr>
      <td bgcolor="#FF66FF" class="witbold">備註:</td>
      <td><input name="memo" type="text" id="memo" size="100"></td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF">&nbsp;</td>
      <td><input type="hidden" name="action" value="save">
          <input type="submit" name="Submit" value="送出">          
      </td>
    </tr>
  </table>
</form>
