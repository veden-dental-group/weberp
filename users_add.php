<?php
  session_start();
  $pagetitle = "系統設定 &raquo; 新增帳號";
  include("_data.php");
  auth("users_add.php");

  if ($_POST["action"] == "save") {
      $queryus = "insert into users ( guid, account, password, cname, ename, phone, email, maker, memo,creator ) values (
                  '" . uuid()                           . "',   
				          '" . safetext($_POST["account"])      . "',
                  '" . md5(safetext($_POST["password1"])). "',   
                  '" . safetext($_POST["cname"])        . "',
                  '" . safetext($_POST["ename"])        . "',      
                  '" . safetext($_POST["phone"])        . "',  
                  '" . safetext($_POST["email"])        . "',    
                  '" . safetext($_POST["maker"])        . "',    
				          '" . safetext($_POST["memo"])         . "',  
                  '" . $_SESSION['account']             . "' )";     
	  $resultus = mysql_query($queryus) or die ('17 Accounts added error!!.' .mysql_error());
	  forward('users.php');
  }
  $showodata = "請輸入以下資料以新增帳號.";
  
  //for xajax
  require ('xajax/xajax_core/xajax.inc.php');
  $xajax = new xajax();
  $xajax->configure('javascript URI', 'xajax/');

  Function checkAccount($v){      
      $objResponse=new xajaxResponse();    
      if ( $v != "" ) {  // validate if input 
          $query="select account from users where account = '" .$v . "' limit 1"; 
          $result=mysql_query($query); 
          if (!$result) {
             $objResponse->alert("42 Users error!!");   
          } else {
              if ( mysql_num_rows($result)>0) {
                  $row = mysql_fetch_array($result);
                  $objResponse->assign("accountstatus","innerHTML", "警告, 帳號已存在!!");    
              }else{
                  $objResponse->assign("accountstatus","innerHTML", "");   
              }
          }  
      }          
      return $objResponse;
  } 

  $xajax->register(XAJAX_FUNCTION,'checkAccount');       
  $xajax->processRequest();

  echo '<?xml version="1.0" encoding="UTF-8"?>'; 
  $IsAjax = True;  
                       
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p><?=$showodata;?> </p>
<form name="form1" method="post" action="<?=$PHP_SELF;?>" onsubmit="return validUserAdd()">
  <table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
      <td bgcolor="#FF66FF" class="witbold">帳號: *</td>
      <td><input name="account" type="text" id="account" size="50"
          onKeyDown="xajax_checkAccount(document.getElementById('account').value)">&nbsp;&nbsp;
          <span style="color:red" id="accountstatus"></span>
      </td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF" class="witbold">密碼:</td>
      <td><input name="password1" type="password" id="password1" size="20"></td>
    </tr> 
    <tr>
      <td bgcolor="#FF66FF" class="witbold">密碼(再一次):</td>
      <td><input name="password2" type="password" id="password2" size="20"></td>
    </tr>       
    <tr>
      <td bgcolor="#FF66FF" class="witbold">中文姓名: </td>
      <td><input name="cname" type="text" id="cname" size="50"></td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF" class="witbold">英文姓名: </td>
      <td><input name="ename" type="text" id="ename" size="50"></td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF" class="witbold">電話: </td>
      <td><input name="phone" type="text" id="phone" size="50"></td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF" class="witbold">Email: </td>
      <td><input name="email" type="text" id="email" size="50"></td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF" class="witbold">製處: </td>
      <td><input name="maker" type="text" id="maker" size="50"></td>
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
