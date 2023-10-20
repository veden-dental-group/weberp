<?
  session_start();
  $pagetitle = "系統設定 &raquo; 編輯帳號";
  include("_data.php");
  auth("users_edit.php");

  if ($_POST["action"] == "save") {
	    $queryul = "update users set        
                  account      = '" . safetext($_POST["account"])     . "',   
                  cname        = '" . safetext($_POST["cname"])       . "',
                  ename        = '" . safetext($_POST["ename"])       . "',    
                  phone        = '" . safetext($_POST["phone"])       . "',  
                  email        = '" . safetext($_POST["email"])       . "',  
                  maker        = '" . safetext($_POST["maker"])       . "',  
		        	    memo         = '" . safetext($_POST["memo"])        . "'
                  where guid   = '" . safetext($_POST["guid"])         . "' limit 1";
	    $resultul = mysql_query($queryul) or die ('16 Users updated error!! ' . mysql_error()); 
      if ($_POST["password1"] != "") {
        $queryul = "update users set password='". md5(safetext($_POST["password1"])) . "'  
                   where guid   = '" . safetext($_POST["guid"])         . "' limit 1";
        $resultul = mysql_query($queryul) or die ('20 Users updated error!! ' . mysql_error());  
      } 
	    msg('更新完畢.');
	    forward('users.php');
  }

  $queryu = "select * from users where guid='" . safetext($_GET["guid"]) . "' limit 1";
  $resultu= mysql_query($queryu) or die ('28 Users Error!!');
  if (mysql_num_rows($resultu) > 0) {
	  $rowu = mysql_fetch_array($resultu);
  }else{
	  msg("帳號不存在!!");
	  forward("users.php");
  }

  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>請填入以下資料.</p>
<form name="form1" method="post" action="<?=$PHP_SELF;?>" onsubmit="return validUserAdd()">
  <table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
      <td bgcolor="#FF66FF" class="witbold">帳號: *</td>
      <td><input name="account" type="text" id="account" size="50" value="<?=$rowu["account"];?>"></td>
    </tr>  
    <tr>
      <td bgcolor="#FF66FF" class="witbold">Password: </td>
      <td><input name="password1" type="password" id="password1" size="20"> (要修改時才輸入)</td>
    </tr>    
    <tr>
      <td bgcolor="#FF66FF" class="witbold">Password again: </td>
      <td><input name="password2" type="password" id="password2" size="20"> (要修改時才輸入)</td>
    </tr>   
    <tr>
      <td bgcolor="#FF66FF" class="witbold">中文姓名: </td>
      <td><input name="cname" type="text" id="cname" size="50" value="<?=$rowu["cname"];?>"></td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF" class="witbold">英文姓名: </td>
      <td><input name="ename" type="text" id="ename" size="50" value="<?=$rowu["ename"];?>"></td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF" class="witbold">電話: </td>
      <td><input name="phone" type="text" id="phone" size="50" value="<?=$rowu["phone"];?>"></td>
    </tr>  
    <tr>
      <td bgcolor="#FF66FF" class="witbold">Email:</td>
      <td><input name="email" type="text" id="email" size="50" value="<?=$rowu["email"];?>"></td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF" class="witbold">製處:</td>
      <td><input name="maker" type="text" id="maker" size="50" value="<?=$rowu["maker"];?>"></td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF" class="witbold">備註: </td>
      <td><input name="memo" type="text" id="memo" size="100" value="<?=$rowu["memo"];?>"></td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF" class="witbold">LastLogin: </td>
      <td><input name="lastlogin" type="text" id="lastlogin" size="50" readonly="true" value="<?=$rowu["lastlogin"];?>"></td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF" class="witbold">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>   
    <tr>
      <td bgcolor="#FF66FF">&nbsp;</td>
      <td><input type="hidden" name="action" value="save">
          <input type="hidden" name="guid" value="<?=safetext($_GET["guid"]);?>">
          <input type="submit" name="Submit" value="送出"></td>
    </tr>
  </table>
</form>
