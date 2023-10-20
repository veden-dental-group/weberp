<?php
  require_once('_data.php');
?>
﻿<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta NAME="robots" content="noindex,nofollow">
<title><?=$_SESSION["systemname"];?></title>
<link href="css.css" rel="stylesheet" type="text/css">
<script language="javascript" src="scripts.js"></script>
<script type="text/javascript" src="calendarDateInput.js"></script>
<style type="text/css">
<!--
body {
	background-color: #eeeeee;
}
.style1 {
	font-size: 11px;
	font-weight: bold;
}
-->
</style></head>

<body>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<form name="form1" method="post" action="doLogin.php">
  <table width="100%"  border="0" cellspacing="0" cellpadding="0" style="border: 1px solid #333333" background="i/logobg.jpg">
  <tr>
    <td background="i/topbg2.gif">&nbsp;</td>
    <td width="182"><a href="index.php"><img src="i/logo.gif" width="182" height="161" border="0"></a></td>
      <td width="300" background="i/topbg.gif" class="topwhite"><span class="style1">請輸入帳號及密碼: </span><br>
          <br>
          <table width="100%"  border="0" cellspacing="1" cellpadding="2">
            <tr>
              <td>帳號:</td>
              <td><input name="account" type="text" id="account"></td>
            </tr>
            <tr>
              <td>密碼:</td>
              <td><input name="pass" type="password" id="pass"></td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td><input name="login" type="submit" value="登入">
              <input name="action" type="hidden" id="action" value="doLogin">
              </td>
            </tr>
          </table>
      </td>
      <td background="i/topbg2.gif">&nbsp;</td>
  </tr>
  </table>
</form>
<p align="center">&nbsp;</p>   