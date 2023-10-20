<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title> <?=$_SESSION["systemname"];?> </title>
<?
  //for xajax
   if ($IsAjax){
      $xajax->printJavascript();  
   }
?>
  

<script type="text/javascript" src="scripts.js"></script>
<script type="text/javascript" src="calendarDateInput.js"></script>
<script type="text/javascript" src="My97DatePicker/WdatePicker.js"></script> 
<link href="css.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="182"><a href="doLogout.php"><h1>登出</h1></a></td>
    <td width="50" background="i/topbg.gif" class="topwhite"></td>
    <td width="560"><h1><?=$_SESSION["systemname"];?>  歡迎使用: <?=$_SESSION["name"];?></h1></td>
    <td background="i/topbg2.gif"><img src="i/topbg2.gif" width="1" height="10"></td>
  </tr>
</table>

<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="182" bgcolor="6D6E71" valign=top>
      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
       <?php
        //顯示選項標題
        $queryar = "select  a.id aid, a.name aname, a.filename afilename, a.iconname aiconname, a.isitem aisitem, a.isshow aisshow ".
                   "from aprights r , aps a " .
                   "where userguid='" . $_SESSION["userguid"] . "' and r.apguid = a.guid order by a.id";
        $resultar = mysql_query($queryar) or die ('46 ApRights error!!');
        while ($rowar= mysql_fetch_array($resultar)) {
            if ($rowar["aisshow"]=="Y") {
                if ($rowar["aisitem"]=="Y"){
                    ?>    
                    <tr>
                        <td class="navbg1"><img src="<?=$rowar["aiconname"];?>" width="16" height="16" align="absmiddle">&nbsp;<?=$rowar["aname"];?></td> 
                    </tr>
                    <?
                } else {
                    ?> 
                    <tr>
                        <td class="subnav" onmouseover="ChangeColor(this, true);" onmouseout="ChangeColor(this, false);" style="cursor:hand;" onclick=location.href="<?=$rowar["afilename"];?>">&#8226; <?=$rowar["aname"];?> </td>
                    </tr>
                    <? 
                }     
                    ?>             
          <?  
            }  
        } 
      ?>     
      <tr>
        <td>&nbsp</td>
      </tr>    
      <tr>
        <td class="topwhite"><div align="center"><?=$_SESSION["systemname"];?> </div></td>
      </tr>
    </table>      
      <p>&nbsp;</p>
    </td>
    <td class="main"><h1><?=$pagetitle;?></h1>   