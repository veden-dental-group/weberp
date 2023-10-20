<?php
  session_start();
  $pagetitle = "系統設定 &raquo; 帳號";
  include("_data.php");
  auth("users.php");
  
  if ($_GET["action"] == "del") {
      //delete the user's app. rights
      $query = "delete from aprights where userguid='" . safetext($_GET["guid"]) . "'";
      $result=mysql_query($query) or die ("Aprights deleted error!!");
      //delet the user' data
      $query = "delete from users where guid = '" . safetext($_GET["guid"]) . "' limit 1";
      $result = mysql_query($query) or die ('Users deleted error!!');        
      msg('帳號已刪除.');
      forward('users.php'); 
  }
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>以下全部的帳號資料!! </p>   
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
        <th width="16">&nbsp;</th>                
        <th><a href="<?=$PHP_SELF;?>?order=account">帳號</th>   
        <th><a href="<?=$PHP_SELF;?>?order=cname">中文姓名</th>          
        <th><a href="<?=$PHP_SELF;?>?order=ename">英文姓名</th>
        <th><a href="<?=$PHP_SELF;?>?order=phone">電話</th>  
        <th><a href="<?=$PHP_SELF;?>?order=email">EMAIL</th>  
        <th>製處</th> 
        <th><a href="<?=$PHP_SELF;?>?order=lastlogin">最後登入</th>
        <th>&nbsp;</th>
        <th>&nbsp;</th> 
    </tr>
    <?
      if(empty($_GET["order"])) $_GET["order"] = "account";  
      $query = "select * from users order by " . $_GET["order"];
                                                                                                  
	    $result = mysql_query($query) or die ('37 Users error!!');   
      $result=mysql_query($query);   
	    while ($row= mysql_fetch_array($result)) {
	  	    $bgkleur = "ffffff";
    ?>
	        <tr bgcolor="#<?=$bgkleur;?>">
		          <td><img src="i/arrow.gif" width="16" height="16"></td>
              <td><?=$row["account"];?></td>  
		          <td><?=$row["cname"];?></td>
		          <td><?=$row["ename"];?></td>
		          <td><?=$row["phone"];?></td>
              <td><?=$row["email"];?></td>    
              <td><?=$row["maker"];?></td>      
				      <td><?=$row["lastlogin"];?></td>
              <td width="16"><a onclick='return confirm("確定刪除?")' href=users.php?guid=<?=$row["guid"];?>&action=del><img border="0" src="i/delete.gif" alt="刪除"></a></td>
              <td width="16"><a href="users_edit.php?guid=<?=$row["guid"];?>"><img src="i/edit.gif" width="16" height="16" border="0" alt="編輯"></a></td>
			    </tr>
		  <?
			}
			?>
</table>   
