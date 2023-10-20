<?
  session_start();
  $pagetitle = "系統設定 &raquo; 編輯程式";
  include("_data.php");
  auth("aps_edit.php");

  if ($_POST["action"] == "save") {
	    $queryap      = "update aps set
				  id        = '" . safetext($_POST["id"])         . "',
				  name      = '" . safetext($_POST["name"])       . "',  
          filename  = '" . safetext($_POST["filename"])   . "',
          iconname  = '" . safetext($_POST["iconname"])   . "',
          isitem    = '" . safetext($_POST["isitem"])     . "',
          isshow    = '" . safetext($_POST["isshow"])     . "',  
				  memo      = '" . safetext($_POST["memo"])       . "' 
          where guid = '" . safetext($_GET["guid"]) . "' limit 1";
	  $resultap = mysql_query($queryap) or die ('APs update error. ' . mysql_error());  
	  msg('更新完畢.');
	  forward('aps.php');
  }

  $queryap = "select * from aps where guid = '" . safetext($_GET["guid"]) . "' limit 1";
  $resultap = mysql_query($queryap) or die ('23 Aps error!!');
  if (mysql_num_rows($resultap) > 0) {
	  $rowap = mysql_fetch_array($resultap);
  }else{
	  msg("Applicaton not found.");
	  forward("aps.php");
  }

  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>請填入以下資料.<br></p>
<form name="form1" method="post" action="<?=$PHP_SELF;?>">
  <table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
      <td width="150" bgcolor="#FF66FF" class="witbold">序號:</td>
      <td><input name="id" type="text" id="id" size="10" value="<?=$rowap["id"];?>"></td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF" class="witbold">程式用途:</td>
      <td><input name="name" type="text" id="name" size="20" value="<?=$rowap["name"];?>"></td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF" class="witbold">檔名:</td>
      <td><input name="filename" type="text" id="filename" size="20" value="<?=$rowap["filename"];?>"></td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF" class="witbold">圖示:</td>
      <td><input name="iconname" type="text" id="iconname" size="20" value="<?=$rowap["iconname"];?>"></td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF"><span class="witbold">是否標題: </span></td>
      <td><input name="isitem" type="radio" id="isitem" value="Y" 
                <?php
                    if ($rowap["isitem"]=='Y') echo " checked";
                ?> >Yes
              <input name="isitem" type="radio" id="isitem" value="N" 
               <?php
                     if ($rowap["isitem"]=='N' ) echo " checked";
               ?> >No</td>    
    </tr> 
    <tr>
      <td bgcolor="#FF66FF"><span class="witbold">是否顯示: </span></td>
      <td><input name="isshow" type="radio" id="isshow" value="Y" 
                <?php
                    if ($rowap["isshow"]=='Y') echo " checked";
                ?> >Yes
              <input name="isshow" type="radio" id="isshow" value="N" 
               <?php
                     if ($rowap["isshow"]=='N' ) echo " checked";
               ?> >No</td>    
    </tr>  
    <tr>
      <td bgcolor="#FF66FF" class="witbold">備註: </td>
      <td><input name="memo" type="text" id="memo" size="100" value="<?=$rowap["memo"];?>"></td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF" class="witbold">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>   
    <tr>
      <td bgcolor="#FF66FF">&nbsp;</td>
      <td><input type="hidden" name="action" value="save">
          <input type="submit" name="Submit" value="送出"></td>
    </tr>
  </table>
</form>
