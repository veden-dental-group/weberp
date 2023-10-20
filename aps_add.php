<?
  session_start();
  $pagtitel = "系統設定 &raquo; 新增程式";
  include("_data.php");
  //auth("aps_add.php");

  if ($_POST["action"] == "save") {
	    $queryap   = "insert into aps ( guid, id, name, filename, iconname, isitem, isshow, memo ) values (
                  '" . uuid()                         . "', 
                  '" . safetext($_POST["id"])         . "',  
				          '" . safetext($_POST["name"])       . "',      
                  '" . safetext($_POST["filename"])   . "', 
                  '" . safetext($_POST["iconname"])   . "',       
                  '" . safetext($_POST["isitem"])     . "',  
                  '" . safetext($_POST["isshow"])     . "',     
				          '" . safetext($_POST["memo"])       .  "')";
	    $resultap= mysql_query($queryap) or die ('17 Aps added error. ' .mysql_error());
      forward('aps_add.php?oname= ( ' . safetext($_POST["id"]). ". ". safetext($_POST["name"]).'--> Added)');     
  }
  $showoname = "請輸入以下資料以新增程式.  ".$_GET["oname"]; 
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p><?=$showoname;?></p>
<form name="form1" method="post" action="<?=$PHP_SELF;?>">
  <table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
      <td width="150" bgcolor="#FF66FF" class="witbold">序號:</td>
      <td><input name="id" type="text" id="id" size="10"></td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF" class="witbold">程式用途: </td>
      <td><input name="name" type="text" id="name" size="20"></td>
    </tr>  
    <tr>
      <td bgcolor="#FF66FF" class="witbold">檔名: </td>
      <td><input name="filename" type="text" id="filename" size="40"></td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF" class="witbold">圖示: </td>
      <td><input name="iconname" type="text" id="iconname" size="40"></td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF"><span class="witbold">是否標題: </span></td>
      <td><input name="isitem" type="radio" id="isitem" value="Y">Yes
          <input name="isitem" type="radio" id="isitem" value="N" checked>No</td>    
    </tr>
    <tr>
      <td bgcolor="#FF66FF"><span class="witbold">是否顯示: </span></td>
      <td><input name="isshow" type="radio" id="isshow" value="Y" checked>Yes
          <input name="isshow" type="radio" id="isshow" value="N" >No</td>    
    </tr> 
    <tr>
      <td bgcolor="#FF66FF" class="witbold">備註:</td>
      <td><input name="memo" type="text" id="memo" size="100"></td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF" class="witbold">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
     <tr>
      <td bgcolor="#FF66FF">&nbsp;</td>
      <td><input type="hidden" name="action" value="save">
          <input type="submit" name="Submit" value="送出">          
      </td>
    </tr>
  </table>
</form>
