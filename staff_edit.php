<?
  session_start();
  $pagetitle = "系統設定 &raquo; 編輯人員";
  include("_data.php");
  auth("staff_edit.php");

  $stafftypes = $_POST['stafftype'];
  $stafftype='';
  foreach ($stafftypes as $v) {
    $stafftype .= $v; 
  }
  
  if ($_POST["action"] == "save") {
	    $queryul = "update staff set                                           
                  name         = '" . safetext($_POST["name"])       . "',    
                  stafftype    = '" . $stafftype . "',      
                  memo         = '" . safetext($_POST["memo"])       . "'  
                  where guid   = '" . safetext($_GET["guid"])         . "' limit 1";
	    $resultul = mysql_query($queryul) or die ('16 Staff updated error!! ' . mysql_error()); 
	    msg('更新完畢.');
	    forward('staff.php');
  }

  $queryu = "select * from staff where guid='" . safetext($_GET["guid"]) . "' limit 1";
  $resultu= mysql_query($queryu) or die ('28 Staff Error!!');
  if (mysql_num_rows($resultu) > 0) {
	  $rowu = mysql_fetch_array($resultu);
  }else{
	  msg("資料不存在!!");
	  forward("staff.php");
  }

  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>請填入以下資料.</p>
<form name="form1" method="post" action="<?=$PHP_SELF;?>" onsubmit="return validStaffAdd()">
  <table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
      <td bgcolor="#FF66FF" class="witbold">名稱: * </td>
      <td><input name="name" type="text" id="name" size="50" value="<?=$rowu["name"];?>"></td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF" class="witbold">身份: </td>
      <td>
        <select multiple name="stafftype[]" id="stafftype[]" size='10'>
          <option value="1" <? if (strpos($rowu["stafftype"],"1")!==false) echo " selected"; ?>>包腊 </option>            
          <option value="2" <? if (strpos($rowu["stafftype"],"2")!==false) echo " selected"; ?>>包埋</option>
          <option value="3" <? if (strpos($rowu["stafftype"],"3")!==false) echo " selected"; ?>>鑄造</option>  
          <option value="4" <? if (strpos($rowu["stafftype"],"4")!==false) echo " selected"; ?>>切牙</option>  
          <option value="5" <? if (strpos($rowu["stafftype"],"5")!==false) echo " selected"; ?>>領牙</option>  
          <option value="6" <? if (strpos($rowu["stafftype"],"6")!==false) echo " selected"; ?>>鑄造後倉管</option>  
          <option value="7" <? if (strpos($rowu["stafftype"],"7")!==false) echo " selected"; ?>>存牙倉管</option>      
          <option value="8" <? if (strpos($rowu["stafftype"],"8")!==false) echo " selected"; ?>>領倉管</option>      
          <option value="Z" <? if (strpos($rowu["stafftype"],"Z")!==false) echo " selected"; ?>>離職</option>     
        </select>
      </td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF" class="witbold">備註: </td>
      <td><input name="memo" type="text" id="memo" size="100" value="<?=$rowu["memo"];?>"></td>    
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
