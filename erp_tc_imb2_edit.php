<?
  session_start();
  $pagetitle = "資材部 &raquo; 編輯料號";
  include("_data.php");
  //auth("erp_tc_imb2_edit.php");

  if ($_POST["action"] == "save") {
	    $queryul = "update erp_tc_imb2 set        
                  code        = '" . safetext($_POST["code"])     . "',   
                  name        = '" . safetext($_POST["name"])     . "',
                  lotno       = '" . safetext($_POST["lotno"])    . "'    
                  where pkey  = '" . safetext($_POST["pkey"])     . "' limit 1";
	    $resultul = mysql_query($queryul) or die ('16 erp_tc_imb2 updated error!! ' . mysql_error());  
	    msg('更新完畢.');
	    forward('erp_tc_imb2.php');
  }

  $queryu = "select * from erp_tc_imb2 where pkey='" . safetext($_GET["pkey"]) . "' limit 1";
  $resultu= mysql_query($queryu) or die ('28 erp_tc_imb2 Error!!');
  if (mysql_num_rows($resultu) > 0) {
	  $rowu = mysql_fetch_array($resultu);
  }else{
	  msg("料號不存在!!");
	  forward("erp_tc_imb2.php");
  }

  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>請填入以下資料.</p>
<form name="form1" method="post" action="<?=$PHP_SELF;?>">
  <table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
      <td bgcolor="#FF66FF" class="witbold">料號: *</td>
      <td><input name="code" type="text" id="code" size="15" maxlength="15" value="<?=$rowu["code"];?>"></td>
    </tr>   
    <tr>
      <td bgcolor="#FF66FF" class="witbold">名稱: </td>
      <td><input name="name" type="text" id="name" size="50" maxlength="50" value="<?=$rowu["name"];?>"></td>
    </tr>
    <tr>
      <td bgcolor="#FF66FF" class="witbold">Lot No.: </td>
      <td><input name="lotno" type="text" id="lotno" size="20" maxlength="20" value="<?=$rowu["lotno"];?>"></td>
    </tr>        
    <tr>
      <td bgcolor="#FF66FF" class="witbold">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>   
    <tr>
      <td bgcolor="#FF66FF">&nbsp;</td>
      <td><input type="hidden" name="action" value="save">
          <input type="hidden" name="pkey"   value="<?=safetext($_GET["pkey"]);?>">
          <input type="submit" name="Submit" value="送出"></td>
    </tr>
  </table>
</form>
