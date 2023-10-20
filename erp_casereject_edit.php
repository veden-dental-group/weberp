<?
  session_start();
  $pagetitle = "廠務 &raquo; 編輯內返CASE";
  include("_data.php");
  auth("erp_casereject_edit.php");

  if ($_POST["action"] == "save") {
	    $queryco            = "update casereject set
            rid           = '" . safetext($_POST["rid"])            . "',
            rdate         = '" . safetext($_POST["rdate"])          . "', 
            rqty1         = '" . safetext($_POST["rqty1"])          . "',
            rqty2         = '" . safetext($_POST["rqty2"])          . "'  
            where pkey    = '" . safetext($_POST["pkey"])           . "' limit 1";
	  $resultco = mysql_query($queryco) or die ('32 Casereject update error! ' . mysql_error());  
	  msg('資料已更新!!');
	  forward('erp_casereject.php');
  }

  $query = "select * from casereject where pkey = '" . safetext($_GET["pkey"]) . "' limit 1";
  $result = mysql_query($query) or die ('23 Casereject error!!');
  if (mysql_num_rows($result) > 0) {
	  $row = mysql_fetch_array($result); 
  }else{
	  msg("資料不存在.");
	  forward("erp_casereject.php");
  }
 
  include("_header.php");
?>
<link href="oos.css" rel="stylesheet" type="text/css">

<p>請填入以下資料<br></p>
<form name="form1" method="post" action="<?=$PHP_SELF;?>" >
  <table width="100%"  border="0" cellpadding="2" cellspacing="1" class="tabel4">
    <tr>
      <td width="150" bgcolor="#FF66FF" class="witbold">日期: *</td>
      <td>
        <input name="rdate" type="text" id="rdate" size="12" maxlength="12"  onfocus="new WdatePicker()" value=<?=$row['rdate'];?>>   
      </td>
    </tr>
    <tr>
      <td width="150" bgcolor="#FF66FF" class="witbold">製處: *</td>
      <td>
          <select name="rid" id="rid">      
              <?
                $q2 = "select mid, mname from maker order by mid";
                $r2 = mysql_query($q2) or die ('51 maker error!!');
                while ($rr2 = mysql_fetch_array($r2)) {
                   echo "<option value=" . $rr2["mid"];
                  if ($_row["rid"] == $rr2["mid"]) echo " selected";
                  echo ">" . $rr2["mid"] . ' '.$rr2['mname'] . "</option>";
                }       
              ?>
          </select>  
      </td>
    </tr>    
    <tr>
      <td bgcolor="#FF66FF" class="witbold">內返顆/床數: * </td>
      <td><input name="rqty1" type="text" id="rqty1" size="3" value="<?=$row['rqty1'];?>" onmouseover="this.focus()" onfocus="this.select()" onkeypress="return numberOnly(event,'rqty1')"></td>
    </tr>    
    <tr>
      <td bgcolor="#FF66FF" class="witbold">重修顆/床數: * </td>
      <td><input name="rqty2" type="text" id="rqty2" size="3" value="<?=$row['rqty2'];?>" onmouseover="this.focus()" onfocus="this.select()" onkeypress="return numberOnly(event,'rqty2')"></td>
    </tr>           
    <tr>
      <td bgcolor="#FF66FF" class="witbold">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>   
    <tr>
      <td bgcolor="#FF66FF">&nbsp;</td>
      <td><input type="hidden" name="action" value="save">
          <input type="hidden" name="pkey" value="<?=safetext($_GET["pkey"]);?>">
          <input type="submit" name="Submit" value="Submit"></td>
    </tr>
  </table>
</form>
