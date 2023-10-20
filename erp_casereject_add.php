<?
  session_start();
  $pagetitle = "廠務 &raquo; 修改內返CASE資料";
  include("_data.php");
  auth("erp_casereject_add.php");

  if (is_null($_GET['rdate'])) {
    $rdate =  date('Y-m-d');
  } else {
    $rdate=$_GET['rdate'];
  }     
  if ($_POST["action"] == "save") {
      $queryus = "insert into casereject (rid,rdate, rqty1,rqty2) values (              
				        '" . safetext($_POST["rid"])              . "',   
                '" . safetext($_POST["rdate"])            . "', 
                '" . safetext($_POST["rqty1"])            . "',   
				        '" . safetext($_POST["rqty2"])            . "')";    
	  $resultus = mysql_query($queryus) or die ('18 Casereject added error!!.'.mysql_error());     
	  forward('erp_casereject_add.php?rdate= '.$_POST["rdate"]) ;
  }

  $showodata = "請輸入以下資料";
  include("_header.php");
?>
    
<link href="oos.css" rel="stylesheet" type="text/css">
<p><?=$showodata;?> </p>
<form name="form1" method="post" action="<?=$PHP_SELF;?>">
  <table width="100%"  border="0" cellpadding="2" cellspacing="1" class="tabel4">
    <tr>
      <td width="150" bgcolor="#FF66FF" class="witbold">日期</td>
      <td>
          <input name="rdate" type="text" id="rdate" size="12" maxlength="12" value=<?=$rdate;?> onfocus="new WdatePicker()">   
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
                  if ($_GET["mid"] == $rr2["mid"]) echo " selected";
                  echo ">" . $rr2["mid"] . ' '.$rr2['mname'] . "</option>";
                }       
              ?>
          </select>  &nbsp;&nbsp;&nbsp;&nbsp; </td>
    </tr>                                             
    <tr>
      <td bgcolor="#FF66FF" class="witbold">內返顆/床數: * </td>
      <td><input name="rqty1" type="text" id="rqty1" size="3" value="0" onmouseover="this.focus()" onfocus="this.select()" onkeypress="return numberOnly(event,'rqty1')"></td>
    </tr>    
    <tr>
      <td bgcolor="#FF66FF" class="witbold">重修返顆/床數: * </td>
      <td><input name="rqty2" type="text" id="rqty2" size="3" value="0" onmouseover="this.focus()" onfocus="this.select()" onkeypress="return numberOnly(event,'rqty2')"></td>
    </tr>     
    <tr>
      <td bgcolor="#FF66FF">&nbsp;</td>
      <td><input type="hidden" name="action" value="save">
          <input type="submit" name="Submit" value="Submit">              
      </td>
    </tr>
  </table>
</form>
