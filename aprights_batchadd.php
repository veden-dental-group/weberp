<?
  session_start();
  $pagtitle = "系統設定 &raquo; 設定權限"; 
  include("_data.php");
  //auth("aprights_batchadd.php");

  if ($_POST["action"] == "save") {
      //先刪除所有的資料 再新增
      $userguid=$_POST['userguid'];
      $queryr = "delete from aprights where userguid = '$userguid'";
      $resultr = mysql_query($queryr) or die ('10 ApRights deleted error!!');
      
      foreach ($_POST["aguidarray"] as $apguid){
          if ($_POST["risok" . $apguid] == "Y") {
                  $queryr = "insert into aprights (guid, apguid, userguid, isok)values ( 
                            '" . uuid()                   . "',
                            '" . safetext($apguid)        . "',
                            '" . $userguid                . "','Y')";     
                  $resultr = mysql_query($queryr) or die ('19 ApRights added error!!'); 
          }
      }    
      
      forward("aprights_batchadd.php?userguid=$userguid");                                                            
  }

  include("_header.php");
?>
<script language='JavaScript'>
checked = false;
function checkedAll () {
  if (checked == false) {
    checked = true
  }else{
    checked = false
  }
  
  for (var i = 0; i < document.form1.elements.length; i++) {
    var e = document.form1.elements[i];
        if (e.type == 'checkbox' && e.disabled==false) {
            e.checked = checked
        }                                                          
  }  
}
</script>
<link href="css.css" rel="stylesheet" type="text/css">
<p>以下全部的程式權限資料. </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">帳號:
            <select name="userguid" id="userguid">  
              <option value="">請選擇一個帳號</option>  
              <?
                $q2 = "select guid, account, cname from users order by account";
                $r2 = mysql_query($q2) or die ('51 Users error!!');
                while ($rr2 = mysql_fetch_array($r2)) {
                   echo "<option value=" . $rr2["guid"];
                  if ($_GET["userguid"] == $rr2["guid"]) echo " selected";
                  echo ">" . $rr2["account"] . '--'.$rr2['cname'] . "</option>";
                }       
              ?>
            </select>
            &nbsp;
            <input type="submit" name="Submit2" value="送出" />
            </div></td>        
        </tr>
    </table>
  </div>
</form>

<form action="<?=$PHP_SELF;?>" method="post" name="form1">
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
  <tr>
    <th width="16">&nbsp;</th>
    <th>序號</th>
    <th>程式用途</th>
    <th>檔名</th>
    <th>是否標題</th>  
    <th width="16"><input type="checkbox" value="checkbox" name='checkall' onclick='checkedAll();'></th>
  </tr>
  <?
  if ($_GET["userguid"]=="") {
      $query = "select a.guid aguid, a.id aid, a.name aname, a.filename afilename, a.isitem aisitem " .
               "from aps a order by aid";
  } else {
      $query = "select a.guid aguid, a.id aid, a.name aname, a.filename afilename, a.isitem aisitem, " . 
               "ar.risok arisok from aps a left join ".
               "(select r.apguid rapguid, r.isok risok, r.userguid  from aprights r where r.userguid='" . 
               $_GET["userguid"] . "') as ar on a.guid = ar.rapguid order by aid asc ";
  }
  $result = mysql_query($query) or die ('87 ApRights error!!');
  $result=mysql_query($query);
  while ($row = mysql_fetch_array($result)) {  
    $bgcolor = "ffffff"; 
    if ($row["aisitem"] == 'Y') { $bgcolor = "FB6A3E"; }    
    ?>    
	    <tr bgcolor="#<?=$bgcolor;?>"> 
		  <td><img src="i/arrow.gif" width="16" height="16">  
          <input type="hidden" name="aguidarray[]" value="<?=$row["aguid"];?>">  </td>
			<td><?=$row["aid"];?></td>
			<td><?=$row["aname"];?></td>
			<td><?=$row["afilename"];?></td>
			<td><?=$row["aisitem"];?></td>
      <td width=16><input name="risok<?=$row["aguid"];?>" type="checkbox" id="risok" value="Y" 
      <?
        if ($row["arisok"]=="Y") {echo "checked"; }
       ?>>
      </td>  
     </tr> 
  <?  
	}
	?>    
  <tr>
    <td>&nbsp;</td>
    <td><input type="hidden" name="userguid" value=<?=$_GET["userguid"];?>>
        <input type="hidden" name="action" value="save">
        <input type="submit" name="Submit" value="送出">        
    </td>
  </tr>
</table>  
</form>
