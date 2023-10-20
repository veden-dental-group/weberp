<?
  session_start();
  $pagtitle = "系統設定 &raquo; 權限"; 
  include("_data.php");
  auth("aprights.php");
  
  $idfilter = " order by aid";
  // select an user
  if ($_GET["Submit2"] != "") {  
	  $querys  = "select guid from aprights where userguid ='" . $_GET["userguid"] . "' limit 1";
	  $results = mysql_query($querys) or die ('14 ApRights error!!');
	  if (mysql_num_rows($results) == 0) {
		  msg('查無資料!!');        
	  } else {
          $idfilter = "where r.userguid ='" . $_GET["userguid"] . "' order by aid ";
    }
  //application rights    
  } else if ($_GET["Submit3"] != "") {
      $queryr  = "select guid from aprights where apguid ='" . $_GET["apguid"] . "' limit 1";  
      $resultr = mysql_query($queryr) or die ('APRights error!!');
      if (mysql_num_rows($resultr) == 0) {
          msg('查無資料!!');
      } else {        
          $idfilter = "where r.apguid ='" . $_GET["apguid"] . "' order by uaccount ";   
      }  
  }

  if ($_GET["action"] == "del") {
      $query = "delete from aprights where guid = '" . safetext($_GET["rguid"]) . "' limit 1";
      $result = mysql_query($query) or die ('38 ApRights delete error!!');        
      msg('刪除成功.');
      forward('aprights.php');    
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
	for (var i = 0; i < document.getElementById('form1').elements.length; i++) {
		document.getElementById('form1').elements[i].checked = checked;
	}
}
</script>
<link href="css.css" rel="stylesheet" type="text/css">
<p>以下全部的程式權限資料!! </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
      
        <td bgcolor="#eeeeee"><div align="left">帳號:
            <select name="userguid" id="userguid">  
              <option value="">請選擇一個帳號</option>  
              <?
                $q2 = "select guid, account from users order by account";
                $r2 = mysql_query($q2) or die ('71 Users error!!');
                while ($rr2 = mysql_fetch_array($r2)) {
   	              echo "<option value=" . $rr2["guid"];
	                if ($_GET["userguid"] == $rr2["guid"]) echo " selected";
	                echo ">" . $rr2["account"] . "</option>";
                }    
              ?>
            </select>
            &nbsp;
            <input type="submit" name="Submit2" value="查詢" />
       </div></td>
        
        <td bgcolor="#eeeeee"><div align="left">程式:
            <select name="apguid" id="apguid">  
              <option value="">請選擇一個程式</option>  
              <?
                $q3 = "select guid, id, name  from aps order by id";  
                $r3 = mysql_query($q3) or die ('88 Aps error!!');
                while ($rr3 = mysql_fetch_array($r3)) {
                   echo "<option value=" . $rr3["guid"];
                   if ($_GET["apguid"] == $rr3["guid"]) echo " selected";
                   echo ">" . $rr3["id"]. ". ". $rr3["name"] . "</option>";
                }    
              ?>
            </select>
            &nbsp;
            <input type="submit" name="Submit3" value="查詢" />
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
    <th>帳號</th>
    <th>中文名</th>       
    <th>權限</th>
    <th>&nbsp;</th>     
  </tr>
  <?    
  $query = "select a.id aid, a.name aname, a.filename afilename, a.isitem aisitem, r.guid rguid, r.apguid rapguid, " . 
           "r.isok risok, u.guid uguid, u.account uaccount, u.cname ucname " .
           "from aps a left join aprights r on a.guid=r.apguid inner join users u on r.userguid=u.guid " . $idfilter  ;                 
           
  $result = mysql_query($query) or die ('125 ApRights error!!');
  $num_max = mysql_numrows($result);       
  $limit=200;
  if(empty($_GET["start"])) {$start = 0;}
  else {$start = $_GET["start"];}
  $query = $query . " limit " . $start . ',' . $limit;
  $result=mysql_query($query);
  $sqtytotal=0;
  while ($row = mysql_fetch_array($result)) {
    $bgcolor = "ffffff"; 
    if ($row["aisitem"] == 'Y') { $bgcolor = "FB6A3E"; }   
  ?>     
	  <tr bgcolor="#<?=$bgcolor;?>">  
		  <td><img src="i/arrow.gif" width="16" height="16"></td>
			<td><?=$row["aid"];?></td>
			<td><?=$row["aname"];?></td>
			<td><?=$row["afilename"];?></td>
			<td><?=$row["aisitem"];?></td>
			<td><?=$row["uaccount"];?></td>
      <td><?=$row["ucname"];?></td>    
      <td><?=$row["risok"];?></td>
      <td width=16>
      <?
        echo '<a onclick=\'return confirm("確認刪除?")\' href=aprights.php?rguid='.$row["rguid"].'&action=del' . '><img border=0 src=i/delete.gif alt="刪除"></a>'
			?>
      </td>
   </tr> 
  <?  
	 }
	?>  
</table>
<?
  $prve=$start-$limit;
  if ($prve >=0 )       { echo "<a href=aprights.php?start=$prve>Prev. Page</a>" ; }   
  echo ("&nbsp;");
  $next =$start+$limit;
  if ($next < $num_max) { echo "<a href=aprights.php?start=$next>Next Page</a>" ; } 
?>    
</form>
