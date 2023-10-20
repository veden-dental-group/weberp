<?
  session_start();
  $pagetitle = "廠務 &raquo; 內返CASE";
  include("_data.php");
  auth("erp_casereject.php");
  if ($_GET["action"] == "del") {  
      $query = "delete from casereject where pkey = '" . safetext($_GET["pkey"]) . "' limit 1";
      $result = mysql_query($query) or die ('Casereject deleted error!!'.mysql_error());        
      msg('內返記錄已刪除.');
      forward('erp_casereject.php?rid='.$_GET['mid'].'&bdate='.$_GET['bdate'].'&edate='.$_GET['edate'].'&submit=submit');  
  }
  
  if (is_null($_GET['bdate'])) {
    $bdate =  date('Y-m-d');
  } else {
    $bdate=$_GET['bdate'];
  }                                                   
  
  if (is_null($_GET['edate'])) {
    $edate =  date('Y-m-d');
  } else {
    $edate=$_GET['edate'];
  }                                                 
                                
  include("_header.php");
?>
<link href="oos.css" rel="stylesheet" type="text/css">
<p>以下為各製處內返記錄!! </p>

<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form2">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee">製處:
            <select name="mid" id="mid">  
              <option value="">全部</option>  
              <?
                $q2 = "select mid, mname from maker order by mid";
                $r2 = mysql_query($q2) or die ('51 maker error!!');
                while ($rr2 = mysql_fetch_array($r2)) {
                   echo "<option value=" . $rr2["mid"];
                  if ($_GET["mid"] == $rr2["mid"]) echo " selected";
                  echo ">" . $rr2["mid"] . ' '.$rr2['mname'] . "</option>";
                }       
              ?>
            </select>  &nbsp;&nbsp;&nbsp;&nbsp;    
            起訖日期:   
            <input name="bdate" type="text" id="bdate" size="12" maxlength="12"  onfocus="new WdatePicker()" value=<?=$bdate;?>> ~~
            <input name="edate" type="text" id="edate" size="12" maxlength="12"  onfocus="new WdatePicker()" value=<?=$edate;?>> &nbsp;&nbsp; 
            <input type="submit" name="submit" value="送出" />      
        </td>
      </tr>
    </table>
  </div>
</form>



<form action="<?=$_SERVER['PHP_SELF'];?>" method="post" name="form1">  
  <table width="100%"  border="0" cellpadding="2" cellspacing="1" class="tabel4">
    <tr>
      <th width="16">&nbsp;</th>                        
      <th>製處</th>          
      <th>日期</th>               
      <th>內返顆/床數</th>   
      <th>重修顆/床數</th>  
      <th>&nbsp;</th> 
      <th>&nbsp;</th> 
    </tr>
    <?         
        
      if ($_GET['mid']==''){
          $midfilter=" ";
      } else {
          $midfilter="and mid='" . $_GET['mid'] . "' "; 
      } 
       
      $query = "select pkey, rid, mname, rdate, rqty1, rqty2 from casereject, maker where rid=mid  $midfilter and rdate >='$bdate' and rdate<='$edate' order by rid,rdate";
	    $result = mysql_query($query) or die ('54 Casereject error!!' . mysql_error());
      $num_max = mysql_numrows($result);       
      $limit=30;
      if(empty($_GET["start"])) {
         $start = 0;
      } else {
         $start = $_GET["start"];
      }
      $query = $query . " limit " . $start . ',' . $limit;
      $result=mysql_query($query);   
      $total1=0;
      $total2=0;
	    while ($row= mysql_fetch_array($result)) {  
        $total1 += $row["rqty1"];
        $total2 += $row["rqty2"];
	  	  $bgkleur = "ffffff";
    ?>
		  <tr bgcolor="#<?=$bgkleur;?>">
			  <td><img src="i/arrow.gif" width="16" height="16"></td>  
			  <td><?=$row["rid"];?>  <?=$row["mname"];?></td>  
			  <td><?=$row["rdate"];?></td>         
        <td style="text-align:right" ><?=number_format($row["rqty1"],1,'.',',');?></td> 
        <td style="text-align:right" ><?=number_format($row["rqty2"],1,'.',',');?></td>  
        <td width="16">
        <?
          echo '<a onclick=\'return confirm("Deleted Confirm?")\' href=erp_casereject.php?pkey='.$row["pkey"].'&mid='.$_GET['mid'].'&bdate='.$_GET['bdate']. '&edate='.$_GET['edate'].'&action=del><img border=0 src=i/delete.gif alt="Delete"></a>'
        ?>
        </td>                
			  <td width="16"><a href="erp_casereject_edit.php?pkey=<?=$row["pkey"];?>"><img src="i/edit.gif" width="16" height="16" border="0" alt="Modify"></a></td>
      </tr>    
		<?
		}
    ?> 
    <tr>
      <td colspan="3"> Total </td>
      <td style="text-align:right" ><?=number_format($total1,1,'.',',');?></td> 
      <td style="text-align:right" ><?=number_format($total2,1,'.',',');?></td>  
      <td colspan="2">&nbsp;</td>  
    </tr>    
  </table>
  <?
    $prev=$start-$limit;
    if ($prev >=0 )       { echo "<a href=erp_casereject.php?start=".$prev."&mid=".$_GET['mid'].'&bdate='.$bdate.'&edate=' . $edate. ">Prev. Page</a>" ; }   
    echo "&nbsp"; 
    $next =$start+$limit;
    if ($next < $num_max) { echo "<a href=erp_casereject.php?start=".$next."&mid=".$_GET['mid'].'&bdate='.$bdate.'&edate=' . $edate. ">Next Page</a>" ; }
  ?> 
</form>