<?
  session_start();
  $pagetitle = "主控臺 &raquo; 歡迎登入";
  include("_data.php");
  auth("main.php");
  include("_header.php");   
  $today=date('Y-m-d');
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>歡迎使用 <?=$_SESSION["systemname"];?> </p>
<table width="100%"  border="0" cellspacing="10" cellpadding="0">
  <tr>
    <td width="50%"><h2></h2></td>
    <td><h2>今日20筆賤金屬領用資料 </h2></td>
  </tr>
  <tr>
    <td valign=top><table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table3">
      <tr>
        <th width="16">&nbsp;</th>
        <th>領用日期</th>      
        <th>製處</th>
        <th>金屬</th>  
        <th align="right">腊重</th>  
        <th align="right">領取重量</th>      
      </tr>
      <?
        $query="select v.laweight vlaweight, v.takeweight vtakeweight, p.name pname, m.name mname " .  
               "from valued v, processes p, metals m " .
               "where v.processguid=p.guid and v.metalguid=m.guid and v.inputdate=' " . $today . "' " . 
               "order by p.name limit 20";     
	      $result = mysql_query($query) or die ('31 Valued error!!' .mysql_error());
	      while ($row = mysql_fetch_array($result)) {
	  	?>
      <tr>        
        <td><img src="i/arrow.gif" width="16" height="16"></td>
        <td><?=$today;?></td>
        <td><?=$row["pname"];?></td> 
        <td><?=$row["mname"];?></td> 
        <td align="right"><?=$row["vlaweight"];?></td>
        <td align="right" ><?=$row["vtakeweight"];?></td>  
      </tr>
      <?
	    }
	    ?>
      
    </table></td>
    <td valign=top><table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table2">
      <tr>
        <th width="16">&nbsp;</th>
        <th>領用日期</th>      
        <th>製處</th>
        <th>金屬</th>  
        <th align="right">腊重</th>  
        <th align="right">領取重量</th>      
      </tr>
      <?
        $query="select c.laweight claweight, c.takeweight ctakeweight, p.name pname, m.name mname " .  
               "from cheap c, processes p, metals m " .
               "where c.processguid=p.guid and c.metalguid=m.guid and c.inputdate=' " . $today . "' " . 
               "order by p.name limit 20";     
          $result = mysql_query($query) or die ('61 Cheap error!!' .mysql_error());
          while ($row = mysql_fetch_array($result)) {
	    ?>
      <tr>        
        <td><img src="i/arrow.gif" width="16" height="16"></td>
        <td><?=$today;?></td>
        <td><?=$row["pname"];?></td> 
        <td><?=$row["mname"];?></td> 
        <td align="right"><?=$row["claweight"];?></td>
        <td align="right"><?=$row["ctakeweight"];?></td>   
      </tr>
      <?
	     }
	    ?>  
    </table>
    <p>&nbsp;</p>
    </td>
  </tr>
</table>
