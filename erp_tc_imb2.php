<?php
  session_start();
  $pagetitle = "資材部 &raquo; 磁粉Lot No.";
  include("_data.php");
  //auth("erp_tc_imb2.php");
  
  if ($_GET["action"] == "del") {     
      $query = "delete from erp_tc_imb2 where pkey = '" . safetext($_GET["pkey"]) . "' limit 1";
      $result = mysql_query($query) or die ('erp_tc_imb2 deleted error!!');        
      msg('資料已刪除.');
      forward('erp_tc_imb2.php'); 
  }
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>以下全部的磁粉Lot No.資料!! </p>   
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
        <th width="16">&nbsp;</th>                
        <th><a href="<?=$PHP_SELF;?>?order=code,ldate">料號</th>   
        <th><a href="<?=$PHP_SELF;?>?order=name,ldate">名稱</th>
        <th><a href="<?=$PHP_SELF;?>?order=ldate,code">日期</th> 
        <th><a href="<?=$PHP_SELF;?>?order=lotno,code">Lot No.</th>   
        <th>&nbsp;</th>
        <th>&nbsp;</th> 
        <th>&nbsp;</th>  
    </tr>
    <? 
      if(empty($_GET["order"])) $_GET["order"] = "code";                                         
      $query = "select * from erp_tc_imb2 order by " . $_GET["order"];
                                                                                                  
	    $result = mysql_query($query) or die ('37 erp_tc_imb2 error!!');   
      $result=mysql_query($query);   
	    while ($row= mysql_fetch_array($result)) {
	  	    $bgkleur = "ffffff";
    ?>
	        <tr bgcolor="#<?=$bgkleur;?>">
		          <td><img src="i/arrow.gif" width="16" height="16"></td>
              <td><?=$row["code"];?></td>  
		          <td><?=$row["name"];?></td>
              <td><?=$row["ldate"];?></td>   
		          <td><?=$row["lotno"];?></td> 
              <td width="16"><a href="erp_tc_imb2_add.php?code=<?=$row["code"];?>&name=<?=$row['name'];?>"><img src="i/add.png" width="16" height="16" border="0" alt="新增"></a></td>   
              <td width="16"><a onclick='return confirm("確定刪除?")' href=erp_tc_imb2.php?pkey=<?=$row["pkey"];?>&action=del><img border="0" src="i/delete.gif" alt="刪除"></a></td>
              <td width="16"><a href="erp_tc_imb2_edit.php?pkey=<?=$row["pkey"];?>"><img src="i/edit.gif" width="16" height="16" border="0" alt="編輯"></a></td>
			    </tr>
		  <?
			}
			?>
</table>   
