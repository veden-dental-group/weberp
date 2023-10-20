<?php
    session_start();                    
    $pagtitle = "系統設定 &raquo; 程式";
    include("_data.php");
    //auth("aps.php");
    if ($_GET["action"] == "del") {
        $query = "delete from aps where guid = '" . safetext($_GET["guid"]) . "' limit 1";
        $result = mysql_query($query) or die ('Aps deleted error!!');        
        msg('Application deleted.');
        forward('aps.php');    
    }

    include("_header.php");
?>

<link href="css.css" rel="stylesheet" type="text/css">
<p>以下全部的程式資料!! </p>
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
        <th width="16">&nbsp;</th>
        <th>序號</th>
        <th>程式用途</th> 
        <th>檔名</th>        
        <th>圖示</th>   
        <th>是否標題</th>   
        <th>是否顯示</th>  
        <th>備註</th>  
        <th>&nbsp;</th>
        <th>&nbsp;</th> 
    </tr>
    <?
   	$queryap = "select * from aps  order by id";
	  $resultap = mysql_query($queryap) or die ('32 APs load error');
	  while ($rowap= mysql_fetch_array($resultap)) {
        $bgcolor = "ffffff"; 
        if ($rowap["isitem"] == 'Y') { $bgcolor = "FB6A3E"; }
	  ?>
			      <tr bgcolor="#<?=$bgcolor;?>">
				        <td><img src="i/arrow.gif" width="16" height="16"></td>
				        <td><?=$rowap["id"];?></td>
                <td><?=$rowap["name"];?></td>  
                <td><?=$rowap["filename"];?></td>  
                <td><?=$rowap["iconname"];?></td>                  
                <td><?=$rowap["isitem"];?></td> 
                <td><?=$rowap["isshow"];?></td>  
			          <td><?=$rowap["memo"];?></td>
                <td width="16">
                <?
                  echo '<a onclick=\'return confirm("確定刪除?")\' href=aps.php?guid='.$rowap["guid"].'&action=del' . '><img border=0 src=i/delete.gif alt="刪除"></a>'
                ?>
                </td>
                <td width="16"><a href="aps_edit.php?guid=<?=$rowap["guid"];?>"><img src="i/edit.gif" width="16" height="16" border="0" alt="編輯"></a></td>
			    </tr>
	  <?
	  }
   	?>
</table>
