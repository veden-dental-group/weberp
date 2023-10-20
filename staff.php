<?php
  session_start();
  $pagetitle = "系統設定 &raquo; 人員";
  include("_data.php");
  auth("staff.php");
  
  if ($_GET["action"] == "del") {        
      $query = "delete from staff where guid = '" . safetext($_GET["guid"]) . "' limit 1";
      $result = mysql_query($query) or die ('Staff deleted error!!');        
      msg('人員已刪除.');
      forward('staff.php'); 
  }
  
  $typefilter="";   
  $stafftype=$_GET["stafftype"] ;
  if ( $stafftype != "") {
      $typefilter = " where instr(stafftype, '". $stafftype ."')> 0 ";       }
  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>以下全部的人員資料!! </p>   

<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form2">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>  
        <td bgcolor="#eeeeee">
          <span align="left">身份別:
            <select name="stafftype" id="stafftype">
              <option value="" <? if  ($_GET["stafftype"]=="") echo " selected"; ?>>全部</option>
              <option value="1" <? if (strpos($_GET["stafftype"],"1")!==false) echo " selected"; ?>>包腊</option> 
              <option value="2" <? if (strpos($_GET["stafftype"],"2")!==false) echo " selected"; ?>>包埋</option> 
              <option value="3" <? if (strpos($_GET["stafftype"],"3")!==false) echo " selected"; ?>>鑄造</option> 
              <option value="4" <? if (strpos($_GET["stafftype"],"4")!==false) echo " selected"; ?>>切牙</option> 
              <option value="5" <? if (strpos($_GET["stafftype"],"5")!==false) echo " selected"; ?>>領牙</option> 
              <option value="6" <? if (strpos($_GET["stafftype"],"6")!==false) echo " selected"; ?>>鑄造後倉管</option> 
              <option value="7" <? if (strpos($_GET["stafftype"],"7")!==false) echo " selected"; ?>>存牙倉管</option> 
              <option value="8" <? if (strpos($_GET["stafftype"],"8")!==false) echo " selected"; ?>>領牙倉管</option> 
              <option value="Z" <? if (strpos($_GET["stafftype"],"Z")!==false) echo " selected"; ?>>離職</option> 
            </select>
          </span>    
          <input type="submit" name="submit" id="submit" value="查詢">            
        </td>
      </tr>
    </table>
  </div>
</form>


<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
        <th width="16">&nbsp;</th>                
        <th><a href="<?=$PHP_SELF;?>?order=name">姓名</th>   
        <th><a href="<?=$PHP_SELF;?>?order=stafftype">身份別</th>           
        <th>備註</th>
        <th>&nbsp;</th>
        <th>&nbsp;</th> 
    </tr>
    <?    
      if(empty($_GET["order"])) $_GET["order"] = "name ";                                                    
      $query = "select * from staff " . $typefilter . " order by " . $_GET['order'];
                                                                                        
	    $result = mysql_query($query) or die ('37 Staff error!!');   
      $result=mysql_query($query);   
	    while ($row= mysql_fetch_array($result)) {
	  	    $bgkleur = "ffffff";
    ?>
	        <tr bgcolor="#<?=$bgkleur;?>">
		          <td><img src="i/arrow.gif" width="16" height="16"></td>
              <td><?=$row["name"];?></td>  
              <td>
                <?
                  $stafftype='';
                  if (strpos($row["stafftype"],"1")!==false) $stafftype .= '包腊,';
                  if (strpos($row["stafftype"],"2")!==false) $stafftype .= '包埋,'; 
                  if (strpos($row["stafftype"],"3")!==false) $stafftype .= '鑄造,'; 
                  if (strpos($row["stafftype"],"4")!==false) $stafftype .= '切牙,'; 
                  if (strpos($row["stafftype"],"5")!==false) $stafftype .= '領牙,'; 
                  if (strpos($row["stafftype"],"6")!==false) $stafftype .= '鑄造後倉管,'; 
                  if (strpos($row["stafftype"],"7")!==false) $stafftype .= '存牙倉管,'; 
                  if (strpos($row["stafftype"],"8")!==false) $stafftype .= '領牙倉管,'; 
                  if (strpos($row["stafftype"],"Z")!==false) $stafftype .= '離職,'; 
                  
                  echo $stafftype; 
                ?>
              </td>  
		          <td><?=$row["memo"];?></td>      
              <td width="16"><a onclick='return confirm("確定刪除?")' href=staff.php?guid=<?=$row["guid"];?>&action=del><img border="0" src="i/delete.gif" alt="刪除"></a></td>
              <td width="16"><a href="staff_edit.php?guid=<?=$row["guid"];?>"><img src="i/edit.gif" width="16" height="16" border="0" alt="編輯"></a></td>
			    </tr>
		  <?
			}
			?>
</table>   
