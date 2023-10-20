<?php
  session_start();
  $pagetitle = "資材部 &raquo; 查詢鑄造後重量";    
  include("_data.php");
  //auth("erp_metal.php");

  
  $bdatefilter="";   
  $ticketnofilter="";   
  $limit = "";
  if ($_GET["bdate"] != "") {
      $bdatefilter = " and bdate='" . $_GET["bdate"] . "' ";        
  } 
  if ($_GET["ticketno"] != "") {
      $ticketnofilter = " and ticketno like '%" . $_GET["ticketno"] . "%' ";        
  }
 //如果沒有給任何條件 不秀 否則資料量太大
  if ($_GET["bdate"]. $_GET["ticketno"] == ""){
      $limit = " limit 0 ";
  }   

  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>以下為金屬鑄造後重量!! </p>   
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr> 
        <td bgcolor="#eeeeee">
          <span align="left">日期:  
            <input name="bdate"   type="text" id="bdate" size="12" maxlength="12" onfocus="new WdatePicker()"  value=<?=$_GET["bdate"];?> >           
          </span>       
        </td>   
        <td bgcolor="#eeeeee">工單:<input type="text" name="ticketno" id="ticketno" size="15" ></td>  
        <td bgcolor="#eeeeee">   
          <input type="submit" name="submit" id="submit" value="查詢">  
        </td>
      </tr>
    </table>
  </div>
</form>

<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table3">
    <tr>
        <th width="16">&nbsp;</th>
        <th>日期</th>        
        <th>金屬</th>  
        <th>工單號</th>    
        <th>鑄造後重量</th> 
        <th>研磨後重量</th>   
        <th>備註</th>   
        <th>&nbsp;</th> 
    </tr>
    <?
       
      $query = "select pkey, bdate, metal, ticketno, weight1, weight2, memo from  erp_metal_add where 1=1 " . $bdatefilter . $ticketnofilter . " order by bdate,metal,ticketno "  . $limit;                                                                                                   
	    $result = mysql_query($query) or die ('37 erp_metal_add error!!'.mysql_error());   
        $result=mysql_query($query);   
        $i=0;
	    while ($row= mysql_fetch_array($result)) {       
            if ($i % 2 == 0) $bgcolor = "ffffff";
            if ($i % 2 == 1) $bgcolor = "DDDDDD";    
    ?>
	        <tr bgcolor="#<?=$bgcolor;?>">
		      <td><img src="i/arrow.gif" width="16" height="16"></td>  
              <td><?=$row["bdate"];?></td>   
              <td><?=$row["metal"];?></td>    
              <td><?=$row["ticketno"];?></td>   
              <td><?=$row["weight1"];?></td>  
              <td><?=$row["weight2"];?></td> 
              <td><?=$row["memo"];?></td>                                                                                                                                                                          
              <td width="16"><a href="erp_metal_edit.php?pkey=<?=$row["pkey"];?>&bdate=<?=$row["bdate"];?>"><img src="i/edit.gif" width="16" height="16" border="0" alt="編輯"></a></td>
			</tr>   
		  <?
             $i++;
			}
			?>
</table>   
