<?
  session_start();
  $pagtitle = "IT &raquo; 更改Delay 狀態"; 
  include("_data.php");
  include("_erp.php");
  auth("erp_changedelaycase.php");
  if (is_null($_GET['tdate'])) {
    $tdate = date('Y-m-d');
  } else {
    $tdate=$_GET['tdate'];
  }    
  
  if (is_null($_GET['rxno'])) {
    $rxno = '';
  } else {
    $rxno=$_GET['rxno'];
  }    

  if ($_POST["action"] == "save") { 
      foreach ($_POST["pkeyarray"] as $pkey){
          $d13=$_POST['d13'.$pkey];
          $d16=$_POST['d16'.$pkey];
          if ($_POST["risok" . $pkey] == "Y") {
              //先將要修改的資料記錄下來                                    
              $queryap   = "insert into erp_changedelaycase_updaterecord ( updatepkey, oldmaker, newmaker, olddelay, newdelay, username, ip ) values ( 
                          '" . safetext($pkey)                    . "',  
                          '" . safetext($_POST["oldd13".$pkey])   . "', 
                          '" . safetext($_POST["d13".$pkey])   . "',   
                          '" . safetext($_POST["oldd16".$pkey])   . "',   
                          '" . safetext($_POST["d16".$pkey])   . "',   
                          '" . safetext($_SESSION['account'])     . "',       
                          '" . safetext($_SERVER['REMOTE_ADDR'])  . "')";
              $resultap= mysql_query($queryap) or die ('23 erp_changedelaycase_updaterecord added error. ' .mysql_error());       
              $d131=substr($d13,0,6);
              $d132=substr($d13,7);
              $s1="update casedelay set d13='$d131', d14='$d132', d16='$d16' where pkey='$pkey'"  ;        
              $r1=mysql_query($s1) or die ('37 Casedelay updated error!!'. mysql_error());
          }      
      }          
      msg ('更新完畢'); 
      forward("erp_changedelaycase.php");                                                              
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
<p>更改case delay 狀態 </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">
            到貨日期:   
            <input name="tdate" type="text" id="tdate" size="12" maxlength="12" onfocus="new WdatePicker()" value=<?=$tdate;?>>&nbsp; &nbsp;  
            製處:   
            <select name="maker" id="maker">  
                <option value=''>全部製處</option>
                <?
                  $s1= "select mid, mname from maker order by mid ";
                  $r1=mysql_query($s1) or die ('186 Makeer error!!' . mysql_error());    
                  while ($row1 = mysql_fetch_array($r1)) {
                      echo "<option value=" . $row1["mid"];  
                      if ($_GET['maker'] == $row1["mid"]) echo " selected";                  
                      echo ">" . $row1['mid'] ."  " .$row1["mname"] . "</option>"; 
                  }   
                ?>
            </select> &nbsp; &nbsp;
            RX #:
            <input name="rxno" type="text" id="rxno" size="20" maxlength="20"> &nbsp;  &nbsp; 
            <input type="submit" name="Submit2" value="送出" />
            </div></td>        
        </tr>
    </table>
  </div>
</form>

<? if (is_null($_GET['Submit2'])) die ; ?> 
<? if (($rxno.$_GET['maker'].$_GET['tdate'])=='') die ; ?>   

<form action="<?=$PHP_SELF;?>" method="post" name="form1">
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
  <tr>
    <th width="16">&nbsp;</th>
    <th>RX #</th>       
    <th>訂單號碼</th>    
    <th>工單號碼</th> 
    <th>到貨日期</th>   
    <th>Delay日期</th> 
    <th>品代</th>  
    <th>品名</th>
    <th>客戶別</th>  
    <th>客戶名</th> 
    <th>製處代號</th>
    <th>Delay原因</th>  
    <th width="16"><input type="checkbox" value="checkbox" name='checkall' onclick='checkedAll();'></th>
  </tr>
  <?            
      if ($rxno=='') {
          $rxfilter = " ";
      } else {
          $rxfilter = " and d02='$rxno' ";
      }
      if ($_GET['maker']=='') {
          $makerfilter = " ";
      } else {
          $makerfilter = " and d13='". $_GET['maker'] . "' ";
      }
      if ($tdate=='') {
          $tdatefilter = " ";
      } else {
          $tdatefilter = " and d01='$tdate' ";
      }
       
      $s1="select pkey, d01, d02, d03, d031, d04, d09, d10, d11, d12, d13, d14, d16,d21 from casedelay " .
            "where 1=1 " . $tdatefilter . $rxfilter . $makerfilter .
            "order by d02, d03, d031,d04,d01 "; 
      $r1=mysql_query($s1) or die ('241 casedelay error!!' . mysql_error());   
      $i=0; 
      while ($row1 = mysql_fetch_array($r1)) {    
          $i++;
      ?>    
	      <tr bgcolor="#FFFFFF"> 
		      <td><?=$i;?> 
              <input type="hidden" name="pkeyarray[]" value="<?=$row1['pkey'];?>">  
              <input type="hidden" name="oldd13<?=$row1['pkey'];?>" value="<?=$row1['d13'];?>">
              <input type="hidden" name="oldd16<?=$row1['pkey'];?>" value="<?=$row1['d16'];?>">  
          </td> 
          <td><?=$row1["d02"];?></td> 
          <td><?=$row1["d03"]. " " .$row1['d031'];?></td>  
          <td><?=$row1["d04"];?></td>  
          <td><?=$row1["d01"];?></td> 
          <td><?=$row1["d21"];?></td>
          <td><?=$row1["d09"];?></td>  
          <td><?=$row1["d10"];?></td>  
          <td><?=$row1["d11"];?></td>    
          <td><?=$row1["d12"];?></td> 
          <td>
            <select name="d13<?=$row1['pkey'];?>" id="d13<?=$row1['pkey'];?>">    
                <?
                  $s2= "select mid, mname from maker order by mid ";
                  $r2=mysql_query($s2) or die ('186 Makeer error!!' . mysql_error());    
                  while ($row2 = mysql_fetch_array($r2)) {
                      echo "<option value=" . $row2["mid"].','.$row2['mname'];  
                      if ($row1['d13'] == $row2["mid"]) echo " selected";                  
                      echo ">" . $row2['mid'] ."  " .$row2["mname"] . "</option>"; 
                  }   
                ?>
            </select> &nbsp; &nbsp;
          </td>                         
          <td>
            <select name="d16<?=$row1['pkey'];?>" id="d16<?=$row1['pkey'];?>">  
                <option value=''>Delay</option> 
                <option value='1' <? if($row1['d16']=='1') echo " selected";?> >六顆以上</option> 
                <option value='2' <? if($row1['d16']=='2') echo " selected";?> >三種產品</option> 
                <option value='3' <? if($row1['d16']=='3') echo " selected";?> >固定+活動</option> 
                <option value='4' <? if($row1['d16']=='4') echo " selected";?> >傳真扣留未刷</option> 
                <option value='5' <? if($row1['d16']=='5') echo " selected";?> >其他</option> 
            </select>
          </td> 
          <td width=16><input name="risok<?=$row1['pkey'];?>" type="checkbox" id="risok<?=$row1['pkey'];?>"" value="Y"></td>  
      </tr> 
      <?  
      }   
  ?>    
  <tr>
    <td>&nbsp;</td>
    <td>                 
        <input type="hidden" name="action" value="save">
        <input type="submit" name="Submit" value="更新">        
    </td>
  </tr>
</table>  
</form>
