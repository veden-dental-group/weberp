<?
  session_start();
  $pagtitle = "IT &raquo; 更改Delay 狀態V2"; 
  include("_data.php");
  include("_erp.php");
  auth("erp_changedelaycasev2.php");
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
          $maker  =$_POST['maker'.$pkey];
          $status =$_POST['status'.$pkey];
          $tdate  =$_POST['tdate'.$pkey]; 
          if ($_POST["risok" . $pkey] == "Y") {
              //先將要修改的資料記錄下來                                    
              $queryap   = "insert into erp_changedelaycase_updaterecord ( updatepkey, oldmaker, newmaker, olddelay, newdelay, username, ip ) values ( 
                          '" . safetext($pkey)                          . "',  
                          '" . safetext($_POST["oldmakercode".$pkey])   . "', 
                          '" . safetext($_POST["maker".$pkey])      . "',   
                          '" . safetext($_POST["oldstatus".$pkey])      . "',   
                          '" . safetext($_POST["status".$pkey])         . "',   
                          '" . safetext($_SESSION['account'])           . "',       
                          '" . safetext($_SERVER['REMOTE_ADDR'])        . "')";
              $resultap= mysql_query($queryap) or die ('23 erp_changedelaycase_updaterecord added error. ' .mysql_error());       
              $makercode=substr($maker,0,6);
              $makername=substr($maker,7);
              $s2="update delaydetail set makercode='$makercode', makername='$makername' where pkey='$pkey' limit 1"  ;        
              $r2=mysql_query($s2) or die ('38 DelayDetail updated error!!'. mysql_error());
              
              $s1="update delay set status='$status' where tdate='$tdate'and orderno=(select orderno from delaydetail where pkey='$pkey') limit 1 "  ;        
              $r1=mysql_query($s1) or die ('41 Delay updated error!!'. mysql_error());  
              
          }      
      }          
      msg ('更新完畢'); 
      forward("erp_changedelaycasev2.php");                                                              
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
            出貨日期/Delay日期:   
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
    <th>應出貨日期</th> 
    <th>品代</th>  
    <th>品名</th>
    <th>客戶別</th>  
    <th>客戶名</th>  
    <th>顆數</th>   
    <th>業績數</th>   
    <th>製處</th>        
    <th>Delay原因</th>  
    <th width="16"><input type="checkbox" value="checkbox" name='checkall' onclick='checkedAll();'></th>
  </tr>
  <?            
      if ($rxno=='') {
          $rxfilter = " ";
      } else {
          $rxfilter = " and d.rx='$rxno' ";
      }
      if ($_GET['maker']=='') {
          $makerfilter = " ";
      } else {
          $makerfilter = " and dd.makercode='". $_GET['maker'] . "' ";
      }
      if ($tdate=='') {
          $tdatefilter = " ";
      } else {
          $tdatefilter = " and d.tdate='$tdate'  ";
      }
       
      $s1="select dd.pkey pkey, d.rx rx, d.orderno orderno, dd.ticketno ticketno, d.orderdate orderdate, d.duedate duedate, dd.productcode productcode, dd.productname productname, d.clientid clientid, d.clientname clientname, dd.makercode makercode, dd.makername makername, d.status status, d.tdate tdate, dd.qty qty, dd.plus plus   ".
          "from delay d, delaydetail dd " .
          "where d.orderno=dd.orderno and d.tdate=dd.tdate and d.tdate>=d.duedate " . $tdatefilter . $rxfilter . $makerfilter .
          "order by dd.makercode, d.rx, d.orderdate, d.duedate "; 
      $r1=mysql_query($s1) or die ('241 casedelay error!!' . mysql_error());   
      $i=0; 
      while ($row1 = mysql_fetch_array($r1)) {    
          $i++;
      ?>    
	      <tr bgcolor="#FFFFFF"> 
		      <td><?=$i;?> 
              <input type="hidden" name="pkeyarray[]" value="<?=$row1['pkey'];?>">  
              <input type="hidden" name="oldmakercode<?=$row1['pkey'];?>" value="<?=$row1['makercode'];?>">
              <input type="hidden" name="oldstatus<?=$row1['pkey'];?>" value="<?=$row1['status'];?>">  
              <input type="hidden" name="tdate<?=$row1['pkey'];?>" value="<?=$row1['tdate'];?>">  
          </td>                             
          <td><?=$row1["rx"];?></td>   
          <td><?=$row1["orderno"];?></td>            
          <td><?=$row1["ticketno"];?></td> 
          <td><?=$row1["orderdate"];?></td>
          <td><?=$row1["duedate"];?></td>  
          <td><?=$row1["productcode"];?></td>  
          <td><?=$row1["productname"];?></td>    
          <td><?=$row1["clientid"];?></td> 
          <td><?=$row1["clientname"];?></td> 
          <td><?=$row1["qty"];?></td> 
          <td><?=$row1["plus"];?></td> 
          <td>
            <select name="maker<?=$row1['pkey'];?>" id="maker<?=$row1['pkey'];?>">    
                <?
                  $s2= "select mid, mname from maker order by mid ";
                  $r2=mysql_query($s2) or die ('186 Makeer error!!' . mysql_error());    
                  while ($row2 = mysql_fetch_array($r2)) {
                      echo "<option value=" . $row2["mid"].','.$row2['mname'];  
                      if ($row1['makercode'] == $row2["mid"]) echo " selected";                  
                      echo ">" . $row2['mid'] ."  " .$row2["mname"] . "</option>"; 
                  }   
                ?>
            </select> &nbsp; &nbsp;
          </td>                         
          <td>
            <select name="status<?=$row1['pkey'];?>" id="status<?=$row1['pkey'];?>">  
                <option value=''>Delay</option>                                                    
                <option value='4' <? if($row1['status']=='4') echo " selected";?> >傳真扣留未刷</option> 
                <option value='6' <? if($row1['status']=='6') echo " selected";?> >高難度顆數多</option> 
                <option value='5' <? if($row1['status']=='5') echo " selected";?> >其他</option> 
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
