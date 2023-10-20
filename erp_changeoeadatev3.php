<?
  session_start();
  $pagtitle = "IT &raquo; 更改Order Date & Due Date"; 
  include("_data.php");
  include("_erp.php");
  //auth("erp_changeoeadate.php");
   
  if (is_null($_GET['orderdate'])) {
    $orderdate = '';
  } else {
    $orderdate=$_GET['orderdate'];
  }       
  
  if (is_null($_GET['moredays'])) {
    $moredays =  1;
  } else {
    $moredays=$_GET['moredays'];
  }       
  
  if (is_null($_GET['occ01'])) {
    $occ01 =  '';
  } else {
    $occ01=$_GET['occ01'];
  }   
  
  if (is_null($_GET['rxno'])) {
    $rxno =  '';
  } else {
    $rxno=$_GET['rxno'];
  }  
                               
  if ($_POST["action"] == "save") {
      $neworderdate=$_POST['neworderdate']; 
      $newduedate=$_POST['newduedate'];   
      $moredays = $_POST['moredays'];
      $neworderdate1=substr($neworderdate,2,2).substr($neworderdate,5,2).substr($neworderdate,8,2); 
      $newduedate1=substr($newduedate,2,2).substr($newduedate,5,2).substr($newduedate,8,2); 
      foreach ($_POST["oeaarray"] as $oea01){
          if ($_POST["risok" . $oea01] == "Y") {
              //先將要修改的資料記錄下來                                    
              $queryap   = "insert into erp_orderdate_updaterecord ( oea01, oldoea02, oldta_oea005, newoea02, newta_oea005, username, ip ) values ( 
                          '" . safetext($oea01)                   . "',  
                          '" . $_POST["oea02" . $oea01]            . "',  
                          '" . $_POST["taoea005" . $oea01]        . "',  
                          '" . safetext($neworderdate)            . "',      
                          '" . safetext($newduedate)              . "', 
                          '" . safetext($_SESSION['account'])     . "',       
                          '" . safetext($_SERVER['REMOTE_ADDR'])  . "')";
              $resultap= mysql_query($queryap) or die ('19 erp_orderdate_updaterecord added error. ' .mysql_error());
                                                   
              $msg='';                           
              //要由工單的sfb22, sfb221 去找訂單的單號及項次
              $soea="select oea99 from oea_file where oea01='$oea01'";
              $erp_sqloea = oci_parse($erp_conn2,$soea );
              oci_execute($erp_sqloea); 
              $rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC);
              $oea99=$rowoea["oea99"];   //多角序號   
                                                               
              //vd210  訂單
              $soea= "update oea_file set oea02=oea02+$moredays, ta_oea005= ta_oea005+$moredays where oea01='$oea01'";
              $erp_sqloea=oci_parse($erp_conn2,$soea);
              $rs=oci_execute($erp_sqloea);
              if ($rs) {
                  $msg.="更新成功!!";  
              } else {
                  $msg.="VD210 oea_file訂單檔 $oea01 更新失敗!!";  
              } 
              
              //vd210  採購單
              $soea= "update pmm_file set pmm04= pmm04 + $moredays where pmm99='$oea99'";
              $erp_sqloea=oci_parse($erp_conn2,$soea);
              $rs=oci_execute($erp_sqloea);
              if ($rs) {
                  $msg.="更新成功!!";  
              } else {
                  $msg.="VD210 pmm_file採購檔 $oea01 更新失敗!!";  
              } 
              
              //vd110 訂單 因為一單到底 vd110, vd210的訂單單號是一樣的
              $soea= "update oea_file set oea02=oea02+$moredays, ta_oea005=ta_oea005 + $moredays where oea01='$oea01'";   
              $erp_sqloea=oci_parse($erp_conn1,$soea);
              $rs=oci_execute($erp_sqloea);
              if ($rs) {
                  $msg.="更新成功!!";  
              } else {
                  $msg.="VD110 oea_file訂單檔 $oea01 更新失敗!!";  
              } 
              
              //vd110 工單
              $soea= "update sfb_file set sfb13=sfb13+$moredays, sfb15=sfb15+$moredays where sfb22='$oea01'";   
              $erp_sqloea=oci_parse($erp_conn1,$soea);
              $rs=oci_execute($erp_sqloea);
              if ($rs) {
                  $msg.="更新成功!!";  
              } else {
                  $msg.="VD210 oea_file訂單檔 $oea01 更新失敗!!";  
              } 
                
          }      
      } 
      msg($msg);          
      forward("erp_changeoeadatev2.php");                                                              
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
<p>更改Order Date & Due Date </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">
            客戶:
            <select name="occ01" id="occ01">  
                <option value=''></option>
                <?
                  $s1= "select occ01,occ02 from occ_file order by occ01 ";
                  $erp_sql1 = oci_parse($erp_conn1,$s1 );
                  oci_execute($erp_sql1);  
                  while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                      echo "<option value=" . $row1["OCC01"];  
                      if ($occ01 == $row1["OCC01"]) echo " selected";                  
                      echo ">" . $row1['OCC01'] ."--" .$row1["OCC02"] . "</option>"; 
                  }   
                ?>
            </select> &nbsp; &nbsp;
            Order Date:
            <input name="orderdate" type="text" id="orderdate" size="12" maxlength="12" onfocus="new WdatePicker()" value=<?=$orderdate;?>>&nbsp; &nbsp;
            RX #:   
            <input name="rxno" type="text" id="rxno" size="20" maxlength="20"> &nbsp; &nbsp; &nbsp; &nbsp;       
            
            延幾天出貨:
            <input name="moredays" type="text" id="moredays" size="5" maxlength="5" value=<?=$moredays;?> >&nbsp; &nbsp; 
            
            <input type="submit" name="Submit2" value="送出" />
            </div></td>        
        </tr>
    </table>
  </div>
</form>

<? if (is_null($_GET['Submit2'])) die ; ?>
<form action="<?=$PHP_SELF;?>" method="post" name="form1">
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
  <tr>
    <th width="16">&nbsp;</th>
    <th width="16"><input type="checkbox" value="checkbox" name='checkall' onclick='checkedAll();'></th>
    <th>RX #</th>  
    <th>Order No.</th>
    <th>KEY-IN</th>
    <th>Order Date</th>
    <th>Due Date</th> 
    <th>New Order Date</th>  
    <th>New Due Date</th> 
    
  </tr>
  <?   
      if ($rxno=='') {      
          $rxnofilter='';
      } else  {
          $rxnofilter=" and ta_oea006 like '$rxno%' ";
      }
      
      if ($occ01=='') {      
          $occfilter='';
      } else  {
          $occfilter=" and oea04='$occ01' ";
      }
      
      if ($orderdate==''){
          $orderdatefilter='';
      } else {
          $orderdate1=substr($orderdate,2,2).substr($orderdate,5,2).substr($orderdate,8,2); 
          $orderdatefilter=" and oea02=to_date('$orderdate1','yy/mm/dd') " ;
      }                                                              
      
      $soea="select oea01, to_char(ta_oea005,'yyyy-mm-dd') ta_oea005, to_char(oea02,'yyyy-mm-dd') oea02, ta_oea006, oea14 from vd210.oea_file " .
            "where 1=1 " . $occfilter . $rxnofilter . $orderdatefilter . " order by oea14 ";
      $erp_sqloea = oci_parse($erp_conn2,$soea );
      oci_execute($erp_sqloea); 

      while ($rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC)) { 
        //date('Y-m-d',strtotime("$date1 +5 day"));
        $oea002    = $rowoea['OEA02'] . ' ' . $moredays . ' days';
        $ta_oea005 = $rowoea['TA_OEA005'] . ' ' . $moredays . ' days';
        $neworderdate = date('Y-m-d', strtotime($oea002));
        $newduedate = date('Y-m-d', strtotime($ta_oea005));
      ?>    
	    <tr bgcolor="#FFFFFF"> 
		      <td><img src="i/arrow.gif" width="16" height="16">  
              <input type="hidden" name="oeaarray[]" value="<?=$rowoea['OEA01'];?>">    
          </td>  
          <td width=16><input name="risok<?=$rowoea['OEA01'];?>" type="checkbox" id="risok" value="Y" </td> 
          <td><?=$rowoea["TA_OEA006"];?></td>
          <td><?=$rowoea["OEA01"];?></td>
          <td><?=$rowoea["OEA14"];?></td>
          <td><?=$rowoea["OEA02"];?></td>
          <td><?=$rowoea["TA_OEA005"];?></td> 
          <td><?=$neworderdate;?></td>      
          <td><?=$newduedate;?></td>
          <input name="oea02<?=$rowoea['OEA01'];?>" type="hidden" value="<?=$neworderdate;?>">  
          <input name="taoea005<?=$rowoea['OEA01'];?>" type="hidden" value="<?=$newduedate;?>">

      </tr> 
      <?  
      }   
  ?>    
  <tr>
    <td>&nbsp;</td>
    <td>     
        <input type="hidden" name="neworderdate"  value="<?=$neworderdate;?>">    
        <input type="hidden" name="newduedate"    value="<?=$newduedate;?>">    
        <input type="hidden" name="moredays"      value="<?=$moredays;?>">    
        <input type="hidden" name="action"        value="save">
        <input type="submit" name="Submit"        value="更新">        
    </td>
  </tr>
</table>  
</form>
