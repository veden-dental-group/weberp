<?
  session_start();
  $pagtitle = "業務部 &raquo; 更改Doctor";
  include("_data.php");
  include("_erp.php");
  auth("erp_changedoctor.php");
  

  if ($_POST["action"] == "save") {
      $newpatient=$_POST['newpatient']; 
      foreach ($_POST["oeaarray"] as $oea01){
          if ($_POST["risok" . $oea01] == "Y") {
              //先將要修改的資料記錄下來                                    
              $queryap   = "insert into erp_patient_updaterecord ( rxno, oldpatient, newpatient, orderno, username, ip ) values ( 
                          '" . safetext($_POST["rxno"])               . "',  
                          '" . safetext($_POST["oldpatient".$oea01])  . "',
                          '" . safetext($_POST["newpatient"])         . "',      
                          '" . safetext($oea01)                       . "', 
                          '" . safetext($_SESSION['account'])         . "',       
                          '" . safetext($_SERVER['REMOTE_ADDR'])      . "')";
              $resultap= mysql_query($queryap) or die ('23 erp_patient_updaterecord added error. ' .mysql_error());
                
              //oea_file: ta_oea003 訂單 vd110, vd210 目前採用一單到底 故vd110的單號和vd210完全相同                 
              $msg='';
              
              //vd210
              $soea= "update oea_file set ta_oea002='$newpatient' where oea01 = '$oea01'";
              $erp_sqloea=oci_parse($erp_conn2,$soea);
              $rs=oci_execute($erp_sqloea);
              if ($rs) {
                  $msg.="更新成功!!";  
              } else {
                  $msg.="VD210 oea_file訂單檔 $oea01 更新失敗!!";  
              } 
              
              //vd110
              $soea= "update oea_file set ta_oea002='$newpatient' where oea01 = '$oea01'";
              $erp_sqloea=oci_parse($erp_conn1,$soea);
              $rs=oci_execute($erp_sqloea);
              if ($rs) {
                  $msg.="更新成功!!";  
              } else {
                  $msg.="VD110 oea_file訂單檔 $oea01 更新失敗!!";  
              }    
          }      
      } 
      msg($msg); 
      msg ('更新完畢'); 
      forward("erp_changedoctor.php");
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
<p>更改客戶的Doctor. </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">
            RX #:   
            <input name="rxno" type="text" id="rxno" size="20" maxlength="20"> &nbsp;  &nbsp;
            對的 Doctor:
            <input name="newpatient" type="text" id="newpatient" size="20" maxlength="20"> &nbsp;  &nbsp; 
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
    <th>RX #</th>
    
    <th>VD210訂單號碼</th>
    <th>Order Date</th>  
    <th>客戶代碼</th>
    <th>客戶姓名</th>  
    <th>Key單人員</th>
    <th>Key單姓名</th>  
    <th>多角貿易代碼</th> 
    <th>原Doctor</th>
    <th>新Doctor</th>
    <th width="16"> &nbsp; </th>
  </tr>
  <?  
      $rxno=$_GET['rxno']; 
      $rxfilter = " ta_oea006='$rxno' ";
      $soea="select oea01, to_char(oea02,'yyyy/mm/dd') oea02, ta_oea002, oea03, occ02, oea14, gen02, oea99 from oea_file, occ_file, gen_file 
            where $rxfilter and oea03=occ01 and oea14=gen01 "; 
      $erp_sqloea = oci_parse($erp_conn2,$soea );
      oci_execute($erp_sqloea); 
      while ($rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC)) {    
      ?>    
	    <tr bgcolor="#FFFFFF"> 
		      <td><img src="i/arrow.gif" width="16" height="16">  
              <input type="hidden" name="oeaarray[]" value="<?=$rowoea['OEA01'];?>"> 
              <input type="hidden" name="oldpatient<?=$rowoea['OEA01'];?>"  id="oldpatient<?=$rowoea['OEA01'];?>" value="<?=$rowoea['TA_OEA003'];?>"> 
          </td>
			    <td><?=$rxno;?></td>
            
			    <td><?=$rowoea["OEA01"];?></td>
          <td><?=$rowoea["OEA02"];?></td>   
          <td><?=$rowoea["OEA03"];?></td> 
          <td><?=$rowoea["OCC02"];?></td>  
          <td><?=$rowoea["OEA14"];?></td>  
          <td><?=$rowoea["GEN02"];?></td>  
          <td><?=$rowoea["OEA99"];?></td> 
          <td><?=$rowoea["TA_OEA002"];?></td>
          <td><?=$_GET['newpatient'];?></td>   
          <td width=16> <input name="risok<?=$rowoea['OEA01'];?>" type="checkbox" id="risok" value="Y"> </td>  
      </tr> 
      <?  
      }   
  ?>    
  <tr>
    <td>&nbsp;</td>
    <td><input type="hidden" name="newpatient"  value="<?=$_GET['newpatient'];?>">   
        <input type="hidden" name="rxno"  value="<?=$_GET['rxno'];?>">  
        <input type="hidden" name="action" value="save">
        <input type="submit" name="Submit" value="更新">        
    </td>
  </tr>
</table>  
</form>
