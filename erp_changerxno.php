<?
  session_start();
  $pagtitle = "IT &raquo; 更改RX No"; 
  include("_data.php");
  include("_erp.php");
  auth("erp_changerxno.php");
  

  if ($_POST["action"] == "save") {
      $newrxno=$_POST['newrxno']; 
      foreach ($_POST["oeaarray"] as $oea01){
          if ($_POST["risok" . $oea01] == "Y") {
              //先將要修改的資料記錄下來                                    
              $queryap   = "insert into erp_rx_updaterecord ( oldrxno, newrxno, orderno, username, ip ) values ( 
                          '" . safetext($_POST["oldrxno"])        . "',  
                          '" . safetext($_POST["newrxno"])        . "',      
                          '" . safetext($oea01)                   . "', 
                          '" . safetext($_SESSION['account'])     . "',       
                          '" . safetext($_SERVER['REMOTE_ADDR'])  . "')";
              $resultap= mysql_query($queryap) or die ('23 erp_rx_updaterecord added error. ' .mysql_error());
                
              //oea_file: ta_oea006 訂單 vd110, vd210 目前採用一單到底 故vd110的單號和vd210完全相同
              //sfb_file:sfbud02 工單單頭 vd110
              //tc_srg_file: tc_srg002 報工單 vd110               
              //tg_ogb_file: tc_ogb011 掃描秤重 vd110
              
              
              //oga_file: ogaud02 出貨單 vd210, vd110
              //tc_dex_file: tc_dex011 出貨單收費檔 vd110 vd210  -->有問題                                     
              //tc_ofb_file: tc_ofb11 Invoice vd210  
              $msg='';
              
              //vd210
              $soea= "update oea_file set ta_oea006='$newrxno' where oea01 = '$oea01'";
              $erp_sqloea=oci_parse($erp_conn2,$soea);
              $rs=oci_execute($erp_sqloea);
              if ($rs) {
                  $msg.="更新成功!!";  
              } else {
                  $msg.="VD210 oea_file訂單檔 $oea01 更新失敗!!";  
              } 
              
              //vd110
              $soea= "update oea_file set ta_oea006='$newrxno' where oea01 = '$oea01'";
              $erp_sqloea=oci_parse($erp_conn1,$soea);
              $rs=oci_execute($erp_sqloea);
              if ($rs) {
                  $msg.="更新成功!!";  
              } else {
                  $msg.="VD110 oea_file訂單檔 $oea01 更新失敗!!";  
              } 
              
              //vd110
              $ssfb= "update sfb_file set sfbud02='$newrxno' where sfb22 = '$oea01'";
              $erp_sqlsfb=oci_parse($erp_conn1,$ssfb);
              $rs=oci_execute($erp_sqlsfb);
              if ($rs) {
                  $msg.="更新成功!!";  
              } else {
                  $msg.="VD110 sfb_file工單檔 $oea01 更新失敗!!";  
              }                                        
                
              //vd110                                              
              $stcsrg= "update tc_srg_file set tc_srg002='$newrxno' where tc_srg001 in (select sfb01 from sfb_file where sfb22 = '$oea01') ";
              $erp_sqltcsrg=oci_parse($erp_conn1,$stcsrg);                                
              $rs=oci_execute($erp_sqltcsrg);
              if ($rs) {
                  $msg.="更新成功!!";  
              } else {
                  $msg.="VD110 tc_srg_file報工檔 $oea01 更新失敗!!";  
              } 
              
              //vd110 客服扣留
              $stcohf= "update tc_ohf_file set tc_ohf002='$newrxno' where tc_ohf001 in (select sfb01 from sfb_file where sfb22 = '$oea01') ";
              $erp_sqltcohf=oci_parse($erp_conn1,$stcohf);                                
              $rs=oci_execute($erp_sqltcohf);
              if ($rs) {
                  $msg.="更新成功!!";  
              } else {
                  $msg.="VD110 tc_ohf_file客服扣留檔 $oea01 更新失敗!!";  
              } 
                     
              //vd110                                               
              $stcogb= "update tc_ogb_file set tc_ogb011='$newrxno' where tc_ogb003 = '$oea01'";
              $erp_sqltcogb=oci_parse($erp_conn1,$stcogb);
              $rs=oci_execute($erp_sqltcogb); 
              if ($rs) {
                  $msg.="更新成功!!";  
              } else {
                  $msg.="VD110 tc_ogb_file秤重檔 $oea01 更新失敗!!";  
              } 
                                                                 
              //vd110    oga_file: oga16是訂單編號
              $soga= "update oga_file set ogaud02='$newrxno' where oga16 = '$oea01'";
              $erp_sqloga=oci_parse($erp_conn1,$soga);
              $rs=oci_execute($erp_sqloga);
              if ($rs) {
                  $msg.="更新成功!!";  
              } else {
                  $msg.="VD110 oga_file出貨檔 $oea01 更新失敗!!";  
              } 
              
              //vd210   oga_file: 出貨單 oga16是訂單編號     
              $soga= "update oga_file set ogaud02='$newrxno' where oga16 = '$oea01'";
              $erp_sqloga=oci_parse($erp_conn2,$soga);
              $rs=oci_execute($erp_sqloga); 
              if ($rs) {
                  $msg.="更新成功!!";  
              } else {
                  $msg.="VD210 oga_file出貨檔 $oea01 更新失敗!!";  
              }                 
              
              //vd110
              $stcdex= "update tc_dex_file set tc_dex011='$newrxno' where tc_dex001 in ( select oga01 from oga_file where oga16='$oea01')";
              $erp_sqltcdex=oci_parse($erp_conn1,$stcdex);
              $rs=oci_execute($erp_sqltcdex);
              if ($rs) {
                  $msg.="更新成功!!";  
              } else {
                  $msg.="VD110 tc_dex_file收費檔 $oea01 更新失敗!!";  
              }                           
              
              //vd210
              $stcdex= "update tc_dex_file set tc_dex011='$newrxno' where tc_dex001 in ( select oga01 from oga_file where oga16='$oea01')";
              $erp_sqltcdex=oci_parse($erp_conn2,$stcdex);
              $rs=oci_execute($erp_sqltcdex);      
              if ($rs) {
                  $msg.="更新成功!!";  
              } else {
                  $msg.="VD210 tc_dex_file收費檔 $oea01 更新失敗!!";
              }                           
                                              
              //vd210  tc_ofb_file 只有vd210有invoice
              $stcofb= "update tc_ofb_file set tc_ofb11='$newrxno' where tc_ofb31 = '$oea01'";
              $erp_sqltcofb=oci_parse($erp_conn2,$stcofb);
              $rs=oci_execute($erp_sqltcofb);  
              if ($rs) {
                  $msg.="更新成功!! ";
              } else {
                  $msg.="VD110 tc_ofb_file Invoice檔 $oea01 更新失敗!! ";  
              } 
          }      
      } 
      msg($msg); 
      msg ('更新完畢'); 
      forward("erp_changerxno.php");                                                              
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
<p>更改客戶的RX NO. </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">
            錯的 RX #:   
            <input name="oldrxno" type="text" id="oldrxno" size="20" maxlength="20"> &nbsp;  &nbsp;
            對的 RX #:   
            <input name="newrxno" type="text" id="newrxno" size="20" maxlength="20"> &nbsp;  &nbsp; 
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
    <th>錯的RX #</th>
    <th>對的RX #</th> 
    <th>VD210訂單號碼</th>
    <th>Order Date</th>  
    <th>客戶代碼</th>
    <th>客戶姓名</th>  
    <th>Key單人員</th>
    <th>Key單姓名</th>  
    <th>多角貿易代碼</th> 
    <th width="16"><input type="checkbox" value="checkbox" name='checkall' onclick='checkedAll();'></th>
  </tr>
  <?  
      $rxno=$_GET['oldrxno'];
      if ($rxno=='') {
          $rxfilter = " ta_oea006 is null ";
       } else {
          $rxfilter = " ta_oea006='$rxno' ";
       }
      $soea="select oea01, to_char(oea02,'yyyy/mm/dd') oea02, oea03, occ02, oea14, gen02, oea99 from oea_file, occ_file, gen_file 
            where " . $rxfilter . " and oea03=occ01 and oea14=gen01 "; 
      $erp_sqloea = oci_parse($erp_conn2,$soea );
      oci_execute($erp_sqloea); 
      while ($rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC)) {    
      ?>    
	    <tr bgcolor="#FFFFFF"> 
		      <td><img src="i/arrow.gif" width="16" height="16">  
              <input type="hidden" name="oeaarray[]" value="<?=$rowoea['OEA01'];?>">  </td>
			    <td><?=$rxno;?></td>
          <td><?=$_GET['newrxno'];?></td>   
			    <td><?=$rowoea["OEA01"];?></td>
          <td><?=$rowoea["OEA02"];?></td>   
          <td><?=$rowoea["OEA03"];?></td> 
          <td><?=$rowoea["OCC02"];?></td>  
          <td><?=$rowoea["OEA14"];?></td>  
          <td><?=$rowoea["GEN02"];?></td>  
          <td><?=$rowoea["OEA99"];?></td> 
          <? if ($_GET['newrxno']!='') { ?>
                  <td width=16><input name="risok<?=$rowoea['OEA01'];?>" type="checkbox" id="risok" value="Y" </td>  
          <? } else { ?>
                <td>&nbsp; </td> 
          <? } ?>
      </tr> 
      <?  
      }   
  ?>    
  <tr>
    <td>&nbsp;</td>
    <td><input type="hidden" name="newrxno"  value="<?=$_GET['newrxno'];?>">    
        <input type="hidden" name="oldrxno"  value="<?=$_GET['oldrxno'];?>">                     
        <input type="hidden" name="action" value="save">
        <input type="submit" name="Submit" value="更新">        
    </td>
  </tr>
</table>  
</form>
