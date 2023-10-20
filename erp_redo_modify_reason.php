<?
  
  session_start();
  $pagtitle = "廠務部 &raquo; Redo/Modify原因"; 
  include("_data.php");
  include("_erp.php");
  //auth("erp_redo_modify_reason.php"); 
  
  if (is_null($_GET['tdate'])) {
    $tdate = date("Y-m-d") ;     
  } else {      
    $tdate = $_GET['tdate'] ;   
  }       
       
  $gem_filter='';
  if ($_GET['gem01']!='') {
      $gemfilter=" and oebud04 = '". $_GET['gem01'] . "' ";
  }    
  
  if ($_POST["action"] == "save") {
      $msg=''; 
      foreach ($_POST["sfbarray"] as $sfb01){
          if ($_POST["risok" . $sfb01] == "Y") {       
              $reason= $_POST["reason" . $sfb01];
              $oeb01= $_POST["oeb01" . $sfb01];  
              $oeb03= $_POST["oeb03" . $sfb01];  
              //vd110
              $soeb= "update oeb_file set oebud06='$reason' where oeb01 = '$oeb01' and oeb03='$oeb03' ";
              $erp_sqloeb=oci_parse($erp_conn1,$soeb);
              $rs=oci_execute($erp_sqloeb);
              if ($rs) {
                  $msg.="VD110 oeb_file $oeb01 $oeb03 更新成功!!";  
              } else {
                  $msg.="VD110 oeb_file訂單檔 $oeb01 $oeb03 更新失敗!!";  
              } 
          }      
      } 
      msg($msg);                                 
                      
      forward("erp_redo_modify_reason.php");                                                              
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
<p>Redo/Modify 原因類別</p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">
            日期:   
            <input name="tdate" type="text" id="tdate" onfocus="WdatePicker()" value="<?=$tdate;?>"> &nbsp;&nbsp; 
            製處:
              <select name="gem01" id="gem01">  
                <option value="">全部</option>
                  <?
                    $socc= "select gem01, gem02 from gem_file where substr(gem01,1,2) in ('69','6A') order by gem01 ";
                    $erp_sqlocc = oci_parse($erp_conn1,$socc );
                    oci_execute($erp_sqlocc);  
                    while ($rowocc = oci_fetch_array($erp_sqlocc, OCI_ASSOC)) {
                        echo "<option value=" . $rowocc["GEM01"];  
                        if ($_GET["gem01"] == $rowocc["GEM01"]) echo " selected";                  
                        echo ">" . $rowocc['GEM01'] ." -- " .$rowocc["GEM02"] . "</option>"; 
                    }   
                  ?>
              </select>   
            <input type="submit" name="Submit2" value="查詢" />
            </div></td>        
        </tr>
    </table>
  </div>
</form>

<?  if (is_null($_GET['Submit2'])) die ;  ?>
<form action="<?=$PHP_SELF;?>" method="post" name="form1">  
  <table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>  
      <th width="16">&nbsp;</th>                         
      <th>RX #</th>
      <th>訂單單號</th> 
      <th>項次</th>
      <th>類別</th>  
      <th>Redo/Modify</th>
      <th>客戶編號</th>   
      <th>客戶名稱</th>   
      <th>產品代號</th>   
      <th>產品名稱</th>   
      <th>部門代號</th> 
      <th>部門名稱</th>  
      <th>原因類別</th> 
      <th width="16"><input type="checkbox" value="checkbox" name='checkall' onclick='checkedAll();'></th>     
    </tr>   
    <?  
    $soea="select ta_oea006, oea01, oeb03, ta_oea004, ta_oea011, oea04, occ02, oeb04, ima02, oebud04, gem02, sfb01, oebud06 " .
          "from oea_file, oeb_file, occ_file, gem_file, sfb_file, ima_file " .
          "where oea02=to_date('$tdate','yy/mm/dd') $gemfilter and ta_oea004 in ('2','3') and oea04=occ01 and oeb04=ima01 and oebud04=gem01 and oea01=oeb01 and sfb22=oea01 and sfb221=oeb03 "; 
    $erp_sqloea = oci_parse($erp_conn1,$soea );
    oci_execute($erp_sqloea); 
    while ($rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC) ) {
        if ($rowoea['TA_OEA011']=='1'){
            $ta_oea011='Crown and Bridge Work';
        } else if ($rowoea['TA_OEA011']=='2'){  
            $ta_oea011='Removable Prosthesis';  
        } else {
            $ta_oea011='Combination Work'; 
        } 
        
        if ($rowoea['TA_OEA004']=='2'){
            $ta_oea004='Redo';               
        } else {
            $ta_oea004='Modify'; 
        }  
        $reason=$rowoea['OEBUD06'];
      ?>
    <tr bgcolor="#FFFFFF">                                
          <td><?=$rxno;?></td>
          <td><?=$rowoea["TA_OEA006"];?></td>
          <td><?=$rowoea["OEA01"];?></td>   
          <td><?=$rowoea["OEB03"];?></td>   
          <td><?=$ta_oea004;?></td>   
          <td><?=$ta_oea011;?></td>    
          <td><?=$rowoea["OEA04"];?></td>    
          <td><?=$rowoea["OCC02"];?></td>
          <td><?=$rowoea["OEB04"];?></td> 
          <td><?=$rowoea["IMA02"];?></td>   
          <td><?=$rowoea["OEBUD04"];?></td>   
          <td><?=$rowoea["GEM02"];?></td>   
          <td>
            <select name="reason<?=$rowoea['SFB01'];?>" id="reason<?=$rowoea['SFB01'];?>" >  
              <option value='0' <? if ($reason=='0') echo " selected" ;?> >無</option>
              <option value='1' <? if ($reason=='1') echo " selected" ;?> >比色</option>       
              <option value='2' <? if ($reason=='2') echo " selected" ;?> >接觸點</option>  
            </select>
          </td>
          <td><input name="risok<?=$rowoea['SFB01'];?>" type="checkbox" id="risok<?=$rowoea['SFB01'];?>" value="Y"> 
              <input name="oeb01<?=$rowoea['SFB01'];?>" type="hidden" value="<?=$rowoea['OEA01'];?>">
              <input name="oeb03<?=$rowoea['SFB01'];?>" type="hidden" value="<?=$rowoea['OEB03'];?>"> 
              <input type="hidden" name="sfbarray[]" value="<?=$rowoea['SFB01'];?>" >
          </td>  
      </tr>     
      
      <?  
      }      
      ?>   
      
      <tr>
        <td>&nbsp;</td>
        <td colspan="13">      
            <input type="hidden" name="action" value="save">
            <input type="submit" name="Submit" value="更新">        
        </td>
      </tr>
  </table>  
</form>
