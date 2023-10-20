<?
  session_start();
  $pagtitle = "IT &raquo; 新增客戶傳真"; 
  include("_data.php");
  include("_erp.php");
 // auth("erp_clientfax_add.php");
  

  if ($_POST["action"] == "save") {
      //先刪除所有的資料 再新增      
      foreach ($_POST["oeaarray"] as $oea01){
          if ($_POST["risok" . $oea01] == "Y") {
              $occ01=$_POST['occ01'];
              $socc="select occ01, occ09, occ07, occ44 from occ_file where occ01='$occ01'"; 
              $erp_sqlocc = oci_parse($erp_conn2,$socc );
              oci_execute($erp_sqlocc); 
              $rowocc = oci_fetch_array($erp_sqlocc, OCI_ASSOC);
              $oea03=$rowocc['OCC01'];
              $oea04=$rowocc['OCC09'];  
              $oea17=$rowocc['OCC07'];  
              $oea32=$rowocc['OCC44']; 
              
              $soea= "update oea_file set oea03='$oea03', oea04='$oea04', oea17='$oea17', oea32='$oea32'  where oea01 = '$oea01'";
              $erp_sqloea=oci_parse($erp_conn2,$soea);
              oci_execute($erp_sqloea);
              
              //取出多角流程代碼
              $soea="select oea99 from oea_file where oea01='$oea01'"; 
              $erp_sqloea = oci_parse($erp_conn2,$soea );
              oci_execute($erp_sqloea); 
              $rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC);
              $oea99=$rowoea['OEA99'];
              
              $soea= "update oea_file set oea04='$oea04' where oea99 = '$oea99'";
              $erp_sqloea=oci_parse($erp_conn1,$soea);
              oci_execute($erp_sqloea);  
          }
          //更新價格表
          $soeb= "UPDATE oeb_file a SET a.oeb13=(SELECT xmf07 FROM oea_file,oeb_file b,xmf_file " .
                 "WHERE oea01=b.oeb01 AND xmf01=oea31 AND xmf02=oea23 AND xmf03=b.oeb04 AND xmf07<>b.oeb13 " .
                 "AND a.oeb01=b.oeb01 AND a.oeb03=b.oeb03)  WHERE (a.oeb01,a.oeb03) IN " . 
                 "(SELECT oeb01,oeb03 FROM oea_file,oeb_file,xmf_file WHERE oea01=oeb01 AND xmf01=oea31 AND xmf02=oea23 AND xmf03=oeb04 AND xmf07<>oeb13)";
          $erp_sqloeb=oci_parse($erp_conn2,$soeb);
          oci_execute($erp_sqloeb);  
      }  
      msg ('更新完畢'); 
      forward("erp_changecustomer.php");                                                              
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
<p>新增問題CASE與客戶的FAX記錄. </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">
            RX #:   
            <input name="rxno" type="text" id="rxno" size="100" maxlength="100">  ( 請用 , 隔開)   
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
    <th>RX</th>
    <th>Order No.</th>
    <th>Order Date</th>  
    <th>Due Date</th>
    <th>客戶編號</th>  
    <th>客戶簡稱</th>
    <th>產品編號</th> 
    <th>產品名稱</th> 
    <th>數量</th> 
    <th>傳真日期</th>  
    <th>新帳款代碼</th>
    <th>說明</th>       
    <th width="16"><input type="checkbox" value="checkbox" name='checkall' onclick='checkedAll();'></th>
  </tr>
  <?
  $rxarray=explode(',', $_GET['rxno']);
  $max=count($rxarray);
  $msg='';
  for($i=0; $i<$max; $i++){
      $rxno=$rxarray[$i];
      $soea="select oea01, to_char(oea02,'yyyy/mm/dd') oea02, to_char(oea05,'yyyy/mm/dd') oea05, oea03, occ02, oea04, oea17 from oea_file, occ_file where oea03=occ01 and ta_oea006='$rxno'"; 
      $erp_sqloea = oci_parse($erp_conn2,$soea );
      oci_execute($erp_sqloea);  
      
      while ($rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC)) {
          $occ01=$_GET['occ01'];
          $socc="select occ01, occ09, occ07, occ44 from occ_file where occ01='$occ01'"; 
          $erp_sqlocc = oci_parse($erp_conn2,$socc );
          oci_execute($erp_sqlocc); 
          $rowocc = oci_fetch_array($erp_sqlocc, OCI_ASSOC) 
      ?>    
	    <tr bgcolor="#FFFFFF"> 
		      <td><img src="i/arrow.gif" width="16" height="16">  
              <input type="hidden" name="oeaarray[]" value="<?=$rowoea['OEA01'];?>">  </td>
			    <td><?=$rxno;?></td>
			    <td><?=$rowoea["OEA01"];?></td>
          <td><?=$rowoea["OEA02"];?></td>   
          <td><?=$rowoea["OEA03"];?></td>   
          <td><?=$rowoea["OEA04"];?></td>   
          <td><?=$rowoea["OEA17"];?></td>    
          <td><?=$rowoea["OEA31"];?></td>    
          <td><?=$rowoea["OEA14"];?></td>
          <td><?=$rowocc["OCC01"];?></td> 
          <td><?=$rowocc["OCC09"];?></td>   
          <td><?=$rowocc["OCC07"];?></td>   
          <td><?=$rowocc["OCC44"];?></td>   
          <td><?=$rowoea["OEA99"];?></td> 
          <td width=16><input name="risok<?=$rowoea['OEA01'];?>" type="checkbox" id="risok" value="Y" </td>  
      </tr> 
      <?  
      }
  }
        
  ?>    
  <tr>
    <td>&nbsp;</td>
    <td><input type="hidden" name="occ01"  value="<?=$_GET['occ01'];?>">                      
        <input type="hidden" name="action" value="save">
        <input type="submit" name="Submit" value="更新">        
    </td>
  </tr>
</table>  
</form>
