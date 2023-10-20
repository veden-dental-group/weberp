<?
  session_start();
  $pagtitle = "業務部 &raquo; 更改為New"; 
  include("_data.php");
  include("_erp.php");
  auth("erp_changenewredomodify.php");
  

  if ($_POST["action"] == "save") {
      $ta_oea004=$_POST['ta_oea004']; 
      foreach ($_POST["oeaarray"] as $oea01){
          if ($_POST["risok" . $oea01] == "Y") {
              //先將要修改的資料記錄下來                                    
              $queryap   = "insert into erp_newredomodify_updaterecord ( rxno, oldta_oea004, newta_oea004, orderno, username, ip ) values ( 
                          '" . safetext($_POST["rxno"])               . "',  
                          '" . safetext($_POST["oldta_oea004".$oea01]). "',
                          '" . safetext($_POST["ta_oea004"])          . "',      
                          '" . safetext($oea01)                       . "', 
                          '" . safetext($_SESSION['account'])         . "',       
                          '" . safetext($_SERVER['REMOTE_ADDR'])      . "')";
              $resultap= mysql_query($queryap) or die ('23 erp_rewredomodify_updaterecord added error. ' .mysql_error());
                
              //oea_file: ta_oea003 訂單 vd110, vd210 目前採用一單到底 故vd110的單號和vd210完全相同                 
              $msg='';
              
              //vd210
              $soea= "update oea_file set ta_oea004='$ta_oea004' where oea01 = '$oea01'";
              $erp_sqloea=oci_parse($erp_conn2,$soea);
              $rs=oci_execute($erp_sqloea);
              if ($rs) {
                  $msg.="更新成功!!";  
              } else {
                  $msg.="VD210 oea_file訂單檔 $oea01 更新失敗!!";  
              } 
              
              //vd110
              $soea= "update oea_file set ta_oea004='$ta_oea004' where oea01 = '$oea01'";
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
      forward("erp_changenewredomodify.php");                                                              
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
<p>更改訂單為New Redo Modify. </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">
            RX #:   
            <input name="rxno" type="text" id="rxno" size="20" maxlength="100" onfocus="this.select();" value="<?=$_GET['rxno'];?>"> &nbsp;  &nbsp;
            新的訂單類別:               
            <select name="ta_oea004" id="ta_oea004"> 
                <option value="1" <? if($_GET['ta_oea004']=='1') echo " selected"; ?> >New</option>
                <option value="2" <? if($_GET['ta_oea004']=='2') echo " selected"; ?> >Redo</option>
                <option value="3" <? if($_GET['ta_oea004']=='3') echo " selected"; ?> >Modify</option>
                <option value="4" <? if($_GET['ta_oea004']=='4') echo " selected"; ?> >內返</option>
            </select>
            &nbsp;  &nbsp; 
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
    <th>原訂單類別</th> 
    <th>新訂單類別</th> 
    <th width="16"> &nbsp; </th>
  </tr>
  <?  
      if ($_GET['ta_oea004']=='1') {
          $newta_oea004='New';
      } else if ($_GET['ta_oea004']=='2') {
          $newta_oea004='Redo';
      } else if ($_GET['ta_oea004']=='3') {
          $newta_oea004='Modify';
      } else if ($_GET['ta_oea004']=='4') {
          $newta_oea004='內返';    
      }
      $rxno=$_GET['rxno']; 
      $rxfilter = " ta_oea006='$rxno' ";
      $soea="select oea01, to_char(oea02,'yyyy/mm/dd') oea02, ta_oea003, oea03, occ02, oea14, gen02, oea99, ta_oea004 from oea_file, occ_file, gen_file 
            where $rxfilter and oea03=occ01 and oea14=gen01 "; 
      $erp_sqloea = oci_parse($erp_conn2,$soea );
      oci_execute($erp_sqloea); 
      while ($rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC)) {    
          if ($rowoea['TA_OEA004']=='1') {
              $ta_oea004='New';
          } else if ($rowoea['TA_OEA004']=='2') {
              $ta_oea004='Redo';
          } else if ($rowoea['TA_OEA004']=='3') {
              $ta_oea004='Modify';
          } else if ($rowoea['TA_OEA004']=='4') {
              $ta_oea004='內返';    
          }
      ?>    
	    <tr bgcolor="#FFFFFF"> 
		      <td><img src="i/arrow.gif" width="16" height="16">  
              <input type="hidden" name="oeaarray[]" value="<?=$rowoea['OEA01'];?>"> 
              <input type="hidden" name="oldta_oea004<?=$rowoea['OEA01'];?>"  id="oldta_oea004<?=$rowoea['OEA01'];?>" value="<?=$rowoea['TA_OEA004'];?>"> 
          </td>
			    <td><?=$rxno;?></td>              
			    <td><?=$rowoea["OEA01"];?></td>
          <td><?=$rowoea["OEA02"];?></td>   
          <td><?=$rowoea["OEA03"];?></td> 
          <td><?=$rowoea["OCC02"];?></td>  
          <td><?=$rowoea["OEA14"];?></td>  
          <td><?=$rowoea["GEN02"];?></td>  
          <td><?=$rowoea["OEA99"];?></td> 
          <td><?=$ta_oea004;?></td>   
          <td><?=$newta_oea004;?></td>   
          <td width=16> <input name="risok<?=$rowoea['OEA01'];?>" type="checkbox" id="risok" value="Y"> </td>  
      </tr> 
      <?  
      }   
  ?>    
  <tr>
    <td>&nbsp;</td>
    <td><input type="hidden" name="ta_oea004"  value="<?=$_GET['ta_oea004'];?>">   
        <input type="hidden" name="rxno"  value="<?=$_GET['rxno'];?>">  
        <input type="hidden" name="action" value="save">
        <input type="submit" name="Submit" value="更新">        
    </td>
  </tr>
</table>  
</form>
