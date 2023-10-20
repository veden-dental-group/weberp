<?
  session_start();
  $pagtitle = "業務部 &raquo; 更改 Shade"; 
  include("_data.php");
  include("_erp.php");
  //auth("erp_changeshade.php");
  
  if (is_null($_GET['bdate'])) {
    // $bdate = date('Y-m-d');
    $bdate = '';
  } else {
    $bdate=$_GET['bdate'];
  }                              
  

  if ($_POST["action"] == "save") {
      
      foreach ($_POST["oeaarray"] as $oea01){
          if ($_POST["risok" . $oea01] == "Y") {
              $oldshade = $_POST["oldshade".$oea01];
              $newshade = $_POST["newshade".$oea01];  
              $rxno = $_POST["rxno".$oea01];
              
              //先將要修改的資料記錄下來                                    
              $queryap   = "insert into erp_shade_updaterecord ( rxno, oldshade, newshade, orderno, username, ip ) values ( 
                          '" . $rxno                                  . "',  
                          '" . $oldshade                              . "',
                          '" . $newshade                              . "',     
                          '" . safetext($oea01)                       . "', 
                          '" . safetext($_SESSION['account'])         . "',       
                          '" . safetext($_SERVER['REMOTE_ADDR'])      . "')";
              $resultap= mysql_query($queryap) or die ('23 erp_shade_updaterecord added error. ' .mysql_error());
                
              //oea_file: ta_oea003 訂單 vd110, vd210 目前採用一單到底 故vd110的單號和vd210完全相同                 
              $msg='';
              
              //vd210
              $soea= "update oea_file set ta_oea047='$newshade' where oea01 = '$oea01'";
              $erp_sqloea=oci_parse($erp_conn2,$soea);
              $rs=oci_execute($erp_sqloea);
              if ($rs) {
                  $msg.="更新成功!!";  
              } else {
                  $msg.="VD210 oea_file訂單檔 $oea01 更新失敗!!";  
              } 
              
              //vd110
              $soea= "update oea_file set ta_oea047='$newshade' where oea01 = '$oea01'";
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
      forward("erp_changeshade.php");                                                              
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
<p>更改訂單的 Shade. </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">
            送貨客戶: 
            <select name="occ01" id="occ01">  
                <option value=''></option>
                <?
                  $s1= "select occ01,occ02 from occ_file where occ01 in ( select decode(occud03,'Y',occ07, occ09) from occ_file ) group by occ01, occ02 order by occ01 ";
                  $erp_sql1 = oci_parse($erp_conn1,$s1 );
                  oci_execute($erp_sql1);  
                  while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                      echo "<option value=" . $row1["OCC01"];  
                      if ($_GET["occ01"] == $row1["OCC01"]) echo " selected";                  
                      echo ">" . $row1['OCC01'] ."--" .$row1["OCC02"] . "</option>"; 
                  }   
                ?>
            </select>   &nbsp;&nbsp;   &nbsp;&nbsp;
            到貨日期:   
            <input name="bdate" type="text" id="bdate" size="12" maxlength="12" onfocus="new WdatePicker()"  value=<?=$bdate;?>>             
            訂單號碼:   
            <input name="oea01" type="text" id="oea01" size="20" maxlength="40"> &nbsp;  &nbsp;               
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
    <th>VD210訂單號碼</th>
    <th>RX #</th>  
    <th>Order Date</th>  
    <th>客戶代碼</th>
    <th>客戶姓名</th>  
    <th>Key單人員</th>
    <th>Key單姓名</th>  
    <th>多角貿易代碼</th> 
    <th>舊Sahde</th> 
    <th>新Shade</th> 
    <th width="16"> &nbsp; </th>
  </tr>
  <?     
	// echo $_GET['occ01'];   
      if ($_GET['occ01']=='') {
          $occ01filter='';
      } else {
          $occ01filter=" and oea04='" . $_GET['occ01'] . "' "; 
      }
	
	// modified by mao 10:14
	if (!$_GET['bdate']) {
		$sdate = date('Y-m-').'01';
		$edate = date('Y-m-d');
		$oea02filter=" and oea02 between to_date('" . $sdate . "','yy/mm/dd') and to_date('" . $edate . "','yy/mm/dd') "; 
	} else {
		$bdate = $_GET['bdate']; 
		$oea02filter=" and oea02=to_date('" . $bdate . "','yy/mm/dd') "; 
	}
	// $oea02filter=" and oea02=to_date('" .  . "','yy/mm/dd') "; 
     
	/* 
      if ($_GET['bdate']=='') {
          $oea02filter='';
      } else {
          $oea02filter=" and oea02=to_date('" . $_GET['bdate'] . "','yy/mm/dd') "; 
      }
	*/
      
      if ($_GET['oea01']=='') {
          $oea01filter='';
      } else {
          $oea01filter=" and oea01='" . $_GET['oea01'] . "' "; 
      }
      
      if (($_GET['oea01'] . $_GET['bdate']. $_GET['occ01'])=='' ) {
          $limit = ' and 1 != 1';
      } else {
          $limit ='';
      }
      
      $soea="select oea01, to_char(oea02,'yyyy/mm/dd') oea02, ta_oea003, ta_oea006, oea04, occ02, oea14, gen02, oea99, ta_oea047 from oea_file, occ_file, gen_file 
            where oea04=occ01 and oea14=gen01 $occ01filter $oea02filter $oea01filter $limit order by oea01 "; 
      $erp_sqloea = oci_parse($erp_conn2,$soea );
      oci_execute($erp_sqloea); 
      while ($rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC)) {    
      ?>    
	    <tr bgcolor="#FFFFFF"> 
		      <td><img src="i/arrow.gif" width="16" height="16">  
              <input type="hidden" name="oeaarray[]" value="<?=$rowoea['OEA01'];?>"> 
              <input type="hidden" name="oldshade<?=$rowoea['OEA01'];?>"  id="oldshade<?=$rowoea['OEA01'];?>" value="<?=$rowoea['TA_OEA047'];?>"> 
              <input type="hidden" name="rxno<?=$rowoea['OEA01'];?>"  id="rxno<?=$rowoea['OEA01'];?>" value="<?=$rowoea['TA_OEA006'];?>">   
          </td>                	     
			    <td><?=$rowoea["OEA01"];?></td>
          <td><?=$rowoea["TA_OEA006"];?></td> 
          <td><?=$rowoea["OEA02"];?></td>   
          <td><?=$rowoea["OEA04"];?></td> 
          <td><?=$rowoea["OCC02"];?></td>  
          <td><?=$rowoea["OEA14"];?></td>  
          <td><?=$rowoea["GEN02"];?></td>  
          <td><?=$rowoea["OEA99"];?></td> 
          <td><?=$rowoea["TA_OEA047"];?></td>   
          <td width=16> <input name="newshade<?=$rowoea['OEA01'];?>" type="text" id="newshade<?=$rowoea['OEA01'];?>" value="<?=$rowoea["TA_OEA047"];?>"> </td>  
          <td width=16> <input name="risok<?=$rowoea['OEA01'];?>" type="checkbox" id="risokrisok<?=$rowoea['OEA01'];?>" value="Y"> </td>  
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
