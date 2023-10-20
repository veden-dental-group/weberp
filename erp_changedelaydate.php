<?
  session_start();
  $pagtitle = "業務部 &raquo; 設定客戶交件日期"; 
  include("_data.php");
  include("_erp.php");
  //auth("erp_changedelaydate.php");   
  date_default_timezone_set('Asia/Taipei');

  

  if ($_POST["action"] == "save") { 
      $rec=0;  
      $oldsfb22='';  
      $account=$_SESSION['account'];
      foreach ($_POST["sfb01array"] as $oea01){          
          
          if ($_POST["risok" . $oea01] == "Y") {
              $rxno       = $_POST["rxno".$oea01];
              $oldoeaud13 = $_POST["oldoeaud13".$oea01];
              $oeaud13    = $_POST["oeaud13".$oea01];   
              //先將要修改的資料記錄下來                                    
              $queryap   = "insert into erp_delaydate_updaterecord ( rxno, oea01, oldate, newdate, username, ip ) values ( 
                          '" . $rxno                              . "',    
                          '" . $oea01                             . "', 
                          '" . $oldoeaud13                        . "', 
                          '" . $oeaud13                           . "', 
                          '" . safetext($account)                 . "',       
                          '" . safetext($_SERVER['REMOTE_ADDR'])  . "')";
              $resultap= mysql_query($queryap) or die ('23 erp_delaydate_updaterecord added error. ' .mysql_error());
                
              $msg='';
              
              //vd110               
              
              $stcohf= "update oea_file set oeaud13=to_date('$oeaud13','yy/mm/dd') where oea01='$oea01' ";
              $erp_sqltcohf=oci_parse($erp_conn1,$stcohf);
              $rs=oci_execute($erp_sqltcohf);   
              if ($rs) {     
                  $rec++;   
              } else {
                  $msg.="VD110  $rxno ($oea01) 更新失敗!!  ";  
              }   
          }      
      } 
      oci_commit($er_conn1);
      if ($msg=='') {  
          msg ("共 $rec 組 cases 更改完畢"); 
      } else {
          msg($msg); 
      } 
      forward("erp_changedelaydate.php");                                                              
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
<p>要設定客戶交件日期的CASE # </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2" id="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">
            客戶: 
            <select name="occ01" id="occ01"> 
                <option value="">全部</option> 
                <?
                  $s1= "select occ01,occ02 from occ_file order by occ01 ";
                  $erp_sql1 = oci_parse($erp_conn1,$s1 );
                  oci_execute($erp_sql1);  
                  while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                      echo "<option value=" . $row1["OCC01"];  
                      if ($_GET["occ01"] == $row1["OCC01"]) echo " selected";                  
                      echo ">" . $row1['OCC01'] ."--" .$row1["OCC02"] . "</option>"; 
                  }   
                ?>
            </select>   &nbsp;&nbsp;   &nbsp;&nbsp; 
            RX #:   
            <input name="rxno" type="text" id="rxno" size="50" value="<?=$_GET['rxno'];?>" maxlength="200">    (若有多個 RX#  請用 , 隔開 )&nbsp;&nbsp;   &nbsp;&nbsp;
            新Due Date:<input name="newduedate" type="text" id="newduedate>" size="12" maxlength="12" value="<?=Date('Y-m-d');?>" onfocus="new WdatePicker()">&nbsp;&nbsp;   &nbsp;&nbsp;
            <input type="submit" name="Submit2" value="送出" />
            </div></td>        
        </tr>
    </table>
  </div>
</form>
<? if (is_null($_GET['Submit2'])) die ; ?>
<form action="<?=$PHP_SELF;?>" method="post" name="form1" id="form1">
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
  <tr>
    <th width="16">&nbsp;</th>
    <th>客戶代號</th>
    <th>客戶名稱</th> 
    <th>RX #</th> 
    <th>訂單日期</th>    
    <th>訂單號碼</th>
    <th>固定/活動</th> 
    <th>原客戶日期</th>    
    <th>新客戶日期</th> 
    <th width="16"><input type="checkbox" value="checkbox" name='checkall' onclick='checkedAll();'></th>
  </tr>
  <?  
      if ($_GET['occ01']!='') {
          $occfilter=" and occ01='" . $_GET['occ01'] . "'";
      } else {
          $occfilter='';
      }
      $rx = $_GET['rxno'];
      $rxarray=explode(',', $rx);
      $max=count($rxarray);
      $rxno='';
      for($i=0; $i<$max; $i++){
          $rxno .= "'" . trim($rxarray[$i]) . "',";                   
      } 
       
      $today = $_GET['newduedate'] ? $_GET['newduedate'] : Date('Y-m-d');
      $ssfb="select oea04, occ02, ta_oea006, to_char(oea02,'yyyy-mm-dd') oea02, oea01, ta_oea011, to_char(oeaud13,'yyyy-mm-dd') oeaud13 from oea_file, occ_file ".
            "where ( ta_oea006 in ($rxno'') or ta_oea006 like '$rx%' ) and oea04=occ01 $occfilter order by oea04,ta_oea006" ;
      $erp_sqlsfb = oci_parse($erp_conn1,$ssfb );
      oci_execute($erp_sqlsfb); 
      while ($rowsfb = oci_fetch_array($erp_sqlsfb, OCI_ASSOC)) { 
        if ($rowsfb['TA_OEA011']=='1') {
            $ta_oea011='固定';
        } else if ($rowsfb['TA_OEA011']=='1') {
            $ta_oea011='活動';
        }  else {
            $ta_oea011='固+活';
        }
          
      ?>    
	    <tr bgcolor="#FFFFFF"> 
		      <td><img src="i/arrow.gif" width="16" height="16">  
              <input type="hidden" name="sfb01array[]" value="<?=$rowsfb['OEA01'];?>"> 
              <input type="hidden" name="rxno<?=$rowsfb['OEA01'];?>" value="<?=$rowsfb['TA_OEA006'];?>"> 
              <input type="hidden" name="oldoeaud13<?=$rowsfb['OEA01'];?>" value="<?=$rowsfb['OEAUD13'];?>">  
          </td>   
			    <td><?=$rowsfb["OEA04"];?></td>
          <td><?=$rowsfb["OCC02"];?></td>  
          <td><?=$rowsfb["TA_OEA006"];?></td> 
          <td><?=$rowsfb["OEA02"];?></td>   
          <td><?=$rowsfb["OEA01"];?></td> 
          <td><?=$ta_oea011;?></td> 
          <td><?=$rowsfb["OEAUD13"];?></td>   
          <td><input name="oeaud13<?=$rowsfb['OEA01'];?>" type="text" id="oeaud13<?=$rowsfb['OEA01'];?>" size="12" maxlength="12" onfocus="new WdatePicker()" value=<?=$today;?>></td>
          <td width=16><input name="risok<?=$rowsfb['OEA01'];?>" type="checkbox" id="risok" value="Y"> </td>  
      </tr> 
      <?  
      }   
  ?>    
  <tr>
    <td>&nbsp;</td>
    <td><input type="hidden" name="rxno"  value="<?=$_GET['rxno'];?>">   
        <input type="hidden" name="action" value="save">
        <input type="submit" name="Submit" value="更改">        
    </td>
  </tr>
</table>  
</form>
