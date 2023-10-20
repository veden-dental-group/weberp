<?
  session_start();
  $pagtitle = "業務部 &raquo; 解傳真作業"; 
  include("_data.php");
  include("_erp.php");
  //auth("erp_faxoutv1.php");
  date_default_timezone_set('Asia/Taipei');
  

  if ($_POST["action"] == "save") { 
      $msg='';  
      $rec=0;  
      $oldsfb22='';
      $account=$_SESSION['account'];
      foreach ($_POST["sfbarray"] as $sfb01){
          if ($_POST["risok" . $sfb01] == "Y") {
              $rxno=$_POST["rxno".$sfb01];
              //先將要修改的資料記錄下來                                    
              $queryap   = "insert into erp_faxin_updaterecord ( rxno, sfb01, faxtype, username, ip ) values ( 
                          '" . $rxno                              . "',    
                          '" . $sfb01                             . "', 'out',
                          '" . safetext($account)                 . "',       
                          '" . safetext($_SERVER['REMOTE_ADDR'])  . "')";
              $resultap= mysql_query($queryap) or die ('23 erp_faxin_updaterecord added error. ' .mysql_error()); 
              //vd110  
              //vd110               
              $ssfb="select sfbud02, sfb05, sfb22  from sfb_file where sfb01='$sfb01'" ; 
              $erp_sqlsfb=oci_parse($erp_conn1,$ssfb);
              oci_execute($erp_sqlsfb); 
              $rowsfb = oci_fetch_array($erp_sqlsfb, OCI_ASSOC);
              $sfbud02=$rowsfb["SFBUD02"];
              $sfb05=$rowsfb["SFB05"];    
              $sfb22=$rowsfb['SFB22'];            
              
              $indate=$_POST["faxdate".$sfb01];  
              $intime=date('G:i:s');
              $memo=$_POST["reason".$sfb01].$_POST["rmemo".$sfb01]; 
              $stcohf= "update tc_ohf_file set tc_ohf008=to_date('$indate','yy/mm/dd'), tc_ohf009='$intime', tc_ohf010='$account', tc_ohf011='$memo' where tc_ohf001='$sfb01' and tc_ohf008 is null ";
              $erp_sqltcohf=oci_parse($erp_conn1,$stcohf);
              $rs=oci_execute($erp_sqltcohf);   
              if ($rs) {  
                  if ( $sfb22 != $oldsfb22 ) {
                      $rec++;
                      $oldsfb22=$sfb22; 
                  }                                   
              } else {
                  $msg.="VD110 傳真扣留檔 $sfbud02 ($sfb01) 更新失敗!!";  
              } 
          }      
      } 
      if ($msg=='') {  
          msg ("共 $rec 組 cases 解扣完畢"); 
      } else {
          msg($msg); 
      }
      forward("erp_faxoutv1.php");
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

function validate_form(){
    var errormsg = ''
    for (var i = 0; i < document.form1.elements.length; i++) {
        var e = document.form1.elements[i]
        if (e.type == "checkbox" && e.checked && e.name.substr(0,5)=='risok' ) {
            var reason = 'reason' + e.name.substr(5)
            var memo = 'rmemo' + e.name.substr(5)
            if (document.getElementById(reason).value=='') {
                errormsg += "請選擇原因\n"
            } else {

                if (document.getElementById(reason).value == 'others') {
                    if (document.getElementById(memo).value == '') {
                        errormsg += "請輸入原因\n"
                    }

                }
            }
        }
    }

    if ( errormsg != "" ) {
        alert(errormsg);
        return false;
    }
}
</script>
<link href="css.css" rel="stylesheet" type="text/css">
<p>解傳真 CASE 的 RX NO. </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
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
            <input name="rxno" type="text" id="rxno" size="100" value="<?=$_GET['rxno'];?>" maxlength="200">    (若有多個 RX#  請用 , 隔開 ) 
            <input type="submit" name="Submit2" value="送出" />
            </div></td>        
        </tr>
    </table>
  </div>
</form>
<? if (is_null($_GET['Submit2'])) die ; ?>
<form action="<?=$PHP_SELF;?>" method="post" name="form1" onsubmit="return validate_form()">
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
  <tr>
    <th width="16">&nbsp;</th>
    <th>客戶</th> 
    <th>RX #</th> 
    <th>日期</th>     
    <th>工單號碼</th>    
    <th>品代</th>  
    <th>品名</th>
    <th>扣留日期</th> 
    <th>扣留時間</th>
    <th>解扣日期</th> 
    <th>解扣原因</th>   
    <th>其他原因</th>         
    <th width="16"><input type="checkbox" value="checkbox" name='checkall' onclick='checkedAll();'></th>
  </tr>
  <?  
      if ($_GET['occ01']!='') {
          $occfilter=" and occ01='" . $_GET['occ01'] . "'";
      } else {
          $occfilter='';
      }
      
      $rxarray=explode(',', $_GET['rxno']); 
      $max=count($rxarray);
      $rxno='';
      for($i=0; $i<$max; $i++){
          $rxno .= " or tc_ohf002 like '" . trim($rxarray[$i]) . "%'  ";                   
      } 
      $today=Date('Y-m-d');                                        
      $ssfb="select occ01, occ02, sfbud02, to_char(sfb81,'mm-dd') sfb81, sfb01, sfb05, ima02, to_char(tc_ohf004,'mm-dd') tc_ohf004, tc_ohf005  from tc_ohf_file, sfb_file, oea_file, occ_file, ima_file ".
            "where tc_ohf008 is null and ( 1=2 $rxno ) and tc_ohf001=sfb01 and sfb22=oea01 and oea04=occ01 and sfb05=ima01 " . $occfilter . " order by occ01,sfbud02,sfb81" ; 
      $erp_sqlsfb = oci_parse($erp_conn1,$ssfb );
      oci_execute($erp_sqlsfb); 
      while ($rowsfb = oci_fetch_array($erp_sqlsfb, OCI_ASSOC)) {    
      ?>    
	    <tr bgcolor="#FFFFFF"> 
		      <td><img src="i/arrow.gif" width="16" height="16">  
              <input type="hidden" name="sfbarray[]" value="<?=$rowsfb['SFB01'];?>">  
              <input type="hidden" name="rxno<?=$rowsfb['SFB01'];?>" value="<?=$rowsfb['SFBUD02'];?>"> 
          </td>   
			    <td><?=$rowsfb["OCC01"] . "--" . $rowsfb["OCC02"];?></td> 
          <td><?=$rowsfb["SFBUD02"];?></td> 
          <td><?=$rowsfb["SFB81"];?></td> 
          <td><?=$rowsfb["SFB01"];?></td>  
          <td><?=$rowsfb["SFB05"];?></td>  
          <td><?=$rowsfb["IMA02"];?></td>  
          <td><?=$rowsfb["TC_OHF004"];?></td>  
          <td><?=$rowsfb["TC_OHF005"];?></td> 
          <td><input name="faxdate<?=$rowsfb['SFB01'];?>" type="text" id="faxdate<?=$rowsfb['SFB01'];?>" size="12" maxlength="12" value=<?=$today;?> onfocus="new WdatePicker()"></td>  
          <td>
            <select name="reason<?=$rowsfb['SFB01'];?>" id="reason<?=$rowsfb['SFB01'];?>" size="1">
              <option value="">解扣原因</option>  
              <option value="空間不夠,修對咬">空間不夠,修對咬</option>            
              <option value="空間不夠,修支臺齒">空間不夠,修支臺齒</option> 
              <option value="儘量做">儘量做</option>   
              <option value="模子有损">模子有损</option>   
              <option value="支臺齒不平行,修支臺齒（或鄰牙）">支臺齒不平行,修支臺齒（或鄰牙）</option>   
              <option value="補倒凹">補倒凹</option> 
              <option value="按模上設計做">按模上設計做</option> 
              <option value="蜡型OK，继续制作">蜡型OK，继续制作</option>   
              <option value="补新Tray">补新Tray</option>    
              <option value="补咬蜡">补咬蜡</option>    
              <option value="补种植体配件">补种植体配件</option>    
              <option value="回复产品">回复产品</option>    
              <option value="回复比色">回复比色</option>                        
              <option value="others">其他</option>
            </select>                                 
          </td> 
          <td><input name="rmemo<?=$rowsfb['SFB01'];?>" id="rmemo<?=$rowsfb['SFB01'];?>"  type="text" size="20"></td>
          <td width=16><input name="risok<?=$rowsfb['SFB01'];?>" type="checkbox" id="risok" value="Y"> </td>  
      </tr> 
      <?  
      }   
  ?>    
  <tr>
    <td>&nbsp;</td>
    <td><input type="hidden" name="rxno"  value="<?=$_GET['rxno'];?>">   
        <input type="hidden" name="action" value="save">
        <input type="submit" name="Submit" value="解扣留">        
    </td>
  </tr>
</table>  
</form>
