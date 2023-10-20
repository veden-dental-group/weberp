<?
  session_start();
  $pagtitle = "業務部 &raquo; 扣留/解傳真查詢作業"; 
  include("_data.php");
  include("_erp.php");
  //auth("erp_faxout.php");   
  date_default_timezone_set('Asia/Taipei');
  if (is_null($_GET['bdate'])) {
      $bdate=date('Y-m-d')  ;
  } else {
      $bdate=$_GET['bdate'] ;
  }
  if (is_null($_GET['istype'])) {
      $istype=1 ;
  } else {             
      if ($_GET['istype']==1) {
          $istype=1;
      } else {
          $istype=2;
      }   
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
<p>扣留/解傳真 CASE 的 RX NO. </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">     
            客戶: 
            <select name="occ01" id="occ01">   
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
            </select>   &nbsp;&nbsp;   
            日期:   
            <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> &nbsp;&nbsp;    
            扣留/解扣:
            <input type="radio" name="istype" id="istype1" value='1' <? if ($istype==1) echo " checked";?> >扣留</input>
            <input type="radio" name="istype" id="istype2" value='2' <? if ($istype==2) echo " checked";?> >解扣</input>  &nbsp;&nbsp;    
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
    <th>訂單號碼</th>   
    <th>工單號碼</th>    
    <th>品代</th>  
    <th>品名</th>
    <th>扣留日期</th> 
    <th>扣留時間</th> 
    <th>扣留人員</th> 
    <th>扣留原因</th>  
    <th>解扣日期</th> 
    <th>解扣原因</th>  
    <th>解扣人員</th>  
    <th>其他原因</th>                                                                                      
  </tr>
  <?  
      $occfilter=" and tc_ohf013='" . $_GET['occ01'] . "'";    
      
      if ($_GET['istype']=='1') {
          $datefilter=" and tc_ohf004=to_date('$bdate','yy/mm/dd') ";
      } else {
          $datefilter=" and tc_ohf008=to_date('$bdate','yy/mm/dd') ";  
      }
      
                                           
      $ssfb="select sfbud02, sfb22, sfb01, sfb05, ima02, to_char(tc_ohf004,'mm-dd') tc_ohf004, tc_ohf005, tc_ohf006, tc_ohf007, " .
            "to_char(tc_ohf008,'mm-dd') tc_ohf008, tc_ohf009, tc_ohf010, tc_ohf011 from tc_ohf_file, sfb_file, oea_file, ima_file ".
            "where tc_ohf001=sfb01 and sfb22=oea01 and sfb05=ima01 $occfilter $datefilter order by sfbud02,sfb22,sfb01" ; 
      $erp_sqlsfb = oci_parse($erp_conn1,$ssfb );
      oci_execute($erp_sqlsfb); 
      $rr=0;
      $oldsfbud02='';
      while ($rowsfb = oci_fetch_array($erp_sqlsfb, OCI_ASSOC)) {    
        $g="select gen02 from gen_file where gen01='" . $rowsfb['TC_OHF006'] . "' ";
        $erp_sqlg = oci_parse($erp_conn1,$g );
        oci_execute($erp_sqlg); 
        $rowg = oci_fetch_array($erp_sqlg, OCI_ASSOC);
        $gen021=$rowg['GEN02'];
        
        $g="select gen02 from gen_file where gen01='" . $rowsfb['TC_OHF010'] . "' ";
        $erp_sqlg = oci_parse($erp_conn1,$g );
        oci_execute($erp_sqlg); 
        $rowg = oci_fetch_array($erp_sqlg, OCI_ASSOC);
        $gen022=$rowg['GEN02'];
        
        if ($oldsfbud02!=$rowsfb['SFBUD02']){
            $rr++;
            $oldsfbud02=$rowsfb['SFBUD02'];
        }
        
      ?>    
	    <tr bgcolor="#FFFFFF"> 
		      <td><?=$rr;?></td>  
          <td><?=$rowsfb["SFBUD02"];?></td> 
          <td><?=$rowsfb["SFB01"];?></td>  
          <td><?=$rowsfb["SFB22"];?></td> 
          <td><?=$rowsfb["SFB05"];?></td>  
          <td><?=$rowsfb["IMA02"];?></td>  
          <td><?=$rowsfb["TC_OHF004"];?></td>  
          <td><?=$rowsfb["TC_OHF005"];?></td> 
          <td><?=$rowsfb["TC_OHF006"]. ' ' . $gen021;?></td>  
          <td><?=$rowsfb["TC_OHF007"];?></td>   
          <td><?=$rowsfb["TC_OHF008"];?></td>  
          <td><?=$rowsfb["TC_OHF009"];?></td> 
          <td><?=$rowsfb["TC_OHF010"]. ' ' . $gen022;?></td>  
          <td><?=$rowsfb["TC_OHF011"];?></td>        
      </tr> 
      <?  
      }         
  ?>     
      <tr>
        <td colspan="14">共 <?=$rr;?> 組case(s)</td> 
      </tr>   
</table>  
</form>
