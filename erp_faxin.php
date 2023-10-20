<?
  session_start();
  $pagtitle = "業務部 &raquo; 傳真扣留作業"; 
  include("_data.php");
  include("_erp.php");
  //auth("erp_faxin.php");   
  date_default_timezone_set('Asia/Taipei');
  
  Function checkOutCase($sfb01){    
      $today=Date('Y-m-d');
      $time =Date('H:m:s');    
      
      global $erp_conn1;  
          
      //檢查有無 有在製量, 有刷進未刷出的工單
      $ssfb10="SELECT * FROM tc_srg_file WHERE tc_srg001= '$sfb01' AND tc_srg005 is NULL and tc_srg007 is not null and tc_srg010 is null";
      $erp_sqlsfb10 = oci_parse($erp_conn1,$ssfb10 );
      oci_execute($erp_sqlsfb10); 
      $rowsfb10 = oci_fetch_array($erp_sqlsfb10, OCI_ASSOC);
      $tc_srg004=$rowsfb10['TC_SRG004'];   
      $tc_srg005=$rowsrg10['TC_SRG005'];  
      $tc_srg006=$rowsrg10['TC_SRG006']; 
      $tc_srg007=$rowsrg10['TC_SRG007'];  
      $tc_srg018=$rowsrg10['TC_SRG018'];                        
      $tc_srg019=$rowsrg10['TC_SRG019'];  
      $tc_srg022=$rowsrg10['TC_SRG022'];  
      if (!is_null($tc_srg004)) {   //有在製量
                                           
         if ($tc_srg006=='N') {  #不要QC
             $ssrg11="UPDATE tc_srg_file SET tc_srg005=NULL, tc_srg010=to_date('$today','yy/mm/dd'), tc_srg011='$time', tc_srg012 = '11015996' " .  
                     "WHERE tc_srg001='$sfb01'  AND tc_srg004='$tc_srg004' ";
             $erp_sqlsrg11=oci_parse($erp_conn1,$ssrg11);
             oci_execute($erp_sqlsrg11); 
             
             $ssrg12="UPDATE tc_srg_file SET tc_srg005=$tc_srg005 " .  
                     "WHERE tc_srg001='$sfb01'  AND tc_srg004= " .
                     "(SELECT MIN(tc_srg004) FROM tc_srg_file WHERE tc_srg001='$sfb01' AND tc_srg004 > '$tc_srg004')";
             $erp_sqlsrg12=oci_parse($erp_conn1,$ssrg12);
             oci_execute($erp_sqlsrg12);    
         } else {
             $ssrg13="UPDATE tc_srg_file SET tc_srg010=to_date('$today','yy/mm/dd'), tc_srg011='$time', tc_srg012 = '11015996' " .  
                     "WHERE tc_srg001='$sfb01'  AND tc_srg004='$tc_srg004' ";
             $erp_sqlsrg13=oci_parse($erp_conn1,$ssrg13);
             oci_execute($erp_sqlsrg13); 
         }      
      }  
  } 
  

  if ($_POST["action"] == "save") { 
      $rec=0;  
      $oldsfb22='';  
      $account=$_SESSION['account'];
      foreach ($_POST["sfb01array"] as $sfb01){          
          
          if ($_POST["risok" . $sfb01] == "Y") {
              $rxno=$_POST["rxno".$sfb01];
              //$faxdate=$_POST["faxdate".$sfb01]; 
              //先將要修改的資料記錄下來                                    
              $queryap   = "insert into erp_faxin_updaterecord ( rxno, sfb01, faxtype, username, ip ) values ( 
                          '" . $rxno                              . "',    
                          '" . $sfb01                             . "', 'in',
                          '" . safetext($account)                 . "',       
                          '" . safetext($_SERVER['REMOTE_ADDR'])  . "')";
              $resultap= mysql_query($queryap) or die ('23 erp_faxin_updaterecord added error. ' .mysql_error());
                
              $msg='';
              
              //vd110               
              $ssfb="select sfbud02, sfb05, sfb22, oea04  from sfb_file, oea_file where sfb01='$sfb01' and sfb22=oea01 " ; 
              $erp_sqlsfb=oci_parse($erp_conn1,$ssfb);
              oci_execute($erp_sqlsfb); 
              $rowsfb = oci_fetch_array($erp_sqlsfb, OCI_ASSOC);
              $sfbud02=$rowsfb["SFBUD02"];
              $sfb05=$rowsfb["SFB05"];    
              $sfb22=$rowsfb['SFB22'];
              $oea04=$rowsfb['OEA04'];  
              //20130413 可修改扣留日期
              $indate=$_POST["faxdate".$sfb01];
              $intime=date('G:i:s');
              $memo=$_POST["reason".$sfb01].$_POST["rmemo".$sfb01];
              $stcohf= "insert into tc_ohf_file(tc_ohf001, tc_ohf002, tc_ohf003, tc_ohf004, tc_ohf005, tc_ohf006, tc_ohf007, tc_ohf012, tc_ohf013, tc_ohf014, tc_ohf015) " .
                       "values ('$sfb01','$sfbud02','$sfb05',to_date('$indate','yy/mm/dd'), '$intime','$account','$memo','$sfb22', '$oea04', '4', '9998')";  
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
              
              //有做扣留 要幫CASE出站
              //$today=Date('Y-m-d');
              //$time =Date('H:m:s');    
                                               
              //檢查 有在製量且有刷進未刷出的工單  幫它做刷出動作
             // $ssfb10="SELECT tc_srg004, tc_srg005, tc_srg006 FROM tc_srg_file WHERE tc_srg001= '$sfb01' AND tc_srg005 is not NULL and tc_srg007 is not null and tc_srg010 is null";
             // $erp_sqlsfb10 = oci_parse($erp_conn1,$ssfb10 );
             // oci_execute($erp_sqlsfb10); 
             // $rowsfb10 = oci_fetch_array($erp_sqlsfb10, OCI_ASSOC);
             // $tc_srg004=$rowsfb10['TC_SRG004'];
             // $tc_srg005=$rowsfb10['TC_SRG005'];  
              //$tc_srg006=$rowsfb10['TC_SRG006'];
             // if (!is_null($tc_srg004)) {   //有在製量   
             //    if ($tc_srg006=='N') {  #不要QC
             //        $ssrg11="UPDATE tc_srg_file SET tc_srg005=NULL, tc_srg010=to_date('$today','yy/mm/dd'), tc_srg011='$time', tc_srg012 = '11015996' " .  
                  //           "WHERE tc_srg001='$sfb01'  AND tc_srg004='$tc_srg004' ";
                  //   $erp_sqlsrg11=oci_parse($erp_conn1,$ssrg11);
                  //   oci_execute($erp_sqlsrg11); 
                  //   
                  //   $ssrg12="UPDATE tc_srg_file SET tc_srg005=$tc_srg005 " .  
                 //            "WHERE tc_srg001='$sfb01'  AND tc_srg004= " .
                 //            "(SELECT MIN(tc_srg004) FROM tc_srg_file WHERE tc_srg001='$sfb01' AND tc_srg004 > '$tc_srg004')";
                //     $erp_sqlsrg12=oci_parse($erp_conn1,$ssrg12);
               //      oci_execute($erp_sqlsrg12);    
               //  } else {
              //       $ssrg13="UPDATE tc_srg_file SET tc_srg010=to_date('$today','yy/mm/dd'), tc_srg011='$time', tc_srg012 = '11015996' " .  
             //                "WHERE tc_srg001='$sfb01'  AND tc_srg004='$tc_srg004' ";
             //        $erp_sqlsrg13=oci_parse($erp_conn1,$ssrg13);
             //        oci_execute($erp_sqlsrg13); 
             //    }      
             // }    
          }      
      } 
      oci_commit($er_conn1);
      if ($msg=='') {  
          msg ("共 $rec 組 cases 扣留完畢"); 
      } else {
          msg($msg); 
      } 
      forward("erp_faxin.php");                                                              
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
<p>扣留CASE的RX NO. </p>
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
            <input name="rxno" type="text" id="rxno" size="100" value="<?=$_GET['rxno'];?>" maxlength="200">    (若有多個 RX#  請用 , 隔開 ) 
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
    <th>日期</th>    
    <th>訂單號碼</th>
    <th>工單號碼</th>    
    <th>品代</th>  
    <th>品名</th>
    <th>扣留日期</th>
    <th>扣留原因</th> 
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
          $rxno .= "'" . trim($rxarray[$i]) . "',";                   
      } 
       
      $today=Date('Y-m-d');                                       
      $ssfb="select occ01, occ02, sfbud02, to_char(sfb81,'mm-dd') sfb81, sfb22, sfb01, sfb05, ima02 from sfb_file, oea_file, occ_file, ima_file ".
            "where sfbud02 in ($rxno'') and sfb22=oea01 and sfb09<sfb08 and oea04=occ01 and sfb05=ima01 " . $occfilter . " order by occ01,sfbud02,sfb81" ; 
      $erp_sqlsfb = oci_parse($erp_conn1,$ssfb );
      oci_execute($erp_sqlsfb); 
      while ($rowsfb = oci_fetch_array($erp_sqlsfb, OCI_ASSOC)) {    
      ?>    
	    <tr bgcolor="#FFFFFF"> 
		      <td><img src="i/arrow.gif" width="16" height="16">  
              <input type="hidden" name="sfb01array[]" value="<?=$rowsfb['SFB01'];?>"> 
              <input type="hidden" name="rxno<?=$rowsfb['SFB01'];?>" value="<?=$rowsfb['SFBUD02'];?>"> 
          </td>   
			    <td><?=$rowsfb["OCC01"];?></td>
          <td><?=$rowsfb["OCC02"];?></td>  
          <td><?=$rowsfb["SFBUD02"];?></td> 
          <td><?=$rowsfb["SFB81"];?></td>   
          <td><?=$rowsfb["SFB22"];?></td> 
          <td><?=$rowsfb["SFB01"];?></td>  
          <td><?=$rowsfb["SFB05"];?></td>  
          <td><?=$rowsfb["IMA02"];?></td>                                                                                                       
          <td><input name="faxdate<?=$rowsfb['SFB01'];?>" type="text" id="faxdate<?=$rowsfb['SFB01'];?>" size="12" maxlength="12" value=<?=$today;?> onfocus="new WdatePicker()"></td>
          <td>
            <select name="reason<?=$rowsfb['SFB01'];?>" id="reason<?=$rowsfb['SFB01'];?>" size="1">
              <option value="">扣留原因</option>  
              <option value="空間不夠">空間不夠</option>            
              <option value="Margin不清">Margin不清</option> 
              <option value="確認咬合">確認咬合</option>   
              <option value="模子有损">模子有损</option>   
              <option value="支臺齒不平行；有倒凹">支臺齒不平行；有倒凹</option>   
              <option value="确认產品">确认產品</option> 
              <option value="確認比色">確認比色</option> 
              <option value="確認翻譯">確認翻譯</option>   
              <option value="確認齒位">確認齒位</option>    
              <option value="印模不好">印模不好</option>    
              <option value="索配件">索配件</option>    
              <option value="確認設計">確認設計</option>    
              <option value="确认蜡型">确认蜡型</option>    
              <option value="確認Post 分體/連體">確認Post 分體/連體</option>    
              <option value="">其他</option>         
            </select>                                 
          </td>
          <td><input name="rmemo<?=$rowsfb['SFB01'];?>" type="text" size="20"></td> 
          <td width=16><input name="risok<?=$rowsfb['SFB01'];?>" type="checkbox" id="risok" value="Y"> </td>  
      </tr> 
      <?  
      }   
  ?>    
  <tr>
    <td>&nbsp;</td>
    <td><input type="hidden" name="rxno"  value="<?=$_GET['rxno'];?>">   
        <input type="hidden" name="action" value="save">
        <input type="submit" name="Submit" value="扣留">        
    </td>
  </tr>
</table>  
</form>
