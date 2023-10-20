<?
  session_start();
  $pagtitle = "業務部 &raquo; 傳真扣留作業v2"; 
  include("_data.php");
  include("_erp.php");
  //auth("erp_faxinv2.php");   
  date_default_timezone_set('Asia/Taipei');

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
              $indate=$_POST["faxdate".$sfb01];
              $intime=date('H:i:s');
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
          }      
      } 
      oci_commit($er_conn1);
      // 重新掃描如果使用者漏扣 要把相同訂單的工單也一起扣
      foreach ($_POST["sfb01array"] as $sfb1){        
          
          if ($_POST["risok" . $sfb1] == "Y") {
              //vd110               
              $s1="select sfb01 from sfb_file where sfb22=(select sfb22 from sfb_file where sfb01 ='$sfb1')" ; 
              $erp_sql1=oci_parse($erp_conn1,$s1);
              oci_execute($erp_sql1); 
              while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                  $sfb01 = $row1['SFB01'];

                  $s2="select count(*) qty from tc_ohf_file where tc_ohf001='$sfb01' " .
                      "and tc_ohf004 is not null and tc_ohf008 is null " .
                      "and tc_ohf014='4' and tc_ohf015='9998' " ;  //看看有沒有客服已經扣的
                  $erp_sql2=oci_parse($erp_conn1,$s2);
                  oci_execute($erp_sql2); 
                  $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);
                  if ($row2['QTY']==0) { //沒找到扣留的記錄 補一筆扣留記錄
                      $ssfb="select sfbud02, sfb05, sfb22, oea04  from sfb_file, oea_file where sfb01='$sfb01' and sfb22=oea01 " ; 
                      $erp_sqlsfb=oci_parse($erp_conn1,$ssfb);
                      oci_execute($erp_sqlsfb); 
                      $rowsfb = oci_fetch_array($erp_sqlsfb, OCI_ASSOC);
                      $sfbud02=$rowsfb["SFBUD02"];
                      $sfb05=$rowsfb["SFB05"];    
                      $sfb22=$rowsfb['SFB22'];
                      $oea04=$rowsfb['OEA04'];
                      $indate=$_POST["faxdate".$sfb01];
                      $intime=date('H:i:s');
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
                  }
              }
          }      
      }  
         
      oci_commit($erp_conn1);
      if ($msg=='') {  
          msg ("共 $rec 組 cases 扣留完畢"); 
      } else {
          msg($msg); 
      } 
      forward("erp_faxinv2.php");
  }
  
  include("_header.php");
?>
<script language='JavaScript'>
checked = false;
function checkedAll () {
  checked = checked == false;
  
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
            </select>
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
    <th>客戶</th>
    <th>RX #</th> 
    <th>日期</th>    
    <th>訂單號碼</th>
    <th>工單號碼</th>  
    <th>品名</th>
    <th>目前狀態</th>
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
          $occ = $rowsfb["OCC01"] . ' ' . $rowsfb["OCC02"];  
          $ima = $rowsfb["SFB05"] . ' ' . $rowsfb["IMA02"];
          $sfb01 = $rowsfb['SFB01']; 
          $sfb22 = $rowsfb['SFB22'];
          $casestatus='';
          $cancheckin = true;
          #檢查有無掃描秤重
          $s1 = "SELECT count(*) qty FROM tc_ogb_file " .
                "WHERE tc_ogb002 IN (SELECT sfb01 FROM sfb_file WHERE sfb22 = '$sfb22')";
          $erp_sql1 = oci_parse($erp_conn1,$s1 );
          oci_execute($erp_sql1); 
          $row1 = oci_fetch_array($erp_sql1, OCI_ASSOC);
          if ($row1['QTY']>0) {
              $casestatus = "該訂單已有掃描秤重";
              $cancheckin = false;
          }   

          #檢查相同RX有無進站未出的
          $s2 = "SELECT ecd02 FROM tc_srg_file, ecd_file " .
                " WHERE tc_srg001 IN (SELECT sfb01 FROM sfb_file WHERE sfb22 = '$sfb22') " .
                " AND tc_srg007 IS NOT NULL " .
                " AND tc_srg010 IS NULL " .
                " AND tc_srg030 = ecd01 " .
                " ORDER BY tc_srg007 DESC, tc_srg008 DESC ";
          $erp_sql2 = oci_parse($erp_conn1,$s2 );
          oci_execute($erp_sql2); 
          $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);
          if (!is_null($row2['ECD02'])) {
              $casestatus = "該訂單在 " . $row2['ECD02'] . "尚未出站";
              $cancheckin = false;
          }    

          #檢查相同訂單號有無進站未出的
          $s3 = "SELECT ecd02 FROM tc_ohf_file, ecd_file " .
                " WHERE tc_ohf012 = '$sfb22' " .
                " AND tc_ohf004 IS NOT NULL " .
                " AND tc_ohf008 IS NULL " .
                " AND tc_ohf015 = ecd01 " .
                " ORDER BY tc_ohf004 DESC, tc_ohf005 DESC ";
          $erp_sql3 = oci_parse($erp_conn1,$s3 );
          oci_execute($erp_sql3); 
          $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);
          if (!is_null($row3['ECD02'])) {
              $casestatus = "該訂單在 " . $row3['ECD02'] . "尚未出站";
              $cancheckin = false;
          }      
          //正常的CASE才 目前的工序狀態 
          IF ($cancheckin) {
              $s4 = "select ecd02 from 
                          (select ecd02, indate, intime from 
                            (select ecd02, tc_srg007 indate, tc_srg008 intime from tc_srg_file, ecd_file
                             where tc_srg001='$sfb01' and tc_srg030=ecd01 and tc_srg007 is not null)  
                            union all 
                            (select ecd02, tc_ohf004 indate, tc_ohf005 intime from tc_ohf_file, ecd_file
                             where tc_ohf001='$sfb01' and tc_ohf015=ecd01 )  
                          )
                      order by indate desc, intime desc ";
              $erp_sql4 = oci_parse($erp_conn1, $s4);
              oci_execute($erp_sql4);
              $row4 = oci_fetch_array($erp_sql4, OCI_ASSOC); //取出第一筆
              if (is_null($row4["ECD02"])) {
                  $casestatus = "尚無報工資料";
              } else {
                  $casestatus = $row4["ECD02"] . " 已出站";   
              }               
          }
          
      ?>    
	    <tr bgcolor="#FFFFFF"> 
		      <td><img src="i/arrow.gif" width="16" height="16">  
              <input type="hidden" name="sfb01array[]" value="<?=$sfb01;?>">
              <input type="hidden" name="rxno<?=$sfb01;?>" value="<?=$rowsfb['SFBUD02'];?>">
          </td>   
			    <td><?=$occ ;?></td>
          <td><?=$rowsfb["SFBUD02"];?></td> 
          <td><?=$rowsfb["SFB81"];?></td>   
          <td><?=$rowsfb["SFB22"];?></td> 
          <td><?=$sfb01;?></td>  
          <td><?=$ima;?></td>   
          <td><?=$casestatus;?></td>
          <? if ($cancheckin) { ;?>
                <td><input name="faxdate<?=$sfb01;?>" type="text" id="faxdate<?=$sfb01;?>" size="12" maxlength="12" value=<?=$today;?> onfocus="new WdatePicker()"></td>
                <td>
                    <select name="reason<?=$sfb01;?>" id="reason<?=$sfb01;?>" size="1">
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
                <td><input name="rmemo<?=$sfb01;?>" type="text" size="20"></td>
                <td width=16><input name="risok<?=$sfb01;?>" type="checkbox" id="risok" value="Y"> </td>
          <? } else { ?>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
          <? } ?>
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
