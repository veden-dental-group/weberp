<?
//匯出期間內某客戶出貨CASE的CE LOT NO

session_start();   
include("_data.php");  
$pagetitle = "業務部 &raquo; Email Traceability";
//auth("erp_traceability_email.php");  
  
date_default_timezone_set('Asia/Taipei'); 

if (is_null($_GET['bdate'])) {
  $bdate = date('Y-m-d');
} else {
  $bdate=$_GET['bdate'];
} 
if (is_null($_GET['edate'])) {
  $edate = date('Y-m-d');
} else {
  $edate=$_GET['edate'];
} 

if (is_null($_GET['email'])) {
  $email = 'frank@vedenlabs.com';
} else {
  $email=$_GET['email'];
}

$daterange=$bdate.'--'.$edate;
$occ01=$_GET['occ01']; 

if ($_GET["submit"]=="E-Mail") { 
   
    //檢查有無以客戶名稱命名的目錄  沒有的話要建一個 一個客戶一個目錄放檔案用
    $dirname='traceability/'.$occ01 ."/";                
    if (!(file_exists($dirname))) {       
        mkdir($dirname);                                            
    }
    
    $template_file=  "templates/traceability-" . $occ01 . ".xls";
    if (!(file_exists($template_file))) { 
        msg( $occ01 . '--無template file  無法匯出, 請洽IT!!');      
        forward("erp_traceability_email.php?submit=Y&bdate=".$bdate. "&edate=" .$edate ."&occ01=".$occ01);                                         
    }
                                               
    $email=$_GET['email'];
    error_reporting(E_ALL);  
    require_once 'classes/PHPExcel.php'; 
    require_once 'classes/PHPExcel/IOFactory.php';  
    //$objPHPExcel = new PHPExcel();
    $objReader = PHPExcel_IOFactory::createReader('Excel5');  
    
     
       
    $query = "select * from traceability where flag1='N' and occ01='$occ01' and daterange='$daterange' order by rxno ";
    $result= mysql_query($query) or die ('57 TraceAbility access error!!' . mysql_error());   
    $j=0; 
    $oldrxno='';
    while ($row   = mysql_fetch_array($result) ){   
        $rxno=$row['rxno'];  
        $j++;
        if ($oldrxno==$rxno){
            $k++;
        } else {
            $k=1;  
            $oldrxno=$rxno;
        }
        
        $objPHPExcel = $objReader->load($template_file);  
            // Set properties
        $objPHPExcel ->getProperties()->setCreator("Frank")
                     ->setLastModifiedBy("Frank")
                     ->setTitle("Traceability")
                     ->setSubject("Traceability")
                     ->setDescription("Traceability")
                     ->setKeywords("Traceability")
                     ->setCategory("Traceability");
        //第一份             
        $objPHPExcel->setActiveSheetIndex(0)     
                    ->setCellValue('E1', $row['rxno'])
                    ->setCellValue('E2', $row['dentist'])   
                    ->setCellValue('E3', $row['patient'])   
                    ->setCellValue('C4', $row['pename']) 
                    ->setCellValue('C5', $row['tooth'])   
                    ->setCellValue('C6', $row['shade']);    
        //第二份            
        $objPHPExcel->setActiveSheetIndex(0)     
                    ->setCellValue('E21', $row['rxno'])
                    ->setCellValue('E22', $row['dentist'])   
                    ->setCellValue('E23', $row['patient'])   
                    ->setCellValue('C24', $row['pename']) 
                    ->setCellValue('C25', $row['tooth'])   
                    ->setCellValue('C26', $row['shade']);               
      
        $guid=$row['guid'];
        $queryb="select * from traceabilitybody where traceabilityguid='$guid' order by sort  ";
        $resultb= mysql_query($queryb) or die ('57 TraceAbilityBody access error!!' . mysql_error());     
        $y=9;
        $z=29;
        while ($rowb = mysql_fetch_array($resultb) ){    
            $objPHPExcel->setActiveSheetIndex(0) 
                ->setCellValue('B'.$y ,$rowb['pename'])
                ->setCellValue('D'.$y ,$rowb['manufacturer'])  
                ->setCellValue('E'.$y ,$rowb['ceno'])  
                ->setCellValue('F'.$y ,$rowb['lotno']) ; 
                
            $objPHPExcel->setActiveSheetIndex(0) 
                ->setCellValue('B'.$z ,$rowb['pename'])
                ->setCellValue('D'.$z ,$rowb['manufacturer'])  
                ->setCellValue('E'.$z ,$rowb['ceno'])  
                ->setCellValue('F'.$z ,$rowb['lotno']) ;     
                
            $y++;
            $z++;  
        }  
                
        $objPHPExcel->getActiveSheet()->setTitle($rxno);
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');   
        if ($k==1){
            $filename[$j]=$dirname . $rxno .".xls";   
        } else {
            $filename[$j]=$dirname . $rxno .'-'. $k.".xls";   
        }
        
        $objWriter->save($filename[$j]);    
    }
   
    ////email
    require_once("_classes.php");
    $_SERVER['SERVER_NAME'] = 'www.vedenlabs.com';
    $regards = "Veden Dental Labs Inc.";
    
    $m = new mailer();
    $m->setMessage($regards);
    $m->setPriority( 'High' );
    $m->setFrom( "frankyu@vedenlabs.com", "Veden Dental Labs Inc"  );
    $m->setReplyTo( "frankyu@vedenlabs.com", "Veden Dental Labs Inc" );  
    
    //每100個附件送一個email
    $totalpage=floor($j/100)+1;
    $page=1;
    for ($x=1; $x<=$j; $x++) {          
        $m->attachFile( $filename[$x], $filename[$x], "application/vnd.ms-excel"); 
        if ($x == $page*100) {
            $m->send($email, $occ01 . "  " . $bdate . " Traceability ( $page/$totalpage )  ") ;  
            $m = new mailer();
            $m->setMessage($regards);
            $m->setPriority( 'High' );
            $m->setFrom( "frankyu@vedenlabs.com", "Veden Dental Labs Inc"  );
            $m->setReplyTo( "frankyu@vedenlabs.com", "Veden Dental Labs Inc" );   
            $page++; 
        }
    }      
    $m->send($email, $occ01 . "  " . $bdate . " Traceability ( $page/$totalpage )  ") ;    
    msg('寄送完畢!!');     
}

if ($_GET["submit"]=="計算") {   //重算期間內 某客戶的全部資料 

    //要先刪出貨期間內的資料
    $query="update traceability set flag1='D' where occ01='$occ01' and daterange='$daterange' and oga02 >='$bdate' and oga02 <= '$edate' ";
    $result=mysql_query($query) or die ('189 TraceAbility Updated error!!;' . mysql_error());   
      

    $s1="select ogb31, ogb32, ogaud02, to_char(oga02, 'yyyy-mm-dd') oga02, occ01, occ02, sfb01, sfb82, ta_oea002, ta_oea003, ta_oea046, ta_oea047, ta_oea048, ima01, ima02, ima1002, oebud03 ".
        "from oga_file, occ_file, ogb_file, ( select distinct tc_fro002 from tc_fro_file where tc_fro001='$occ01' or tc_fro001='MISC' ), sfb_file, oea_file, oeb_file, ima_file " . //取出本客戶或MICS客戶有設定的產品清單
        "where oga02 between to_date('$bdate','yy/mm/dd') and to_date('$edate','yy/mm/dd') and oga04='$occ01' and oga04=occ01 " .
        "and oga01=ogb01 and ogb04=tc_fro002 and ogb31=sfb22 and ogb32=sfb221 " .
        "and ogb31=oea01 and ogb31=oeb01 and ogb32=oeb03 and ogb04=ima01 ".
      //  "and oga16='B531-1304260827' " .
        "order by oga02, ogaud02";
    $erp_sql1 = oci_parse($erp_conn1,$s1 );
    oci_execute($erp_sql1);                                 
    while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {          //取出單頭所有資料
        $occ02        =$row1['OCC02'] ;
        $oga02        =$row1['OGA02'];
        $rxno         =$row1['OGAUD02'];
        $orderno      =$row1['OGB31'];
        $ordernosn    =$row1['OGB32'];  
        $workticketno =$row1['SFB01'];
        $dentist      =$row1['TA_OEA002'];
        $patient      =$row1['TA_OEA003'];
        $pcode        =$row1['IMA01'];
        $pename       =$row1['IMA1002'];
        $pcname       =$row1['IMA02']; 
        if (is_null($row1['TA_OEA047'])) {
          if (is_null($row1['TA_OEA046'])) { 
            if (is_null($row1['TA_OEA048'])) { 
              $shade =$row1['TA_OEA048']; 
            } else {
              $shade='';  
            } 
          } else {
            $shade    =$row1['TA_OEA046'];   
          }                   
        } else {
          $shade      =$row1['TA_OEA047'];
        }
        $tooth        =$row1['OEBUD03'];
        $sfb82        =$row1['SFB82'];
        
        //寫入一筆單頭
        $guid=uuid();
        $queryus = "insert into traceability ( guid, occ01, occ02, daterange, oga02, workticketno, orderno, ordernosn, rxno, dentist, patient, pcode, pcname, pename, tooth, shade, flag1 ) values (
                  '" . $guid              . "',   
                  '" . $occ01             . "',  
                  '" . $occ02             . "',       
                  '" . $daterange         . "',
                  '" . $oga02             . "',
                  '" . $workticketno      . "',
                  '" . $orderno           . "',
                  '" . $ordernosn         . "',
                  '" . $rxno              . "',
                  '" . $dentist           . "',
                  '" . $patient           . "',
                  '" . $pcode             . "',
                  '" . $pcname            . "',
                  '" . $pename            . "',
                  '" . $tooth             . "',
                  '" . $shade             . "','N')";     
        $resultus = mysql_query($queryus) or die ('216 TraceAbility added error!!.' .mysql_error());
        
        
        //開始取出已設定要秀出的原物料 
        //先計算要秀幾種原料
        $s2="select distinct tc_frl010 from tc_frl_file where tc_frl001='$pcode' order by tc_frl010 " ;  
        $erp_sql2 = oci_parse($erp_conn1,$s2 );
        oci_execute($erp_sql2);       
        while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) { 
          $tc_frl010=$row2['TC_FRL010'] ;  
          $tc_frl002='MISC';
          $tc_frl003='MISC';
          $tc_frl005='MISC';
          
          //找出第一種要秀出的 製處  顏色 客戶  如果有指定到  就用指定的  如果沒指定到就取MISC
          $s3="select tc_frl002, tc_frl003, tc_frl005 from tc_frl_file where tc_frl001='$pcode' and tc_frl010='$tc_frl010' order by tc_frl002, tc_frl003, tc_frl005" ;  
          $erp_sql3 = oci_parse($erp_conn1,$s3 );
          oci_execute($erp_sql3);  
          while ($row3 = oci_fetch_array($erp_sql3, OCI_ASSOC)) {
            if ($row3['TC_FRL002']==$sfb82){ //製處有指定
                $tc_frl002=$sfb82;              
            }
            if ($row3['TC_FRL003']==$shade){ //顏色有指定
                $tc_frl003=$shade;              
            }
            if ($row3['TC_FRL005']==$occ1){ //客戶有指定
                $tc_frl005=$occ01;              
            }
          }  

          //到此 已經算出要依製處/比色/客戶或MISC來取出物料編號了                  
          $s4="select tc_frl004, ima02, ima1002 from tc_frl_file, ima_file where tc_frl001='$pcode' and tc_frl010='$tc_frl010' and tc_frl002='$tc_frl002' and tc_frl003='$tc_frl003' and tc_frl005='$tc_frl005' and tc_frl004=ima01 " ;  
          $erp_sql4 = oci_parse($erp_conn1,$s4 );
          oci_execute($erp_sql4);  
          $row4 = oci_fetch_array($erp_sql4, OCI_ASSOC);
          $tc_frl004=$row4['TC_FRL004'];
          $ima1002=$row4['IMA1002'];
          $ima02=$row4['IMA02']; 
          
          //取出物料的ceno, lotno, manufacturer
          $s5="select tc_img003, tc_img004, tc_img005 from tc_img_file where tc_img001<=to_date('$oga02','yy/mm/dd') and tc_img002='$tc_frl004' order by tc_img001 desc " ;  
          $erp_sql5 = oci_parse($erp_conn1,$s5 );
          oci_execute($erp_sql5);  
          $row5 = oci_fetch_array($erp_sql5, OCI_ASSOC);
          $tc_img003=$row5['TC_IMG003'];
          $tc_img004=$row5['TC_IMG004'];   
          $tc_img005=$row5['TC_IMG005'];   
          
          //寫入單身
          $queryus = "insert into traceabilitybody ( guid, traceabilityguid, sort, pcode, pcname, pename, manufacturer, ceno, lotno ) values (
                '" . uuid()             . "',   
                '" . $guid             . "',  
                '" . $tc_frl010         . "',       
                '" . $tc_frl004         . "',
                '" . $ima02             . "',
                '" . $ima1002           . "',
                '" . $tc_img005         . "',
                '" . $tc_img004         . "', 
                '" . $tc_img003         . "')";     
          $resultus = mysql_query($queryus) or die ('283 TraceAbilityBody added error!!.' .mysql_error());            
       }                      
    }  
    msg('計算完畢!!');    
}

if ($_POST["action"] == "save") {                     
    foreach ($_POST["tbguidarray"] as $tbguid){ 
        if ($_POST["risok" . $tbguid] == "Y") {
          $bdate        =$_POST['bdate'];
          $edate        =$_POST['edate']; 
          $occ01        =$_POST['occ01']; 
          $tguid        =$_POST['tguid'.$tbguid];      
          $dentist      =$_POST['dentist'.$tbguid];
          $patient      =$_POST['patient'.$tbguid];    
          $tooth        =$_POST['tooth'.$tbguid];    
          $shade        =$_POST['shade'.$tbguid];    
          $tbpename     =$_POST['tbpename'.$tbguid];  
          $manufacturer =$_POST['manufacturer'.$tbguid];       
          $ceno         =$_POST['ceno'.$tbguid];      
          $lotno        =$_POST['lotno'.$tbguid]; 
          $duperxno       =$_POST['duperxno'.$tbguid];   
          $dupecode       =$_POST['dupecode'.$tbguid];
          /*     
          //先將要修改的資料記錄下來                                    
          $queryap   = "insert into erp_traceability_updaterecord ( bdate, occ01, sfe02, sfe28, weight, username, ip ) values (
                      '" . safetext($bdate)                   . "',     
                      '" . safetext($occ01)                   . "',  
                      '" . safetext($sfe02)                   . "',      
                      '" . safetext($sfe28)                   . "',     
                      '" . safetext($tasfe002)                . "',       
                      '" . safetext($_SESSION['account'])     . "',       
                      '" . safetext($_SERVER['REMOTE_ADDR'])  . "')";
          $resultap= mysql_query($queryap) or die ('31 erp_addmetalweight_updaterecord added error. ' .mysql_error());               
          */
          //先更新單頭
          if ($duperxno=='N') {
              $query = "update traceability set        
                        dentist      = '" . safetext($dentist)     . "',   
                        patient      = '" . safetext($patient)     . "',
                        shade        = '" . safetext($shade)       . "'  
                        where guid   = '" . safetext($tguid)       . "' limit 1";
              $result = mysql_query($query) or die ('330 TraceAbility updated error!! ' . mysql_error()); 
          }
          
          if ($dupecode=='N') {
              $query = "update traceability set                        
                        tooth        = '" . safetext($tooth)       . "'                           
                        where guid   = '" . safetext($tguid)       . "' limit 1";
              $result = mysql_query($query) or die ('330 TraceAbility updated error!! ' . mysql_error()); 
          }
          
          //先更新單身
          $query = "update traceabilitybody set        
                    manufacturer = '" . safetext($manufacturer)   . "',   
                    ceno         = '" . safetext($ceno)           . "',
                    pename       = '" . safetext($tbpename)       . "',   
                    lotno        = '" . safetext($lotno)          . "'
                    where guid   = '" . safetext($tbguid)         . "' limit 1";
          $result = mysql_query($query) or die ('330 TraceAbility updated error!! ' . mysql_error()); 
          
        }  
    }  
    //可能會改到shade 所以要重新計算一次使用的品代/品名/manufacture/ce no /lot no
    //若shade為空 就不要重新取資料
    
    
    msg('更新完畢!!');     
    forward("erp_traceability_email.php?submit=Y&bdate=".$bdate. "&edate=" .$edate ."&occ01=".$occ01);  
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
<p>Export Traceability </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">    
            客戶:
            <select name="occ01" id="occ01">    
                <?
                  $s1= "select occ01, occ02 from occ_file order by occ01 ";
                  $erp_sql1 = oci_parse($erp_conn,$s1 );
                  oci_execute($erp_sql1);  
                  while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                      echo "<option value=" . $row1["OCC01"];  
                      if ($_GET["occ01"] == $row1["OCC01"]) echo " selected";                  
                      echo ">" . $row1['OCC01'] ."--" .$row1["OCC02"] . "</option>"; 
                  }   
                ?>
            </select>&nbsp;&nbsp; 
            出貨期間:   
            <input name="bdate" type="text" id="bdate" size="12" maxlength="12" onfocus="new WdatePicker()" value=<?=$bdate;?>> ~~ 
            <input name="edate" type="text" id="edate" size="12" maxlength="12" onfocus="new WdatePicker()" value=<?=$edate;?>> &nbsp; &nbsp;   
            <input type="submit" name="submit" id="submit" value="計算"> &nbsp; &nbsp;   
            <input type="submit" name="submit" id="submit" value="查詢"> &nbsp; &nbsp;    
            <input type="submit" name="submit" id="submit" value="E-Mail">to:<input name="email" type="text" id="email" size='40' value=<?=$email;?>> &nbsp; &nbsp;   
            </div></td>        
        </tr>
    </table>
  </div>
</form>


<? if (is_null($_GET['submit'])) die ; ?>
<form action="<?=$PHP_SELF;?>" method="post" name="form1">
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
  <tr>
    <th width="16">&nbsp;</th>
    <th>出貨日期</th>  
    <th>Case No.</th>   
    <th>訂單號碼</th>  
    <th>工單號碼</th>      
    <th>Dentist</th> 
    <th>Patient</th>       
    <th>Shade</th>  
    <th>品代</th>    
    <th>Restoration</th>   
    <th>Teeth No.</th>     
    <th>原物料代號</th>   
    <th>Alloy/Porcelian</th>   
    <th>Manufacturer</th>   
    <th>CE No.</th>   
    <th>Lot No.</th>  
    <th width="16"><input type="checkbox" value="checkbox" name='checkall' onclick='checkedAll();'></th>                                                                            
  </tr>            
  <?   
      $_GET['submit']=''; 
      $query = "select t.guid tguid, t.occ01 tocc01, t.occ02 tocc02, t.oga02 toga02, t.workticketno tworkticketno, t.orderno torderno, t.rxno trxno, t.dentist tdentist, t.patient tpatient, t.pcode tpcode, t.pename tpename, " .
               "t.tooth ttooth, t.shade tshade, tb.guid tbguid, tb.pcode tbpcode, tb.pename tbpename, tb.manufacturer tbmanufacturer, tb.ceno tbceno, tb.lotno tblotno ".
               "from traceability t, traceabilitybody tb ".
               "where t.flag1='N' and  t.guid=tb.traceabilityguid and t.occ01='$occ01' and t.daterange='$daterange'  " .
               "order by oga02,rxno " ;
      $result = mysql_query($query) or die ('546 TraceAbility error!!' . mysql_error());
      $oldrxno='';    
      $oldpcode='';
      $duperxno='N';    
      $dupecode='N';
      while ($row= mysql_fetch_array($result)) { 
        if ($oldrxno != $row['trxno']) {             
            $oldrxno        =$row['trxno'];  
            $doga02         =$row['toga02']; 
            $drxno          =$row['trxno'];
            $dorderno       =$row['torderno'];
            $dworkticketno  =$row['tworkticketno'];
            $ddentist       =$row['tdentist'];
            $dpatient       =$row['tpatient'];               
            $oldpcode       =$rowp['tpcode'];
            $dpcode         =$row['tpcode'];
            $dpename        =$row['tpename'];
            $dtooth         =$row['ttooth'];
            $dshade         =$row['tshade'];
            $duperxno       ='N';
            $dupecode       ='N';
        } else {
            $doga02         ='';
            $drxno          ='';
            $dorderno       ='';
            $dworkticketno  ='';
            $ddentist       ='';
            $dpatient       ='';
            $duperxno       ='Y';
            if ($oldpcode  != $row['tpcode'] ) {
                $dpcode     =$row['tpcode'];
                $dpename    =$row['tpename'];
                $dtooth     =$row['ttooth'];  
                $dupecode   ='N';
            } else {
                $dpcode     ='';
                $dpename     ='';
                $dtooth     ='';     
                $dupecode   ='Y';
            }
        }
        
          
        $key=$row['tbguid'];                                               
      ?>    
      <tr bgcolor="#FFFFFF"> 
          <? if ($duperxno=='N') { ?>  
             <td><img src="i/arrow.gif" width="16" height="16"></td> 
          <? } else { ?>
            <td>&nbsp;</td>   
          <? } ?>            
      
          <td><?=$doga02;?>
              <input type="hidden" name="tbguidarray[]"   value="<?=$key;?>"> 
              <input type="hidden" name="tguid<?=$key;?>" value="<?=$row['tguid'];?>">        
              <input type="hidden" name="duperxno<?=$key;?>" value="<?=$duperxno;?>">     
              <input type="hidden" name="dupecode<?=$key;?>" value="<?=$dupecode;?>">   
          </td>        
          <td><?=$drxno;?></td>
          <td><?=$dorderno;?></td>   
          <td><?=$dworkticketno;?></td> 
          
          <? if ($duperxno=='N') { ?>
            <td width=16><input name="dentist<?=$key;?>" type="text" id="dentist<?=$key;?>" value="<?=$ddentist;?>" </td> 
            <td width=16><input name="patient<?=$key;?>" type="text" id="patient<?=$key;?>" value="<?=$dpatient;?>" </td> 
            <td width=16><input name="shade<?=$key;?>" type="text" id="shade<?=$key;?>" value="<?=$dshade;?>" </td>    
          <? } else { ?>
            <td>&nbsp;</td>   
            <td>&nbsp;</td>   
            <td>&nbsp;</td>     
          <? } ?>
          
          <td><?=$dpcode;?></td>  
          <td><?=$dpename;?></td>            
          <? if ($dupecode=='N') { ?>  
            <td width=16><input name="tooth<?=$key;?>" type="text" id="tooth<?=$key;?>" value="<?=$dtooth;?>" </td>               
          <? } else { ?>
            <td>&nbsp;</td> 
          <? } ?>          
          
          <td><?=$row['tbpcode'];?></td>
          <td width=16><input name="tbpename<?=$key;?>" type="text" id="tbpename<?=$key;?>" value="<?=$row['tbpename'];?>" </td> 
          <td width=16><input name="manufacturer<?=$key;?>" type="text" id="manufacturer<?=$key;?>" value="<?=$row['tbmanufacturer'];?>" </td> 
          <td width=16><input name="ceno<?=$key;?>" type="text" id="ceno<?=$key;?>" value="<?=$row['tbceno'];?>" </td> 
          <td width=16><input name="lotno<?=$key;?>" type="text" id="lotno<?=$key;?>" value="<?=$row['tblotno'];?>" </td> 
          <td width=16><input name="risok<?=$key;?>" type="checkbox" id="risok<?=$key;?>" value="Y" </td>    
      </tr> 
      <?  
      }   
  ?>    
  <tr>
    <td>&nbsp;</td>
    <td><input type="hidden" name="bdate"  value="<?=$bdate;?>">    
        <input type="hidden" name="edate"  value="<?=$edate;?>">
        <input type="hidden" name="occ01"  value="<?=$occ01;?>">                       
        <input type="hidden" name="action" value="save">
        <input type="submit" name="Submit" value="更新">        
    </td>
  </tr>
</table>  
</form>
