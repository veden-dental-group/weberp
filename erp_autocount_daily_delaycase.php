<?
//每天重算各製處到貨顆數/Delay顆數/內返顆數
//星期五, 六 PTL不計算delay

session_start();   
include("_data.php");  
date_default_timezone_set('Asia/Taipei'); 
$tdate=date('Y-m-d');
$edate=date('Y-m-d', strtotime("-4 days"));
//$edate=date('Y-m-d');

//find the lastest fax date
function findfaxdatebysfb01($erp_conn1,$sfb01){
  $s3= "select to_char(tc_ohf008,'mm-dd-yyyy') tc_ohf008, tc_ohf009 from tc_ohf_file where tc_ohf001='$sfb01' order by tc_ohf008, tc_ohf009 desc";   
  $erp_sql3 = oci_parse($erp_conn1,$s3 );  
  oci_execute($erp_sql3);  
  $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);    
  if (is_null($row3['TC_OHF008'])){
      $faxdate='';
  } else {
      $faxdate= $row3['TC_OHF008']. ' ' . $row3['TC_OHF009'];
  }
  return $faxdate;  
}

//find the case stage
function findstagebysfb01($erp_conn1,$sfb01){
  $s3= "select tc_srg001, tc_srg006,tc_srg010, tc_srg018, tc_srg022, ecb17 from tc_srg_file, ecb_file where tc_srg001='$sfb01' and tc_srg005 is not null and tc_srg003=ecb01 and tc_srg004=ecb03 ";   
  $erp_sql3 = oci_parse($erp_conn1,$s3 );  
  oci_execute($erp_sql3);  
  $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);    
  $tc_srg001=$row3['TC_SRG001'];
  $tc_srg006=$row3['TC_SRG006']; 
  $tc_srg010=$row3['TC_SRG010'];
  $tc_srg018=$row3['TC_SRG018'];
  $tc_srg022=$row3['TC_SRG022'];
  $ecb17=$row3['ECB17'];  
  if (is_null($tc_srg001)){
      $nowstage='工序不明';
  } else {
      if ($tc_srg018=='Y') {  //返工    
          if ($tc_srg006=='Y') { //要QC                
              if (is_null($tc_srg022)) {
                  $nowstage=$ecb17;
              } else {
                  $nowstage=$ecb17 . " QC";  
              }      
          } else {
            $nowstage=$ecb17; 
          }        
      } else {
          if ($tc_srg006=='Y') { //要QC                
              if (is_null($tc_srg010)) {  //未出站
                  $nowstage=$ecb17;
              } else {
                  $nowstage=$ecb17. " QC";  
              }      
          } else {
            $nowstage=$ecb17; 
          }
      }
      
  }
  return $nowstage;  
}

//先刪除同一日期的所有資料
$query2= "delete from casedelay where d01='$edate' ";     
$result2= mysql_query($query2);

$weekday=date('w',strtotime($edate)); //0:日  1:一... 5:五  6:六)


if ($weekday=5 or $weekday=6) {  //PTL, LP, F-1 星期五六 不算delay
    $s2= "select sfb82, gem02, sfbud02, sfb08, sfb22, sfb221, sfb01, to_char(oea02,'yyyy-mm-dd') oea02,  to_char(ta_oea005,'yyyy-mm-dd') ta_oea005, sfb05, ima02, imaud07, occ01, occ02 " .
         "from sfb_file,ima_file, oea_file, gem_file, occ_file " .
         //"where sfb81<=to_date('$edate','yy/mm/dd') and sfb81>=to_date('120201','yy/mm/dd')   " . //某天以前  
         "where oea02<=to_date('$edate','yy/mm/dd') and oea04 !='E129001' and oea04 !='E129002' and oea04 != 'E143001' and oea04 != 'E145001' " .
         "and sfb28 is null " . //未結案
         "and not exists ( select 1 from tc_ogb_file where tc_ogb003=sfb22) " . //未出貨
         "and sfb05=ima01 and ( ta_ima003 !='Y' and ta_ima004 !='Y' and ta_ima005 !='Y') " . //未配件
         "and not exists (select 1 from tc_ohf_file where tc_ohf001=sfb01) " . //有傳真過都會delay都不能算
         "and sfb22=oea01 and oea04=occ01 and ( oeaud05='N' or oeaud05 is null ) " . //做die送回case不算delay
         "and sfb05=ima01 and instr(ima02, '客戶來模')=0 " .
         "and sfb82=gem01 " .
         "order by sfb82,sfb22,sfb221";
} else {
    $s2= "select sfb82, gem02, sfbud02, sfb08, sfb22, sfb221, sfb01, to_char(oea02,'yyyy-mm-dd') oea02,  to_char(ta_oea005,'yyyy-mm-dd') ta_oea005, sfb05, ima02, imaud07, occ01, occ02 " .
         "from sfb_file,ima_file, oea_file, gem_file, occ_file " .
         //"where sfb81<=to_date('$edate','yy/mm/dd') and sfb81>=to_date('120201','yy/mm/dd')   " . //某天以前  
         "where oea02<=to_date('$edate','yy/mm/dd') " .
         "and sfb28 is null " . //未結案
         "and not exists ( select 1 from tc_ogb_file where tc_ogb003=sfb22) " . //未出貨
         "and sfb05=ima01 and ( ta_ima003 !='Y' and ta_ima004 !='Y' and ta_ima005 !='Y') " . //未配件
         "and not exists (select 1 from tc_ohf_file where tc_ohf001=sfb01) " . //有傳真過都會delay都不能算
         "and sfb22=oea01 and oea04=occ01 and ( oeaud05='N' or oeaud05 is null ) " . //做die送回case不算delay  
         "and sfb05=ima01 instr(ima02, '客戶來模')=0 " .
         "and sfb82=gem01 " .
         "order by sfb82,sfb22,sfb221";
}
 
$erp_sql2 = oci_parse($erp_conn1,$s2 );  
oci_execute($erp_sql2);  
                                                   
while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {   
    $faxdate=findfaxdatebysfb01($erp_conn1,$row2['SFB01']);
    $nowstage=findstagebysfb01($erp_conn1,$row2['SFB01']);   
    $queryus = "insert into casedelay ( d01,d02,d03,d031,d04,d05,d06,d07,d08,d09,d10,d11,d12,d13,d14,d15,d16,d21,d25,d26 ) values (  
               '" . $edate                      . "',
               '" . safetext($row2['SFBUD02'])  . "',   
               '" . $row2['SFB22']              . "',  
               '" . $row2['SFB221']         . "',  
               '" . $row2['SFB01']          . "',  
               '" . $row2['OEA02']          . "',  
               '" . $row2['TA_OEA005']      . "',  
               '" . $faxdate                . "',  
               '" . $faxdate                . "',  
               '" . $row2['SFB05']          . "',  
               '" . $row2['IMA02']          . "',  
               '" . $row2['OCC01']          . "',  
               '" . safetext($row2['OCC02']). "',  
               '" . $row2['SFB82']          . "',  
               '" . safetext($row2['GEM02']). "',  
               '" . $nowstage               . "','', 
               '"  . $tdate                 . "',  
               "  . $row2['SFB08']          . ",  
               "  . $row2['IMAUD07']        . " )";     
    $resultus = mysql_query($queryus);       
    $msg=mysql_error();
    
}  



//判斷case的合理性 寫到 d16
// 2:3種產品  3:固+活

$queryu = "select distinct d03 from casedelay where d01='$edate'";
$resultu= mysql_query($queryu);
while ($rowu = mysql_fetch_array($resultu)){
    $i=1;
    $name=array();
    $d03=$rowu['d03'];
    $s3= "select sfb05 from sfb_file where sfb22='$d03' order by sfb05";
    $erp_sql3 = oci_parse($erp_conn1,$s3 );  
    oci_execute($erp_sql3);                                                     
    while ($row3 = oci_fetch_array($erp_sql3, OCI_ASSOC)) {   
        $name[$i]=$row3['SFB05'];              //將同一張訂單的所有品代都放到陣列中
        $i++;  
    }
    
    //判斷是否為固+活
    $type1=0;
    $type2=0;
    $type17=0;  
    $type19=0;
    $type24=0;  
    for ($j=1;$j<$i;$j++){                      //若頭位數1 則為固  2則為活
        if (substr($name[$j],0,1)=='1') $type1=1;
        if (substr($name[$j],0,1)=='2') $type2=1;     
        if (substr($name[$j],0,2)=='17') $type17=1;  
        if (substr($name[$j],0,2)=='19') $type19=1; 
        if (substr($name[$j],0,2)=='24') $type24=1;   
    }          
    
    $type='';
    if (($type1+$type2)==2) $type= '2';
    if (($type17+$type19+$type24)==3) $type= '3';  
    
    $queryv="update casedelay set d16='$type' where d03='$d03' and d01='$edate' ";
    $resultv= mysql_query($queryv);   
}


//判斷case的合理性 寫到 d16
// 1: 六顆以上
$queryu = "select distinct d03 from casedelay where d25*d26> 5 and d01='$edate'";
$resultu= mysql_query($queryu);
while ($rowu = mysql_fetch_array($resultu)){
    $d03=$rowu['d03'];  
    $queryv="update casedelay set d16='1' where d03='$d03' and d01='$edate' ";
    $resultv= mysql_query($queryv);  
}


error_reporting(E_ALL);  
require_once 'classes/PHPExcel.php'; 
require_once 'classes/PHPExcel/IOFactory.php';  
//$objPHPExcel = new PHPExcel();
$objReader = PHPExcel_IOFactory::createReader('Excel5');  
$objPHPExcel = $objReader->load("templates/erp_delaycases_4m.xls");  
// Set properties
$objPHPExcel ->getProperties()->setCreator("Frank")
           ->setLastModifiedBy("Frank")
           ->setTitle("$edate (含)前到貨但未出貨的RX#")
           ->setSubject("$edate (含)前到貨但未出貨的RX#")
           ->setDescription("$edate (含)前到貨但未出貨的RXX#")
           ->setKeywords("$edate (含)前到貨但未出貨的RX#")
           ->setCategory("$edate (含)前到貨但未出貨的RX#");

// Add some data
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A1', $tdate . ' delay cases report ---> ' . $edate. ' (含)前到貨但未出貨的RX#'); 
$objPHPExcel->setActiveSheetIndex(0) ->mergeCells('A1:Q1');  
//$objPHPExcel->setActiveSheetIndex(0)   
//            ->setCellValue('A3', '序號') 
//            ->setCellValue('B3', '製處')   
//            ->setCellValue('P3', '客戶')    
//            ->setCellValue('C3', 'CASE號碼')   
//            ->setCellValue('D3', '訂單號碼')   訂單序號')    
//            ->setCellValue('F3', '工單號碼')  
//            ->setCellValue('G3', '到貨日期') 
//            ->setCellValue('H3', '解傳真日期') 
//            ->setCellValue('I3', '客戶應交貨日期') 
//            ->setCellValue('K3', '預計出貨日期') 
//            ->setCellValue('L3', '產品')        
//            ->setCellValue('O3', '目前工序')   
                     
          
$query2= "select * from casedelay where d01='$edate' and d16='' order by d14,d11,d02,d03,d031";     
$result2= mysql_query($query2);

$y=2;   
$i=0;   
$j=0;
$oldrxno=''; 
$oldgem02='';                                             
while ($row2 = mysql_fetch_array($result2)){
    if ($oldgem02!=$row2['d14']) {
        $oldgem02 =$row2['d14'];
        $i=0;  //不同製處就由1開始算
        $y+=2; //列印位置
    } 
    
    if ($oldrxno!=$row2['d02']){
        $rxno=$row2['d02'];
        $d03=$row2['d03'];
        $d031=$row2['d031'];
        $oldrxno=$row2['d02'];
        $i++;
        $j++;
    } else {
        $rxno='';
        $d03='';      
        $d031='';
    }
    
    $objPHPExcel->setActiveSheetIndex(0)     
                ->setCellValue('A'. $y, $i)             
                ->setCellValue('B'. $y, $row2['d14'])   
                ->setCellValue('C'. $y, $row2['d11'].'--'.$row2['d12'])  
                ->setCellValue('D'. $y, $rxno)   
                ->setCellValue('E'. $y, $d03.' ' . $d031)    
                ->setCellValue('F'. $y, $row2['d04']) 
                ->setCellValue('G'. $y, $row2['d05'])    
                ->setCellValue('H'. $y, '')      
                ->setCellValue('I'. $y, $row2['d06'])    
                ->setCellValue('J'. $y, $row2['d09'].'--'.$row2['d10'])     
                ->setCellValue('K'. $y, $row2['d15'])                       
                ->setCellValue('L'. $y, $row2['d25'])      
                ->setCellValue('M'. $y, $row2['d26']);  
    if (($y%2)==0){
        $objPHPExcel->setActiveSheetIndex(0) ->getStyle('A'.$y.':M'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');    
    }               
                
    $y++;       
}         

$y++;     

$objPHPExcel->setActiveSheetIndex(0)                 
          ->setCellValue('A'. $y, '合計: '. $j. ' 組CASES');
$objPHPExcel->setActiveSheetIndex(0) ->mergeCells('A'.$y.':Q'.$y); 
$objPHPExcel->setActiveSheetIndex(0) ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);   
$y++;
$objPHPExcel->setActiveSheetIndex(0)                 
          ->setCellValue('A'. $y, '-- 以下無資料 --');
$objPHPExcel->setActiveSheetIndex(0) ->mergeCells('A'.$y.':Q'.$y);   
$objPHPExcel->setActiveSheetIndex(0) ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);         
// Rename sheet
$objPHPExcel->getActiveSheet()->setTitle($edate. ' (含)前到貨但未出貨的RX#');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');   
// $filename='email/' . $tdate . '_DelayCasesReport.xls';
$filename='report/' . $tdate . '_DelayCasesReport.xls';
$objWriter->save($filename);    

if ($j>0) {          
    ////email
    require_once("_classes.php");
    $_SERVER['SERVER_NAME'] = 'www.vedenlabs.com';
    $regards = "Veden Dental Labs Inc.";
       
    $m = new mailer();
    $m->setMessage($regards);
    $m->setPriority( 'High' );
    $m->setFrom( "frankyu@vedenlabs.com", "Veden Dental Labs Inc"  );
    $m->setReplyTo( "frankyu@vedenlabs.com", "Veden Dental Labs Inc" );  
    $m->attachFile( $filename, $filename, "application/vnd.ms-excel"); 
    $m->send("frank@vedenlabs.com", $tdate . ' delay cases report') ;  
    $m->send("cs@vedenlabs.com", $tdate . ' delay cases report') ;  
    $m->send("m@vedenlabs.com", $tdate . ' delay cases report') ;   
    $m->send("fa3@vedenlabs.com", $tdate . ' delay cases report') ;    
    $m->send("fa6@vedenlabs.com", $tdate . ' delay cases report') ;
}
