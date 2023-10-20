<?
//檢查三天前到 但未出貨的清單
session_start();   
include("_data.php");  
date_default_timezone_set('Asia/Taipei'); 
$tdate=date('Y-m-d'); 
$edate=date('Y-m-d', strtotime("-3 days"));

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
            ->setCellValue('A1', $edate. ' (含)前到貨但未出貨的RX#'); 
$objPHPExcel->setActiveSheetIndex(0) ->mergeCells('A1:Q1');  
//$objPHPExcel->setActiveSheetIndex(0)   
//            ->setCellValue('A3', '序號') 
//            ->setCellValue('B3', '製處')   
//            ->setCellValue('C3', 'CASE號碼')   
//            ->setCellValue('D3', '訂單號碼')   
//            ->setCellValue('E3', '工單號碼')  
//            ->setCellValue('F3', '到貨日期') 
//            ->setCellValue('G3', '解傳真日期')     
//            ->setCellValue('H3', 'Delay 天數')   
//            ->setCellValue('I3', '產品') 
//            ->setCellValue('J3', 'Delay 原因') 
//            ->setCellValue('K3', '客戶')      
                     
          
$s2= "select sfb82, gem02, sfbud02, sfb08, sfb22, sfb221, sfb01, to_char(sfb81,'mm-dd-yyyy') sfb81,  to_char(ta_oea005,'mm-dd-yyyy') ta_oea005, sfb05, ima02, imaud07, occ01, occ02 " .
     "from sfb_file,ima_file, oea_file, gem_file, occ_file " .
     "where sfb81<=to_date('$edate','yy/mm/dd') and sfb81>=to_date('120201','yy/mm/dd')   " . //某天以前
     "and sfb28 is null " . //未結案
     "and not exists ( select 1 from oga_file where oga16=sfb22) " . //未出貨
     "and sfb05=ima01 and ( ta_ima003 !='Y' and ta_ima004 !='Y' and ta_ima005 !='Y') " . //未配件
     "and not exists (select 1 from tc_ohf_file where tc_ohf001=sfb01) " . //有傳真過都會delay都不能算
     "and sfb22=oea01 and oea04=occ01 " .
     "and sfb05=ima01 " .
     "and sfb82=gem01 " .
     "order by sfb82,sfb22,sfb221";
 
$erp_sql2 = oci_parse($erp_conn1,$s2 );  
oci_execute($erp_sql2);  
$y=2;   
$i=0;   
$j=0;
$oldrxno=''; 
$oldgem02='';                                             
while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {   
    $faxdate=findfaxdatebysfb01($erp_conn1,$row2['SFB01']);
    $nowstage=findstagebysfb01($erp_conn1,$row2['SFB01']);
    if ($oldgem02!=$row2['GEM02']) {
        $oldgem02=$row2['GEM02'];
        $i=0;  //不同製處就由1開始算
        $y+=2;
    } 
    if ($oldrxno!=$row2['SFBUD02']){
        $rxno=$row2['SFBUD02'];
        $sfb22=$row2['SFB22'];
        $sfb221=$row2['SFB221'];
        $oldrxno=$row2['SFBUD02'];
        $i++;
        $j++;
    } else {
        $rxno='';
        $sfb22='';      
        $sfb221='';
    }
    
    $objPHPExcel->setActiveSheetIndex(0)     
                ->setCellValue('A'. $y, $i)             
                ->setCellValue('B'. $y, $row2['GEM02'])
                ->setCellValue('C'. $y, $row2['OCC01'].'--'.$row2['OCC02'])  
                ->setCellValue('D'. $y, $rxno)   
                ->setCellValue('E'. $y, $sfb22 . ' ' . $sfb221)  
                ->setCellValue('F'. $y, $row2['SFB01']) 
                ->setCellValue('G'. $y, $row2['SFB81'])    
                ->setCellValue('H'. $y, $faxdate)      
                ->setCellValue('I'. $y, $row2['TA_OEA005'])   
                ->setCellValue('J'. $y, $row2['SFB05'].'--'.$row2['IMA02']) 
                ->setCellValue('K'. $y, $nowstage);  
    if (($y%2)==0){
        $objPHPExcel->setActiveSheetIndex(0) ->getStyle('A'.$y.':R'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');    
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
$filename='email/' . $edate . '_delaycases.xls';
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
    $m->send("frank@vedenlabs.com", $edate . ' delay cases') ;  
    $m->send("cs@vedenlabs.com", $edate . ' delay cases') ;  
    $m->send("m@vedenlabs.com", $edate . ' delay cases') ;   
    $m->send("fa2@vedenlabs.com", $edate . ' delay cases') ;    
}
