<?
//每天11:30AM寄送昨天delay的資料

session_start();   
include("_data.php");  
date_default_timezone_set('Asia/Taipei'); 
$tdate=date('Y-m-d', strtotime("-1 days")); 
$edate=date('Y-m-d', strtotime("-5 days"));
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
//            ->setCellValue('C3', 'CASE號碼')   
//            ->setCellValue('D3', '訂單號碼')   
//            ->setCellValue('E3', '訂單序號')    
//            ->setCellValue('F3', '工單號碼')  
//            ->setCellValue('G3', '到貨日期') 
//            ->setCellValue('H3', '解傳真日期') 
//            ->setCellValue('I3', '客戶應交貨日期') 
//            ->setCellValue('J3', 'Delay 天數') 
//            ->setCellValue('K3', '預計出貨日期') 
//            ->setCellValue('L3', '產品') 
//            ->setCellValue('M3', 'Delay 原因') 
//            ->setCellValue('N3', '責任工序') 
//            ->setCellValue('O3', '目前工序')
//            ->setCellValue('P3', '客戶')
//            ->setCellValue('Q3', '回覆時間')
//            ->setCellValue('R3', '板色'); 
                     
          
$query2= "select * from casedelay where d01='$edate' and d16='' order by d14, d11,d02,d03,d031";     
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
$filename='email/' . $tdate . '_DelayCasesReport.xls';
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
}
