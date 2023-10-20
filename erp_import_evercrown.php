<?
//檢查七天前到 但未出貨的清單
session_start();   
include("_data.php");  
date_default_timezone_set('Asia/Taipei'); 
$edate=date('Y-m-d', strtotime("-6 days"));
error_reporting(E_ALL);  
require_once 'classes/PHPExcel.php'; 
require_once 'classes/PHPExcel/IOFactory.php';  
$objPHPExcel = new PHPExcel();
// Set properties
$objPHPExcel ->getProperties()->setCreator("Frank")
           ->setLastModifiedBy("Frank")
           ->setTitle("$edate 到貨未做處理的RX#")
           ->setSubject("$edate 到貨未做處理的RX#")
           ->setDescription("$edate 到貨未做處理的RX#")
           ->setKeywords("$edate 到貨未做處理的RX#")
           ->setCategory("$edate 到貨未做處理的RX#");

// Add some data
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A1', $edate. ' 到貨未做處理的RX#'); 
$objPHPExcel->setActiveSheetIndex(0) ->mergeCells('A1:E1');  
$objPHPExcel->setActiveSheetIndex(0)   
            ->setCellValue('A3', '客戶') 
            ->setCellValue('B3', 'RX #')   
            ->setCellValue('C3', '訂單號')   
            ->setCellValue('D3', '工單號')   
            ->setCellValue('E3', '品代')    
            ->setCellValue('F3', '開單人員');    
                     
          
$s2= "select occ01, occ02, sfb01, sfb22, sfbud02, sfb05, ima02, oea14, gen02 from sfb_file, ima_file, oea_file, occ_file, gen_file " .
     "where sfb28 is null and sfb81 = to_date('$edate','yy/mm/dd')   and sfb04<7  and sfb22=oea01 and oea04=occ01 and oea14=gen01 " .
     "and sfb05=ima01 and ta_ima003!='Y'   and ta_ima004!='Y'  and ta_ima005!='Y' " .
     "and sfb01 not in ( select sfl02 from sfl_file) "  .
     "and sfb01 not in (select tc_ohf001 from tc_ohf_file where tc_ohf008 is null) " .
     "and sfb01 not in ( select tc_ogb001 from tc_ogb_file) " .
     "order by occ01, sfb01 " ;
 
$erp_sql2 = oci_parse($erp_conn1,$s2 );  
oci_execute($erp_sql2);  
$y=4;   
$rec=0;                                                 
while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {      
    $objPHPExcel->setActiveSheetIndex(0)                 
                ->setCellValue('A'. $y, $row2['OCC01']. ' -- '.$row2['OCC02'])
                ->setCellValue('B'. $y, $row2['SFBUD02'])   
                ->setCellValue('C'. $y, $row2['SFB22'])                   
                ->setCellValue('D'. $y, $row2['SFB01'])
                ->setCellValue('E'. $y, $row2['SFB05']. ' -- '.$row2['OCC02']) 
                ->setCellValue('F'. $y, $row2['OEA14']. ' -- '.$row2['GEN02']) ;   
    $y++;
    $rec++;
}               

$objPHPExcel->setActiveSheetIndex(0)                 
          ->setCellValue('A'. $y, '-- 以下無資料 --');
// Rename sheet
$objPHPExcel->getActiveSheet()->setTitle($edate. ' 到貨未做處理的RX#');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');   
$filename='email/' . $edate . '_rxinveden.xls';
$objWriter->save($filename);    

if ($rec>0) {          
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
    $m->send("frank@vedenlabs.com", $edate . ' RX delay in Veden') ;  
    $m->send("cs@vedenlabs.com", $edate . ' RX delay in Veden') ;  
}
