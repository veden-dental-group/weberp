<?
//檢查七天前到 但未出貨的清單
session_start();
include("_data.php");
date_default_timezone_set('Asia/Taipei');
$edate=date('Y-m-d', strtotime("-10 days"));
//$edate='2014-09-03';
error_reporting(E_ALL);
require_once 'classes/PHPExcel.php';
require_once 'classes/PHPExcel/IOFactory.php';
$objPHPExcel = new PHPExcel();
// Set properties
$objPHPExcel ->getProperties()->setCreator("Frank")
           ->setLastModifiedBy("Frank")
           ->setTitle("$edate 前(含)到貨未做處理的RX#")
           ->setSubject("$edate 前(含)到貨未做處理的RX#")
           ->setDescription("$edate 前(含)到貨未做處理的RX#")
           ->setKeywords("$edate 前(含)到貨未做處理的RX#")
           ->setCategory("$edate 前(含)到貨未做處理的RX#");

// Add some data
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A1', $edate. ' 前(含)到貨未做處理的RX#');
$objPHPExcel->setActiveSheetIndex(0) ->mergeCells('A1:E1');
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A3', '客戶')
            ->setCellValue('B3', '開單日期')
            ->setCellValue('C3', 'RX #')
            ->setCellValue('D3', '訂單號')
            ->setCellValue('E3', '工單號')
            ->setCellValue('F3', '品代')
            ->setCellValue('G3', '開單人員');


$s2= "select occ01, occ02, to_char(sfb81,'yy-mm-dd') sfb81, sfb01, sfb22, sfbud02, sfb05, ima02, oea14, gen02 from sfb_file, ima_file, oea_file, occ_file, gen_file " .
     "where sfb28 is null and sfb81 <= to_date('$edate','yy/mm/dd')   and sfb04<7  and sfb22=oea01 and oea04=occ01 and oea14=gen01 " .
     "and sfb05=ima01 and ta_ima003!='Y'   and ta_ima004!='Y'  and ta_ima005!='Y' " .
     "and sfb05!='1Z130' " .
     "and sfb38 is null " . //有結案的不算
     "and not exists ( select 1 from sfl_file where sfb01= sfl02)  "  .
     "and not exists ( select 1 from tc_ohf_file where sfb01= tc_ohf001 and tc_ohf008 is null) " .
     "and not exists ( select 1 from tc_ogb_file where sfb01= tc_ogb002 )  " .
     "order by occ01, sfb81, sfb01 " ;

$erp_sql2 = oci_parse($erp_conn1,$s2 );
oci_execute($erp_sql2);
$y=3;
$rec=0;
$oldocc01='';
while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
    if ($oldocc01!=$row2['OCC01']) {
        $y++;
        $oldocc01=$row2['OCC01'];
    }
    $objPHPExcel->setActiveSheetIndex(0)
                ->setCellValue('A'. $y, $row2['OCC01']. ' -- '.$row2['OCC02'])
                ->setCellValue('B'. $y, $row2['SFB81'])
                ->setCellValue('C'. $y, $row2['SFBUD02'])
                ->setCellValue('D'. $y, $row2['SFB22'])
                ->setCellValue('E'. $y, $row2['SFB01'])
                ->setCellValue('F'. $y, $row2['SFB05']. ' -- '.$row2['IMA02'])
                ->setCellValue('G'. $y, $row2['OEA14']. ' -- '.$row2['GEN02']) ;
    $y++;
    $rec++;
}

$objPHPExcel->setActiveSheetIndex(0)
          ->setCellValue('A'. $y, '-- 以下無資料 --');
// Rename sheet
$objPHPExcel->getActiveSheet()->setTitle($edate. ' 前(含)到貨未做處理的RX#');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
// $filename='email/' . $edate . '_rxinveden.xls';
// $objWriter->save($filename);
$filename='report/' . $edate . '_rxinveden.xls';
$objWriter->save($filename);

	$report_name = $edate . ' 前(含)到貨未做處理的RX#';


  require 'PHPMailer/PHPMailerAutoload.php';

  $mail = new PHPMailer();
  $mail->isSMTP();
  $mail->SMTPDebug = 0;
  $mail->Debugoutput = 'html';

  $mail->CharSet = 'utf-8';
  $mail->Host = $_SESSION['emailserver'];
  $mail->Port = $_SESSION['emailport'];
  $mail->SMTPAuth = $_SESSION['emailauth'];
  $mail->Username = $_SESSION['emailusername'];
  $mail->Password = $_SESSION['emailpassword'];
  //Set who the message is to be sent from
  $mail->setFrom('report@veden.dental', 'Daily Report');
  $mail->addReplyTo('report@veden.dental', 'Daily Report');
  $mail->addAddress('it@veden.dental', 'Veden');
  $mail->addAddress('cs@veden.dental', 'Veden');
  $mail->Subject = "Veden $report_name ";
  $mail->Body    = "Veden $report_name ";
  $mail->addAttachment($filename, $filename);
  $mail->send();
