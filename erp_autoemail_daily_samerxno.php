<?
//檢查當天相同RX # 卻有兩張以上的訂單
session_start();
include("_data.php");
date_default_timezone_set('Asia/Taipei');
$edate=date('Y-m-d', strtotime("-1 days"));
error_reporting(E_ALL);
require_once 'classes/PHPExcel.php';
require_once 'classes/PHPExcel/IOFactory.php';
$objPHPExcel = new PHPExcel();
// Set properties
$objPHPExcel ->getProperties()->setCreator("Frank")
           ->setLastModifiedBy("Frank")
           ->setTitle("相同RX#的訂單")
           ->setSubject("相同RX#的訂單")
           ->setDescription("相同RX#的訂單")
           ->setKeywords("相同RX#的訂單")
           ->setCategory("相同RX#的訂單");

// Add some data
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A1', $edate. ' 相同 RX#, 有一張以上訂單');
$objPHPExcel->setActiveSheetIndex(0) ->mergeCells('A1:E1');
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A3', '客戶')
            ->setCellValue('B3', 'RX #')
            ->setCellValue('C3', '開單日期')
            ->setCellValue('D3', '訂單號')
            ->setCellValue('E3', '開單人員');


$s2= "select oea04, occ02, ta_oea006, to_char(oeadate,'mm-dd-yyyy') oeadate, oea01, oea14, gen02 from vd210.oea_file, vd210.occ_file, vd210.gen_file where ta_oea006 in " .
     "(select ta_oea006 from (select oea04, ta_oea006, count(*) total from vd210.oea_file where oea02 =to_date('$edate','yy/mm/dd')  group by oea04, ta_oea006 having count(*) > 1)) " .
     "and oea04=occ01 and oea14=gen01 " .
     "order by oea04, ta_oea006, oeadate ";

$erp_sql2 = oci_parse($erp_conn2,$s2 );
oci_execute($erp_sql2);
$y=4;
$rec=0;
while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
    $objPHPExcel->setActiveSheetIndex(0)
                ->setCellValue('A'. $y, $row2['OEA04']. ' -- '.$row2['OCC02'])
                ->setCellValue('B'. $y, $row2['TA_OEA006'])
                ->setCellValue('C'. $y, $row2['OEADATE'])
                ->setCellValue('D'. $y, $row2['OEA01'])
                ->setCellValue('E'. $y, $row2['OEA14']. ' -- '.$row2['GEN02']) ;
    $y++;
}

$objPHPExcel->setActiveSheetIndex(0)
          ->setCellValue('A'. $y, '-- 以下無資料 --');
// Rename sheet
$objPHPExcel->getActiveSheet()->setTitle('相同RX#的訂單');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
// $filename='email/' . $edate . '_samerxno.xls';
// $objWriter->save($filename);
$filename='report/' . $edate . '_samerxno.xls';
$objWriter->save($filename);


	$report_name = $edate . ' 相同RX#的訂單';

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
