<?
//檢查昨天的訂單號 但訂單日期不一樣的
session_start();
include("_data.php");
date_default_timezone_set('Asia/Taipei');
$edate=date('y-m-d', strtotime("-2 days"));
$code='B531-'.date('ymd', strtotime("-2 days")).'%';


error_reporting(E_ALL);
require_once 'classes/PHPExcel.php';
require_once 'classes/PHPExcel/IOFactory.php';

//報工清單
$objPHPExcel = new PHPExcel();
// Set properties
$objPHPExcel ->getProperties()->setCreator("Frank")
           ->setLastModifiedBy("Frank")
           ->setTitle("更改訂單日期的訂單資料")
           ->setSubject("更改訂單日期的訂單資料")
           ->setDescription("更改訂單日期的訂單資料")
           ->setKeywords("更改訂單日期的訂單資料")
           ->setCategory("更改訂單日期的訂單資料");

// Add some data
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A1', $edate . ' 訂單 更改到貨日期清單 ');
$objPHPExcel->setActiveSheetIndex(0) ->mergeCells('A1:E1');
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A3', '客戶')
            ->setCellValue('B3', '客戶')
            ->setCellValue('C3', 'RX #')
            ->setCellValue('D3', '訂單號碼')
            ->setCellValue('E3', 'Order Date') ;



$s2= "select oea01, to_char(oea02,'yyyy/mm/dd') oea02, ta_oea006, oea04, occ02 " .
     "from oea_file, occ_file " .
     "where oea01 like '$code' and oea02!=to_date('$edate','yy/mm/dd') and oea04=occ01 " .
     "order by oea04, ta_oea006 ";

$erp_sql2 = oci_parse($erp_conn2,$s2 );
oci_execute($erp_sql2);
$y=4;
while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
    $objPHPExcel->setActiveSheetIndex(0)
                ->setCellValue('A'. $y, $row2['OEA04'])
                ->setCellValue('B'. $y, $row2['OCC02'])
                ->setCellValue('C'. $y, $row2['TA_OEA006'].'  ')
                ->setCellValue('D'. $y, $row2['OEA01'])
                ->setCellValue('E'. $y, $row2['OEA02']);
    $y++;
}
$objPHPExcel->setActiveSheetIndex(0)
          ->setCellValue('A'. $y, '-- 以下無資料 --');
// Rename sheet
$objPHPExcel->getActiveSheet()->setTitle('更改訂單日期的訂單資料 ');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
// $filename='email/' . $edate . '_changeorderdatelist.xls';
// $objWriter->save($filename);
$filename='report/' . $edate . '_changeorderdatelist.xls';
$objWriter->save($filename);


  require 'PHPMailer/PHPMailerAutoload.php';
  $report_name = $edate . ' 訂單更改到貨日期清單';

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
  $mail->addAddress('fa3@veden.dental', 'Veden');
  $mail->addAddress('steven.jiao@veden.dental', 'Veden');
  $mail->Subject = "Veden $report_name ";
  $mail->Body = "Veden $report_name ";
  $mail->addAttachment($filename, $filename);

  $mail->send();

?>
