<?
session_start();
include("_data.php");
date_default_timezone_set('Asia/Taipei');
$edate=date('Y-m-d');
error_reporting(E_ALL);
require_once 'classes/PHPExcel.php';
require_once 'classes/PHPExcel/IOFactory.php';
$objPHPExcel = new PHPExcel();
// Set properties
$objPHPExcel ->getProperties()->setCreator("Frank")
           ->setLastModifiedBy("Frank")
           ->setTitle("庫存不足")
           ->setSubject("庫存不足")
           ->setDescription("庫存不足")
           ->setKeywords("庫存不足")
           ->setCategory("庫存不足");

// Add some data
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A1', '品代')
            ->setCellValue('B1', '品名')
            ->setCellValue('C1', '目前庫存')
            ->setCellValue('D1', '安全庫存')
            ->setCellValue('E1', '差額');

$s2= "select ima01, ima02, img10, ima27 from ima_file, ( select img01, sum(img10) img10 from img_file group by img01) " .
     "where ima27>0 and ima01=img01 and substr(ima06,1,1)!='9' and img10 <=ima27 order by ima01";
$erp_sql2 = oci_parse($erp_conn1,$s2 );
oci_execute($erp_sql2);
$y=2;
$rec=0;
while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
    $objPHPExcel->setActiveSheetIndex(0)
                ->setCellValue('A'. $y, $row2['IMA01'])
                ->setCellValue('B'. $y, $row2['IMA02'])
                ->setCellValue('C'. $y, $row2['IMG10'])
                ->setCellValue('D'. $y, $row2['IMA27'])
                ->setCellValue('E'. $y, $row2['IMG10']-$row2['IMA27']);
    $y++;
    $rec++;
}

$objPHPExcel->setActiveSheetIndex(0)
          ->setCellValue('A'. $y, '-- 以下無資料 --');
// Rename sheet
$objPHPExcel->getActiveSheet()->setTitle('庫存不足');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
// $filename='email/' . $edate . '_stocknotenough.xls';
// $objWriter->save($filename);
$filename='report/' . $edate . '_stocknotenough.xls';
$objWriter->save($filename);

if ($rec > 0 ) {

  require 'PHPMailer/PHPMailerAutoload.php';
  $report_name =  $edate . ' 庫存不足清單';

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
  $mail->addAddress('em2@veden.dental', 'Veden');
  $mail->addAddress('p3@veden.dental', 'Veden');
  $mail->addAddress('e2@veden.dental', 'Veden');
  $mail->addAddress('e8@veden.dental', 'Veden');
  $mail->addAddress('p5@veden.dental', 'Veden');
  $mail->Subject = "Veden $report_name ";
  $mail->Body = "Veden $report_name ";
  $mail->addAttachment($filename, $filename);

  $mail->send();

}

?>
