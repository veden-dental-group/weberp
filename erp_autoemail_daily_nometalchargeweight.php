<?
//檢查昨天沒有填研磨後重量的工單
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
           ->setTitle("無研磨後重量")
           ->setSubject("無研磨後重量")
           ->setDescription("無研磨後重量")
           ->setKeywords("無研磨後重量")
           ->setCategory("無研磨後重量");

// Add some data
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A1', $edate. ' 無研磨後重量');
$objPHPExcel->setActiveSheetIndex(0) ->mergeCells('A1:G1');
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A3', '製處')
            ->setCellValue('B3', '客戶')
            ->setCellValue('C3', '出貨日期')
            ->setCellValue('D3', 'RX #')
            ->setCellValue('E3', '工單號碼')
            ->setCellValue('F3', '品名')
            ->setCellValue('G3', '品代');



$s2= "select sfbud02, sfb82, gem02, to_char(tc_oga002,'yyyy-mm-dd') OutDate, sfb01, sfb05, ima02, occ01, occ02 " .
     "from sfe_file, sfb_file, gem_file, ima_file,tc_ogb_file, tc_oga_file, oea_file, occ_file " .
     "where sfe16>0 and ta_sfe002=0 and sfe06='1' " .
     "and sfe01=sfb01 " .
     "and sfb82=gem01 " .
     "and sfb05=ima01 " .
     "and sfe01=tc_ogb002 and tc_ogb001=tc_oga001 and tc_oga002 = to_date('$edate','yy/mm/dd') " .
     "and sfb22=oea01 and oea04=occ01 ".
     "order by gem02, occ01, tc_oga002, sfbud02 ";


$erp_sql2 = oci_parse($erp_conn1,$s2 );
oci_execute($erp_sql2);
$y=4;
$rec=0;
while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
    $objPHPExcel->setActiveSheetIndex(0)
                ->setCellValue('A'. $y, $row2['GEM02'])
                ->setCellValue('B'. $y, $row2['OCC01']. ' -- '.$row2['OCC02'])
                ->setCellValue('C'. $y, $row2['OUTDATE'])
                ->setCellValue('D'. $y, $row2['SFBUD02'])
                ->setCellValue('E'. $y, $row2['SFB01'])
                ->setCellValue('F'. $y, $row2['SFB05'])
                ->setCellValue('G'. $y, $row2['IMA02']);

    $y++;
    $rec++;
}

$objPHPExcel->setActiveSheetIndex(0)
          ->setCellValue('A'. $y, '-- 以下無資料 --');
// Rename sheet
$objPHPExcel->getActiveSheet()->setTitle('無研磨後重量');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
// $filename='email/' . $edate . '_nometalchargeweight.xls';
// $objWriter->save($filename);
$filename='report/' . $edate . '_nometalchargeweight.xls';
$objWriter->save($filename);

//if ($rec > 0) {
	// modify by mao 2013/08/07

	require 'PHPMailer/PHPMailerAutoload.php';
  $report_name = $edate . ' 無研磨後重量';

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
//}

?>
