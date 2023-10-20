<?
//檢查有無當天出貨的訂單中 有工單沒秤重的
session_start();
include("_data.php");

$edate=date('Y-m-d', strtotime("-1 days"));
error_reporting(E_ALL);
require_once 'classes/PHPExcel.php';
require_once 'classes/PHPExcel/IOFactory.php';
$objPHPExcel = new PHPExcel();
// Set properties
$objPHPExcel ->getProperties()->setCreator("Frank")
           ->setLastModifiedBy("Frank")
           ->setTitle("未秤重工單")
           ->setSubject("未秤重工單")
           ->setDescription("未秤重工單")
           ->setKeywords("未秤重工單")
           ->setCategory("未秤重工單");

// Add some data
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A1', 'RX# 和 '.$edate. ' 出貨的 RX# 相同, 但未在 '. $edate .' 出貨的工單');
$objPHPExcel->setActiveSheetIndex(0) ->mergeCells('A1:E1');
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A3', 'RX #')
            ->setCellValue('B3', '客戶')
            ->setCellValue('C3', '工單號')
            ->setCellValue('D3', '訂單號')
            ->setCellValue('E3', '品代')
            ->setCellValue('F3', '品名');

$s2= "select sfb01, sfb22, sfbud02, sfb05, ima02  from
        (select sfb01, sfb22, sfb05, sfbud02, sfb28 from sfb_file where sfbud02 in
          (select tc_ogb011 from tc_ogb_file, tc_oga_file where tc_ogb001=tc_oga001 and tc_oga002 =to_date('$edate','yy/mm/dd'))) , ima_file
      where
      sfb01 not in (select tc_ogb002 from tc_ogb_file, tc_oga_file where tc_ogb001=tc_oga001 and tc_oga002<=to_date('$edate','yy/mm/dd'))
      and sfb28 is null
      and sfb01 not in (select sfl02 from sfl_file)
      and sfb01 not in (select sfe01 from sfe_file where sfe06='4')
      and sfb01 not in (select distinct tc_ohf001 from tc_ohf_file where tc_ohf008 is null)
      and sfb05=ima01 and ta_ima003!='Y' and ta_ima004!='Y' and ta_ima005!='Y' order by sfbud02  ";

$erp_sql2 = oci_parse($erp_conn1,$s2 );
oci_execute($erp_sql2);
$y=4;
$rec=0;
while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
    $s3="select oea04, occ02 from oea_file, occ_file where oea04=occ01 and oea01='" . $row2['SFB22'] . "' ";
    $erp_sql3 = oci_parse($erp_conn1,$s3 );
    oci_execute($erp_sql3) ;
    $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);
    $objPHPExcel->setActiveSheetIndex(0)
                ->setCellValue('A'. $y, $row2['SFBUD02'])
                ->setCellValue('B'. $y, $row3['OEA04']. ' -- '.$row3['OCC02'])
                ->setCellValue('C'. $y, $row2['SFB01'])
                ->setCellValue('D'. $y, $row2['SFB22'])
                ->setCellValue('E'. $y, $row2['SFB05'])
                ->setCellValue('F'. $y, $row2['IMA02']);
    $y++;
    $rec++;
}

$objPHPExcel->setActiveSheetIndex(0)
          ->setCellValue('A'. $y, '-- 以下無資料 --');
// Rename sheet
$objPHPExcel->getActiveSheet()->setTitle('未秤重工單');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
// $filename='email/' . $edate . '_ticketnoweight.xls';
// $objWriter->save($filename);
$filename='report/' . $edate . '_ticketnoweight.xls';
$objWriter->save($filename);

if ($rec>0) {

	require 'PHPMailer/PHPMailerAutoload.php';
  $report_name = $edate . ' 出貨未秤重工單';

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
}

?>
