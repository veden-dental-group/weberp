<?php
  session_start();
  $pagetitle = "車間作業 &raquo; 今天出貨卻未刷卡工單";
  include("_data.php");
  date_default_timezone_set('Asia/Taipei');

  $bdate = $argv[1] ? $argv[1] : date('Y-m-d');
  //$bdate='2015-01-17';


  $filename='templates/casenocheckin.xls';

  error_reporting(E_NONE);
  require_once 'classes/PHPExcel.php';
  require_once 'classes/PHPExcel/IOFactory.php';
  $objReader = PHPExcel_IOFactory::createReader('Excel5');
  $objPHPExcel = $objReader->load($filename);
  $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
  $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
  // Set properties
  $objPHPExcel ->getProperties()->setCreator('Frank' )
               ->setLastModifiedBy('Frank')
               ->setTitle('Frank')
               ->setSubject('Frank')
               ->setDescription('Frank')
               ->setKeywords('Frank')
               ->setCategory('Frank');

  $y=4;
  $s2= "select gem02, sfb01, sfb22, sfbud02, occ02, ima02 from
          (select sfb01, sfb22, sfb05, sfb82, SFBUD02 from sfb_file
          where  sfb01 in ( select tc_ogb002 from tc_ogb_file, tc_oga_file where tc_ogb001=tc_oga001 and tc_oga002 = to_date('$bdate','yy/mm/dd') )
          and sfb01 not in ( select tc_srg001 from tc_srg_file where tc_srg007 is not null or tc_srg010 is not null or tc_srg013 is not null)
          and sfb82 !='6AZ000'
          and sfb05 not like '1Z%' and sfb28 is null
          and sfb05 not like '2Z%' ), gem_file, oea_file, occ_file, ima_file
          where sfb82=gem01
          and sfb22=oea01 and ta_oea004='1'
          and oea04=occ01
          and sfb05=ima01
          order by gem01" ;
  $erp_sql2 = oci_parse($erp_conn,$s2 );
  oci_execute($erp_sql2);
  while ($row = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A'. $y, $bdate)
                    ->setCellValue('B'. $y, $row["GEM02"])
                    ->setCellValue('C'. $y, $row["SFB01"])
                    ->setCellValue('D'. $y, $row["SFB22"])
                    ->setCellValue('E'. $y, $row["SFBUD02"])
                    ->setCellValue('F'. $y, $row["OCC02"])
                    ->setCellValue('G'. $y, $row["IMA02"]);
        $y++;
  }

  $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A'. $y, '--以下無資料--');
  // Rename sheet
  $objPHPExcel->getActiveSheet()->setTitle($bdaet . '出貨 但未刷卡工單');

  // Set active sheet index to the first sheet, so Excel opens this as the first sheet
  $objPHPExcel->setActiveSheetIndex(0);

  $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
  // $filename='email/' . $edate . '_DeliveriedCasesReportv3.xls';
  // $objWriter->save($filename);
  $filename='report/' . $bdate . '_noncheckin.xls';
  $objWriter->save($filename);

  $report_name = $bdate . ' 已出貨未報工清單';

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
  $mail->addAddress('m@veden.dental', 'Veden');
  $mail->addAddress('it@veden.dental', 'Veden');
  //$mail->addAddress('frank@vedendentalgroup.com', 'Veden');
  $mail->Subject = "Veden $report_name ";
  $mail->Body    = "Veden $report_name ";
  $mail->addAttachment($filename, $filename);
  $mail->send();
?>
