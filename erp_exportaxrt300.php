<?php
  session_start();
  $pagetitle = "帳單組 &raquo; 匯出關係人交易";
  include("_data.php");
  include("_erp.php");
  //auth("erp_clientdailyinvoice.php");

  if (is_null($_GET['bdate'])) {
    $bdate =  date('Y-m-d');
  } else {
    $bdate=$_GET['bdate'];
  }

  if ($_GET["submit"]=="Excel") {

        $filename="templates/erp_exportaxrt300.xls";
        error_reporting(E_ALL);

        require_once 'classes/PHPExcel.php';
        require_once 'classes/PHPExcel/IOFactory.php';
        $objReader = PHPExcel_IOFactory::createReader('Excel5');
        $objPHPExcel = $objReader->load("$filename");

        $objPHPExcel ->getProperties()->setCreator( 'axrt300')
                     ->setLastModifiedBy('axrt300')
                     ->setTitle('axrt300')
                     ->setSubject('axrt300')
                     ->setDescription('axrt300')
                     ->setKeywords('axrt300')
                     ->setCategory('axrt300');

        $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);
        $s2="select omb31, omb32, omb04, omb06, omb05, omb12, omb13, omb14, omb15, omb16 " .
            "from oma_file, omb_file " .
            "where oma01=omb01 and oma03='V210000' and oma00='12' and oma02=to_date($bdate1,'yy/mm/dd') ".
            "order by omb31, omb32 ";

        $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
        $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
        $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
        $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
        $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
        $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣

        $objPHPExcel->setActiveSheetIndex(0);
        $osheet = $objPHPExcel->getActiveSheet();

        $erp_sql2 = oci_parse($erp_conn1,$s2 );
        oci_execute($erp_sql2);
        $y=2;
        while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
          $osheet ->setCellValue('A'. $y, $row2["OMB31"])
                  ->setCellValue('B'. $y, $row2["OMB32"])
                  ->setCellValue('C'. $y, $row2["OMB04"])
                  ->setCellValue('D'. $y, $row2["OMB06"])
                  ->setCellValue('E'. $y, $row2["OMB05"])
                  ->setCellValue('F'. $y, $row2["OMB12"])
                  ->setCellValue('G'. $y, $row2["OMB13"])
                  ->setCellValue('H'. $y, $row2["OMB14"])
                  ->setCellValue('I'. $y, $row2["OMB15"])
                  ->setCellValue('J'. $y, $row2["OMB16"]);
          $y++;
        }

        // Rename sheet
        $osheet ->setTitle('axrt300');
        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $objPHPExcel->setActiveSheetIndex(0);
        
        // Redirect output to a client’s web browser (Excel5)
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="'. 'axrt300' . '_' . $bdate . '.xls"');
        header('Cache-Control: max-age=0');
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');

        $objWriter->save('php://output');
        exit;
  }

  $IsAjax = false;

  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>匯出關係人交易 </p>
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
         日期:
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()">
        <input type="submit" name="submit" id="submit" value="Excel">  &nbsp;&nbsp;   &nbsp;&nbsp;
      </td></tr>
    </table>
  </div>
</form>


