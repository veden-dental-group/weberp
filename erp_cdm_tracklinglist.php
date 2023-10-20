<?php
  session_start();
  $pagetitle = "業務部 &raquo; CDM Tracking List";
  include("_data.php");
  include("_erp.php");
  //auth("erp_clientshippingadvicef.php");
  function findlotno($erp_conn1,$metalno, $month) {
    $stcfrc="select tc_frc003 from tc_frc_file where tc_frc001='$metalno' and '$month'>=tc_frc002 order by tc_frc002 desc  ";
      $erp_sqltcfrc = oci_parse($erp_conn1,$stcfrc );
      oci_execute($erp_sqltcfrc);
      $rowtcfrc = oci_fetch_array($erp_sqltcfrc, OCI_ASSOC);
      if(is_null($rowtcfrc["TC_FRC003"])) {
        return ('');
      } else {
        return($rowtcfrc["TC_FRC003"]);
      }
  }

    function findlotnov2($erp_conn1,$metalno, $bdate) {
      $stcfrc="select tc_img003 from tc_img_file where tc_img002='$metalno' and tc_img001 <= to_date('$bdate','yy/mm/dd') order by tc_img001 desc  ";
      $erp_sqltcfrc = oci_parse($erp_conn1,$stcfrc );
      oci_execute($erp_sqltcfrc);
      $rowtcfrc = oci_fetch_array($erp_sqltcfrc, OCI_ASSOC);
      if(is_null($rowtcfrc["TC_IMG003"])) {
        return ('');
      } else {
        return($rowtcfrc["TC_IMG003"].' ');
      }
  }


  if (is_null($_GET['bdate'])) {
    $bdate = date('Y-m-d');
  } else {
    $bdate=$_GET['bdate'];
  }

if (is_null($_GET['edate'])) {
    $edate = date('Y-m-d');
} else {
    $edate = $_GET['edate'];
}

  $occfilter='';
  $datefilter='';

  $bocc01=$_GET['bocc01'];
  $occfilter  = " and oga04='$bocc01' ";

  $datefilter = " and oga02>=to_date('$bdate','yy/mm/dd') and oga02<=to_date('$edate','yy/mm/dd') ";

  if (($_GET["submit"]=="Excel") or  ($_GET["submit"]=="PDF")) {
      //取出客戶的shipping advice格式設定
      $filename="templates/CDMCaseTracking.xls";

      error_reporting(E_ALL);

      require_once 'classes/PHPExcel.php';
      require_once 'classes/PHPExcel/IOFactory.php';
      $objReader = PHPExcel_IOFactory::createReader('Excel5');
      $objPHPExcel = $objReader->load("$filename");

      $objPHPExcel ->getProperties()->setCreator( 'tracking list')
                   ->setLastModifiedBy('tracking list')
                   ->setTitle('tracking list')
                   ->setSubject('tracking list')
                   ->setDescription('tracking list')
                   ->setKeywords('tracking list')
                   ->setCategory('tracking list');
      $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
      $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
      $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
      $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
      $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);


          $objPHPExcel->setActiveSheetIndex(0);
          $osheet = $objPHPExcel->getActiveSheet();

          $socc="select occ02, occ18 from occ_file where occ01='$bocc01' ";
          $erp_sqlocc = oci_parse($erp_conn1,$socc );
          oci_execute($erp_sqlocc);
          $rowocc = oci_fetch_array($erp_sqlocc, OCI_ASSOC);
          $osheet ->setCellValue('A1', 'TO: ' . $rowocc['OCC18']);
          $osheet ->setCellValue('A2', 'VD tracking list from ' . $bdate . ' -- ' . $edate);

          $s2="select to_char(oea02, 'mm/dd/yyyy') oea02, to_char(oga02,'mm/dd/yyyy') oga02, ogaud02, ta_oea003, ta_oea001, ima1002, metalno, metal, ima021, tc_dex004, ogb12, shade, oebud03, ogb04  from " .
                "( select oea02, oga02, oga01, ogb03, ogaud02, ta_oea003, ta_oea001, ima1002, ogb12, (ta_oea046||ta_oea047||ta_oea048) shade, oebud03, ogb04 " .
                   "from oga_file, ogb_file, ima_file, oea_file, oeb_file where oga16=oea01 and ogb04=ima01 and oga01=ogb01 and ogb31=oeb01 and ogb32=oeb03  " . $datefilter . $occfilter . " ) " .
                "left join " .
                "(select tc_dex001, tc_dex002, tc_dex003 metalno, ima1002 metal, ima021, tc_dex004 from tc_dex_file, ima_file where tc_dex003=ima01 and imaud10 is not null ) " .
                "on oga01=tc_dex001 and ogb03=tc_dex002 " .
                "order by oga02, upper(ta_oea001), ogaud02  ";
          $i=0;
          $y=4;
          $totalcase=0;
          $totalunit=0;
          $oldcaseno='';
          $oldoga02='';
          $caseno=='';
          $erp_sql2 = oci_parse($erp_conn1,$s2 );
          oci_execute($erp_sql2);
          while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
          	if ($row2['OGA02']!=$oldoga02 && $y!=4 ){
          		$y++;
          	}

          	$oldoga02 = $row2['OGA02'];
            $totalunit+=$row2["OGB12"];
            if ($row2['OGAUD02']!= $oldcaseno) {
                  $i++;
                  $oldcaseno= $row2['OGAUD02'];
                  $caseno   = $row2['OGAUD02'];
                  $patient  = $row2['TA_OEA003'];
                  $clinic   = $row2['TA_OEA001'];
                  $oga02    = $row2['OGA02'];
                  $oea02    = $row2['OEA02'];
                  $ii=$i;
              } else {            //重複的case no 不秀
                  $caseno   = '';
                  $patient  = '';
                  $clinic   = '';
                  $oga02    = '';
                  $oea02    = '';
                  $ii='';
              }
              if (is_null($row2['METAL'])) {      //沒有金屬重量不秀
                  $metal='';
                  $tcdex004='';
                  $lotno='';
              } else {
                  #金屬資料改由TC_OGI_FILE取出
                  $sm="select tc_ogi002 from tc_ogi_file where tc_ogi001='" . $row2['METALNO'] ."' ";
                  $erp_sqlm = oci_parse($erp_conn1,$sm );
                  oci_execute($erp_sqlm);
                  $rowm = oci_fetch_array($erp_sqlm, OCI_ASSOC);
                  $metal= $row2["METAL"] . "(" . $rowm["TC_OGI002"] . ")";
                  $tcdex004=number_format($row2["TC_DEX004"],2,'.',',');
                  $lotno=findlotnov2($erp_conn1,$row2["METALNO"], $bdate);
              }

            //24111, 24112, 24113, 2411A Alloy都秀GM800+
            if ($row2['OGB04']=='24111' || $row2['OGB04']=='24112' || $row2['OGB04']=='24113' || $row2['OGB04']=='2411A') {
            	$metal = 'GM800+';
            }

            $osheet ->setCellValue('A'. $y, $ii)
                    ->setCellValue('B'. $y, $oea02 )
                    ->setCellValue('C'. $y, $oga02 )
                    ->setCellValue('D'. $y, $caseno.' ')
                    ->setCellValue('E'. $y, $patient)
                    ->setCellValue('F'. $y, $clinic)
                    ->setCellValue('G'. $y, $row2["IMA1002"] )
                    ->setCellValue('H'. $y, $metal)
                    ->setCellValue('I'. $y, $row2["TC_DEX004"] )
                    ->setCellValue('J'. $y, $lotno)
                    ->setCellValue('K'. $y, $row2["OGB12"])
                    ->setCellValue('L'. $y, $row2["OEBUD03"]);
            //$osheet ->getStyle('B'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            //$osheet ->getStyle('K'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );

            if (($y%2)==0){
               $osheet->getStyle('A'.$y.':L'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
            }
            /*
            $osheet ->getStyle('A'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('B'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('C'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('D'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('E'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('F'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('G'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('H'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('I'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('J'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('K'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('L'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

            $osheet ->getStyle('A'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('B'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('C'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('D'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('E'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('F'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('G'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('H'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('I'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('J'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('K'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('L'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

            $osheet ->getStyle('A'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('B'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('C'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('D'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('E'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('F'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('G'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('H'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('I'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('J'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('K'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('L'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

            $osheet ->getStyle('A'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('B'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('C'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('D'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('E'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('F'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('G'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('H'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('I'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('J'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('K'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('L'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            */
            $y++;
          }

/*          $osheet ->getStyle('A'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          $osheet ->getStyle('B'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          $osheet ->getStyle('C'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          $osheet ->getStyle('D'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          $osheet ->getStyle('E'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          $osheet ->getStyle('F'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          $osheet ->getStyle('G'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          $osheet ->getStyle('H'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          $osheet ->getStyle('I'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          $osheet ->getStyle('J'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          $osheet ->getStyle('K'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          $osheet ->getStyle('L'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          $osheet ->setCellValue('D'. $y, "TOTAL:")
                  ->setCellValue('E'. $y, $i)
                  ->setCellValue('F'. $y, "CASES")
                  ->setCellValue('K'. $y, $totalunit)
                  ->setCellValue('L'. $y, "UNITS");*/
          // Rename sheet

          $objPHPExcel->setActiveSheetIndex(0);

      header('Content-Type: application/vnd.ms-excel');
      header('Content-Disposition: attachment;filename="trackinglist_'. $bocc01 . '_'.$bdate.'--'.$edate.'.xls"');
      header('Cache-Control: max-age=0');
      $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
      $objWriter->save('php://output');
      exit;
  }

    //for xajax
  require ('xajax/xajax_core/xajax.inc.php');
  $xajax = new xajax();
  $xajax->configure('javascript URI', 'xajax/');

  $xajax->processRequest();

  echo '<?xml version="1.0" encoding="UTF-8"?>';
  $IsAjax = True;
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>客戶 Shipping Advice Full 列印 </p>
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
         出貨日期:
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> ~~
        <input name="edate" type="text" id="edate" size="12" maxlength="12" value=<?=$edate;?> onfocus="new WdatePicker()"> ~~
        客戶:
        <select name="bocc01" id="bocc01">
            <?
              $s1= "select occ01,occ02 from occ_file order by occ01 ";
              $erp_sql1 = oci_parse($erp_conn2,$s1 );
              oci_execute($erp_sql1);
              while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                  echo "<option value=" . $row1["OCC01"];
                  if ($_GET["bocc01"] == $row1["OCC01"]) echo " selected";
                  echo ">" . $row1['OCC01'] ."--" .$row1["OCC02"] . "</option>";
              }
            ?>
        </select>
        <input type="submit" name="submit" id="submit" value="Submit">  &nbsp;&nbsp;   &nbsp;&nbsp;
        <input type="submit" name="submit" id="submit" value="Excel">  &nbsp;&nbsp;   &nbsp; &nbsp;
      </td></tr>
    </table>
  </div>
</form>
<? if (is_null($_GET['submit'])) die ; ?>
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
        <th width="16">&nbsp;</th>
        <th>Receiving Date</th>
        <th>Shipping Date</th>
        <th>RX No</th>
        <th>Patient</th>
        <th>Clinic</th>
        <th>Product</th>
        <th>Alloy(Composition)</th>
        <th>Weight</th>
        <th>Lot No.</th>
        <th>Unit(s)</th>
        <th>Teeth</th>
    </tr>
    <?
      $soga="select to_char(oea02, 'mm/dd/yyyy') oea02, to_char(oga02,'mm/dd/yyyy') oga02, ogaud02, ta_oea003, ta_oea001, ima1002, metalno, metal, ima021, tc_dex004, ogb12, shade, oebud03, ogb04  from " .
                "( select oea02, oga02, oga01, ogb03, ogaud02, ta_oea003, ta_oea001, ima1002, ogb12, (ta_oea046||ta_oea047||ta_oea048) shade, oebud03, ogb04 " .
                   "from oga_file, ogb_file, ima_file, oea_file, oeb_file where oga16=oea01 and ogb04=ima01 and oga01=ogb01 and ogb31=oeb01 and ogb32=oeb03  " . $datefilter . $occfilter . " ) " .
                "left join " .
                "(select tc_dex001, tc_dex002, tc_dex003 metalno, ima1002 metal, ima021, tc_dex004 from tc_dex_file, ima_file where tc_dex003=ima01 and imaud10 is not null ) " .
                "on oga01=tc_dex001 and ogb03=tc_dex002 " .
                "order by oga02, upper(ta_oea001), ogaud02  ";
      $erp_sqloga = oci_parse($erp_conn1,$soga );
      oci_execute($erp_sqloga);
      $bgkleur = "ffffff";
      $i=0;
      $oldcaseno='';
      $caseno='';
      $patient='';
      while ($rowoga = oci_fetch_array($erp_sqloga, OCI_ASSOC)) {
          if ($rowoga['OGAUD02']!= $oldcaseno) {
              $i++;
              $oldcaseno= $rowoga['OGAUD02'];
              $caseno   = $rowoga['OGAUD02'];
              $patient  = $rowoga['TA_OEA003'];
              $clinic   = $rowoga['TA_OEA001'];
              $oga02    = $rowoga['OGA02'];
              $oea02    = $rowoga['OEA02'];
              $ii=$i;
          } else {            //重複的case no 不秀
              $caseno   = '';
              $patient  = '';
              $clinic   = '';
              $oea02 	= '';
              $oga02 	= '';
              $ii='';
          }
          if (is_null($rowoga['METAL'])) {      //沒有金屬重量不秀
              $metal='';
              $tcdex004='';
              $lotno='';
          } else {
              #金屬資料改由TC_OGI_FILE取出
              $sm="select tc_ogi002 from tc_ogi_file where tc_ogi001='" . $rowoga['METALNO'] ."' ";
              $erp_sqlm = oci_parse($erp_conn1,$sm );
              oci_execute($erp_sqlm);
              $rowm = oci_fetch_array($erp_sqlm, OCI_ASSOC);
              $metal= $rowoga["METAL"] . "(" . $rowm["TC_OGI002"] . ")";
              #$metal= $rowoga["METAL"] . "(" . $rowoga["IMA021"] . ")";
              $tcdex004=number_format($rowoga["TC_DEX004"],2,'.',',');
              $lotno=findlotnov2($erp_conn1,$rowoga["METALNO"], $bdate);
          }

          //24111, 24112, 24113, 2411A Alloy都秀GM800+
            if ($rowoga['OGB04']=='24111' || $rowoga['OGB04']=='24112' || $rowoga['OGB04']=='24113' || $rowoga['OGB04']=='2411A') {
            	$metal = 'GM800+';
            }

    ?>
          <tr bgcolor="#<?=$bgkleur;?>">
              <td><?=$ii;?></td>
              <td><?=$oea02;?></td>
              <td><?=$oga02;?></td>
              <td><?=$caseno;?></td>
              <td><?=$patient;?></td>
              <td><?=$clinic;?></td>
              <td><?=$rowoga['IMA1002'];?></td>
              <td><?=$metal ;?></td>
              <td style="text-align:right" ><?=$tcdex004;?></td>
              <td><?=$lotno;?></td>
              <td style="text-align:right" ><?=$rowoga["OGB12"];?></td>
              <td><?=$rowoga['OEBUD03'];?></td>
          </tr>
    <?
      }
    ?>
</table>
