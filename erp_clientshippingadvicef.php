<?php
  session_start();
  $pagetitle = "業務部 &raquo; Shipping Advice Full Version";
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
    $bdate =  date('Y-m-d');
  } else {
    $bdate=$_GET['bdate'];
  }
  
  $month=substr($bdate,0,4).substr($bdate,5,2);

  $occfilter='';
  $datefilter='';

  $bocc01=$_GET['bocc01'];

  $occfilter  = " and oga04='$bocc01' ";
  
  $datefilter = " and oga02=to_date('$bdate','yy/mm/dd') ";

  if ($_GET["submit"]=="Excel")  {
  
    if ($bocc01=='U196001') {  //F 要取秤重的資料

      $filename="templates/shippingadvicef.xls";
      error_reporting(E_ALL);
      require_once 'classes/PHPExcel.php';
      require_once 'classes/PHPExcel/IOFactory.php';
      $objReader = PHPExcel_IOFactory::createReader('Excel5');
      $objPHPExcel = $objReader->load("$filename");

      $objPHPExcel ->getProperties()->setCreator( 'Shipping Advice')
        ->setLastModifiedBy('Shipping Advice')
        ->setTitle('Shipping Advice')
        ->setSubject('Shipping Advice')
        ->setDescription('Shipping Advice')
        ->setKeywords('Shipping Advice')
        ->setCategory('Shipping Advice');
      $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
      $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
      $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
      $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
      $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
  
      $objPHPExcel->setActiveSheetIndex(0);
      $osheet = $objPHPExcel->getActiveSheet();
  
      $socc="select occ02, occud08 from occ_file where occ01='$bocc01' ";
      $erp_sqlocc = oci_parse($erp_conn1,$socc );
      oci_execute($erp_sqlocc);
      $rowocc = oci_fetch_array($erp_sqlocc, OCI_ASSOC);
      $osheet ->setCellValue('C2', $rowocc['OCC02'])
        ->setCellValue('C5', $bdate);
  
      $occud08=$rowocc['OCCUD08']; //齒位別

      $s2="select tc_ogb011, ta_oea003, ta_oea001, ima1002, tc_ogb006, (ta_oea046||ta_oea047||ta_oea048) shade, oebud03  from " .
          "tc_oga_file, tc_ogb_file, ima_file, oea_file, oeb_file " .
          "where  tc_oga001=tc_ogb001 and tc_ogb003=oea01 and tc_ogb005=ima01 and tc_oga004='U196001' and tc_oga002=to_date('" . $bdate . "', 'yy/mm/dd' ) " .
          "and tc_ogb003 = oeb01 and tc_ogb004=oeb03 " .
          "order by tc_ogb011  ";

      $i=0;
      $y=8;
      $totalcase=0;
      $totalunit=0;
      $oldcaseno='';
      $caseno=='';
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
        $totalunit+=$row2["TC_OGB006"];
        if ($row2['TC_OGB011']!= $oldcaseno) {
          $i++;
          $oldcaseno= $row2['TC_OGB011'];
          $caseno   = $row2['TC_OGB011'];
          $patient  = $row2['TA_OEA003'];
          $clinic   = $row2['TA_OEA001'];
          $ii=$i;
        } else {            //重複的case no 不秀
          $caseno   = '';
          $patient  = '';
          $clinic   = '';
          $ii='';
        }

        $metal='';
        $tcdex004='';
        $lotno='';

        //要替換齒位規則
        $oebud03='';
        if ($row2["OEBUD03"]!=''){
          if ($occud08==2 or $occud08==3) {
            if (substr($row2["OEBUD03"],0,1)!='U' and substr($row2["OEBUD03"],0,1)!='L' ) {
              $oebud03='#'.changeteethno($row2["OEBUD03"],$occud08);
            } else {
              $oebud03=$row2['OEBUD03'];
            }
          } else {
            $oebud03=$row2['OEBUD03'];
          }
        }

        $osheet ->setCellValue('A'. $y, $ii)
          ->setCellValue('B'. $y, $caseno.' ')
          ->setCellValue('C'. $y, $patient)
          ->setCellValue('D'. $y, $clinic)
          ->setCellValue('E'. $y, $row2["IMA1002"] )
          ->setCellValue('F'. $y, '')
          ->setCellValue('G'. $y, '' )
          ->setCellValue('H'. $y, '')
          ->setCellValue('I'. $y, $row2["TC_OGB006"])
          ->setCellValue('J'. $y, '')
          ->setCellValue('K'. $y, $oebud03)
          ->setCellValue('L'. $y, $row2["SHADE"]);
        $osheet ->getStyle('B'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $osheet ->getStyle('G'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
        $osheet ->getStyle('K'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $osheet ->getStyle('L'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);


        if (($y%2)==0){
          $osheet->getStyle('A'.$y.':L'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');

        }
        $y++;
      }

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

      //F秀合計
      $osheet ->setCellValue('B'. $y, "TOTAL:")
        ->setCellValue('C'. $y, $i)
        ->setCellValue('D'. $y, "CASES")
        ->setCellValue('I'. $y, $totalunit)
        ->setCellValue('J'. $y, "UNITS");
      // Rename sheet
      $osheet ->setTitle('Shipping Advice Full Version');
      // Set active sheet index to the first sheet, so Excel opens this as the first sheet
      $objPHPExcel->setActiveSheetIndex(0);

      header('Content-Type: application/vnd.ms-excel');
      header('Content-Disposition: attachment;filename="shippingadvice_'. $bocc01 . '_'.$bdate.'.xls"');
      header('Cache-Control: max-age=0');
      $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
      $objWriter->save('php://output');
      exit;

    } else if ($bocc01=='E137001')   {
  
      $filename="templates/shippingadvicef_E137001.xls";
      error_reporting(E_ALL);
      require_once 'classes/PHPExcel.php';
      require_once 'classes/PHPExcel/IOFactory.php';
      $objReader = PHPExcel_IOFactory::createReader('Excel5');
      $objPHPExcel = $objReader->load("$filename");
  
      $objPHPExcel ->getProperties()->setCreator( 'Shipping Advice')
        ->setLastModifiedBy('Shipping Advice')
        ->setTitle('Shipping Advice')
        ->setSubject('Shipping Advice')
        ->setDescription('Shipping Advice')
        ->setKeywords('Shipping Advice')
        ->setCategory('Shipping Advice');
      $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
      $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
      $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
      $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
      $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
  
      $objPHPExcel->setActiveSheetIndex(0);
      $osheet = $objPHPExcel->getActiveSheet();
  
      $socc="select occ02, occud08 from occ_file where occ01='$bocc01' ";
      $erp_sqlocc = oci_parse($erp_conn1,$socc );
      oci_execute($erp_sqlocc);
      $rowocc = oci_fetch_array($erp_sqlocc, OCI_ASSOC);
      $osheet ->setCellValue('C2', $rowocc['OCC02'])
        ->setCellValue('C5', $bdate);
  
      $occud08=$rowocc['OCCUD08']; //齒位別
         
      $s2="select ogaud02, ta_oea003, ta_oea001, ima1002, metalno, metal, ima021, tc_dex004, ogb12, shade, oebud03  from " .
          "( select oga01, ogb03, ogaud02, ta_oea003, ta_oea001, ima1002, ogb12, (ta_oea046||ta_oea047||ta_oea048) shade, oebud03 " .
              "from oga_file, ogb_file, ima_file, oea_file, oeb_file where oga16=oea01 and ogb04=ima01 and oga01=ogb01 and ogb31=oeb01 and ogb32=oeb03  " . $datefilter . $occfilter . " ) " .
          "left join " .
          "(select tc_dex001, tc_dex002, tc_dex003 metalno, ima1002 metal, ima021, tc_dex004 from tc_dex_file, ima_file where tc_dex003=ima01 and imaud10 is not null ) " .
          "on oga01=tc_dex001 and ogb03=tc_dex002 " .
          "order by ogaud02  ";
      
      $i=0;
      $y=8;
      $totalcase=0;
      $totalunit=0;
      $oldcaseno='';
      $caseno=='';
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
        $totalunit+=$row2["OGB12"];
        if ($row2['OGAUD02']!= $oldcaseno) {
          $i++;
          $oldcaseno= $row2['OGAUD02'];
          $caseno   = $row2['OGAUD02'];
          $patient  = $row2['TA_OEA003'];
          $clinic   = $row2['TA_OEA001'];
          $ii=$i;
        } else {            //重複的case no 不秀
          $caseno   = '';
          $patient  = '';
          $clinic   = '';
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

        //要替換齒位規則
        $oebud03='';
        if ($row2["OEBUD03"]!=''){
          if ($occud08==2 or $occud08==3) {
            if (substr($row2["OEBUD03"],0,1)!='U' and substr($row2["OEBUD03"],0,1)!='L' ) {
              $oebud03='#'.changeteethno($row2["OEBUD03"],$occud08);
            } else {
              $oebud03=$row2['OEBUD03'];
            }
          } else {
            $oebud03=$row2['OEBUD03'];
          }
        }

        $osheet ->setCellValue('A'. $y, $ii)
                ->setCellValue('B'. $y, $caseno.' ')
                ->setCellValue('C'. $y, $row2["IMA1002"] )
                ->setCellValue('D'. $y, $metal)
                ->setCellValue('E'. $y, $row2["TC_DEX004"] )
                ->setCellValue('F'. $y, $lotno)
                ->setCellValue('G'. $y, $row2["OGB12"])
                ->setCellValue('H'. $y, '')
                ->setCellValue('I'. $y, $oebud03)
                ->setCellValue('J'. $y, $row2["SHADE"]);
        $osheet ->getStyle('B'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $osheet ->getStyle('E'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
        $osheet ->getStyle('I'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $osheet ->getStyle('J'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

        if (($y%2)==0){
          $osheet->getStyle('A'.$y.':J'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
        }
        $y++;
      }

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
      
      
      $osheet ->setCellValue('B'. $y, "TOTAL:")
              ->setCellValue('C'. $y, $i)
              ->setCellValue('D'. $y, "CASES")
              ->setCellValue('G'. $y, $totalunit)
              ->setCellValue('H'. $y, "UNITS");
      $y++;
      $osheet ->setCellValue('H'. $y, "Form No.:QF-S01-04 A");

        // Rename sheet
      $osheet ->setTitle('Shipping Advice Full Version');
      // Set active sheet index to the first sheet, so Excel opens this as the first sheet
      $objPHPExcel->setActiveSheetIndex(0);
    
      header('Content-Type: application/vnd.ms-excel');
      header('Content-Disposition: attachment;filename="shippingadvice_'. $bocc01 . '_'.$bdate.'.xls"');
      header('Cache-Control: max-age=0');
      $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
      $objWriter->save('php://output');
      exit;

    } else   {
  
      $filename="templates/shippingadvicef.xls";
      error_reporting(E_ALL);
      require_once 'classes/PHPExcel.php';
      require_once 'classes/PHPExcel/IOFactory.php';
      $objReader = PHPExcel_IOFactory::createReader('Excel5');
      $objPHPExcel = $objReader->load("$filename");
  
      $objPHPExcel ->getProperties()->setCreator( 'Shipping Advice')
        ->setLastModifiedBy('Shipping Advice')
        ->setTitle('Shipping Advice')
        ->setSubject('Shipping Advice')
        ->setDescription('Shipping Advice')
        ->setKeywords('Shipping Advice')
        ->setCategory('Shipping Advice');
      $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
      $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
      $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
      $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
      $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
  
      $objPHPExcel->setActiveSheetIndex(0);
      $osheet = $objPHPExcel->getActiveSheet();
  
      $socc="select occ02, occud08 from occ_file where occ01='$bocc01' ";
      $erp_sqlocc = oci_parse($erp_conn1,$socc );
      oci_execute($erp_sqlocc);
      $rowocc = oci_fetch_array($erp_sqlocc, OCI_ASSOC);
      $osheet ->setCellValue('C2', $rowocc['OCC02'])
        ->setCellValue('C5', $bdate);
  
      $occud08=$rowocc['OCCUD08']; //齒位別
         
      $s2="select ogaud02, ta_oea003, ta_oea001, ima1002, metalno, metal, ima021, tc_dex004, ogb12, shade, oebud03  from " .
          "( select oga01, ogb03, ogaud02, ta_oea003, ta_oea001, ima1002, ogb12, (ta_oea046||ta_oea047||ta_oea048) shade, oebud03 " .
              "from oga_file, ogb_file, ima_file, oea_file, oeb_file where oga16=oea01 and ogb04=ima01 and oga01=ogb01 and ogb31=oeb01 and ogb32=oeb03  " . $datefilter . $occfilter . " ) " .
          "left join " .
          "(select tc_dex001, tc_dex002, tc_dex003 metalno, ima1002 metal, ima021, tc_dex004 from tc_dex_file, ima_file where tc_dex003=ima01 and imaud10 is not null ) " .
          "on oga01=tc_dex001 and ogb03=tc_dex002 " .
          "order by ogaud02  ";
      
      $i=0;
      $y=8;
      $totalcase=0;
      $totalunit=0;
      $oldcaseno='';
      $caseno=='';
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
        $totalunit+=$row2["OGB12"];
        if ($row2['OGAUD02']!= $oldcaseno) {
          $i++;
          $oldcaseno= $row2['OGAUD02'];
          $caseno   = $row2['OGAUD02'];
          $patient  = $row2['TA_OEA003'];
          $clinic   = $row2['TA_OEA001'];
          $ii=$i;
        } else {            //重複的case no 不秀
          $caseno   = '';
          $patient  = '';
          $clinic   = '';
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

        //要替換齒位規則
        $oebud03='';
        if ($row2["OEBUD03"]!=''){
          if ($occud08==2 or $occud08==3) {
            if (substr($row2["OEBUD03"],0,1)!='U' and substr($row2["OEBUD03"],0,1)!='L' ) {
              $oebud03='#'.changeteethno($row2["OEBUD03"],$occud08);
            } else {
              $oebud03=$row2['OEBUD03'];
            }
          } else {
            $oebud03=$row2['OEBUD03'];
          }
        }

        $osheet ->setCellValue('A'. $y, $ii)
                ->setCellValue('B'. $y, $caseno.' ')
                ->setCellValue('C'. $y, $patient)
                ->setCellValue('D'. $y, $clinic)
                ->setCellValue('E'. $y, $row2["IMA1002"] )
                ->setCellValue('F'. $y, $metal)
                ->setCellValue('G'. $y, $row2["TC_DEX004"] )
                ->setCellValue('H'. $y, $lotno)
                ->setCellValue('I'. $y, $row2["OGB12"])
                ->setCellValue('J'. $y, '')
                ->setCellValue('K'. $y, $oebud03)
                ->setCellValue('L'. $y, $row2["SHADE"]);
        $osheet ->getStyle('B'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        $osheet ->getStyle('G'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
        $osheet ->getStyle('K'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $osheet ->getStyle('L'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

        if (($y%2)==0){
          $osheet->getStyle('A'.$y.':L'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
        }
        $y++;
      }

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
      
      
      $osheet ->setCellValue('B'. $y, "TOTAL:")
              ->setCellValue('C'. $y, $i)
              ->setCellValue('D'. $y, "CASES")
              ->setCellValue('I'. $y, $totalunit)
              ->setCellValue('J'. $y, "UNITS");
       $y++;
      $osheet ->setCellValue('J'. $y, "Form No.:QF-S01-04 A");
      // Rename sheet
      $osheet ->setTitle('Shipping Advice Full Version');
      // Set active sheet index to the first sheet, so Excel opens this as the first sheet
      $objPHPExcel->setActiveSheetIndex(0);
    
      header('Content-Type: application/vnd.ms-excel');
      header('Content-Disposition: attachment;filename="shippingadvice_'. $bocc01 . '_'.$bdate.'.xls"');
      header('Cache-Control: max-age=0');
      $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
      $objWriter->save('php://output');
      exit;
    }
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
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()">
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
        <input type="submit" name="submit" id="submit" value="Excel">  &nbsp;&nbsp;   &nbsp; &nbsp; &nbsp;
      </td></tr>
    </table>
  </div>
</form>
<? if (is_null($_GET['submit'])) die ; ?>
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
      <th width="16">&nbsp;</th>
      <th>RX No</th>
      <th>Patient</th>
      <th>Clinic</th>
      <th>Product</th>
      <th>Alloy(Composition)</th>
      <th>Weight</th>
      <th>Lot No.</th>
      <th>Unit(s)</th>
      <th>Account</th>
    </tr>
    <?   
      $soga="select ogaud02, ta_oea003, ta_oea001, ima1002, metalno, metal, ima021, tc_dex004, ogb12  from " . 
            "( select oga01, ogb03, ogaud02, ta_oea003, ta_oea001, ima1002, ogb12 " .
               "from oga_file, ogb_file, ima_file, oea_file where oga16=oea01 and ogb04=ima01 and oga01=ogb01 " . $datefilter . $occfilter . " ) " .
            "left join " .
            "(select tc_dex001, tc_dex002, tc_dex003 metalno, ima1002 metal, ima021, tc_dex004 from tc_dex_file, ima_file where tc_dex003=ima01 and imaud10 is not null ) " .    
            "on oga01=tc_dex001 and ogb03=tc_dex002 " .
            "order by ogaud02  ";        

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
            $ii=$i;
        } else {            //重複的case no 不秀
            $caseno   = '';
            $patient  = '';
            $clinic   = '';
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

    ?>
        <tr bgcolor="#<?=$bgkleur;?>">
            <td><?=$ii;?></td>
            <td><?=$caseno;?></td>
            <td><?=$patient;?></td>
            <td><?=$clinic;?></td>
            <td><?=$rowoga['IMA1002'];?></td>
            <td><?=$metal ;?></td>
            <td style="text-align:right" ><?=$tcdex004;?></td>
            <td><?=$lotno;?></td>
            <td style="text-align:right" ><?=$rowoga["OGB12"];?></td>
            <td>&nbsp;</td>
        </tr>
    <?
      }
    ?>
</table>
