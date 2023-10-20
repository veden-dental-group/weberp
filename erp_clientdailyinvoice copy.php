<?php
  session_start();
  $pagetitle = "帳單組 &raquo; 每日invoice";
  include("_data.php");
  include("_erp.php");
  //auth("erp_clientdailyinvoice.php");

  if (is_null($_GET['bdate'])) {
    $bdate =  date('Y-m-d');
  } else {
    $bdate=$_GET['bdate'];
  }

  if (is_null($_GET['occ01'])) {
    $occ01 =  '';
  } else {
    $occ01=$_GET['occ01'];
  }

  $invoiceno=$_GET['invoiceno'];

  $invoicefilter='';
  if ($invoiceno!='')  $invoicefilter=" and tc_ofa01='$invoiceno' ";

  if (($_GET["submit"]=="Excel") or  ($_GET["submit"]=="PDF")) {
        //根據 occud04, occud05, occud06 取出要使用的template檔名
        $s2= "select occ02, occ42, occud04, occud05, occud06, occud08 from occ_file where occ01='$occ01'";
        $erp_sql2 = oci_parse($erp_conn2,$s2 );
        oci_execute($erp_sql2);
        $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);
        if (is_null($row2['OCCUD04'])) {
            $occud04='N';
        } else {
            $occud04=$row2['OCCUD04'];
        }
        if (is_null($row2['OCCUD05'])) {
            $occud05='N';
        } else {
            $occud05=$row2['OCCUD05'];
        }
        if (is_null($row2['OCCUD06'])) {
            $occud06='N';
        } else {
            $occud06=$row2['OCCUD06'];
        }

        $occ02 = $row2['OCC02'];
        $occud08=$row2['OCCUD08'];

        if ($occ01=='U144002') {  //U144002, UC-R 要加 pan No. 放在doctor 欄位裡
          $template='U144002';
        } else if ($occ01=='A114001') {
           $template='A114001';     // A114001 VL-A
        } else if ($occ01=='A114002') {
           $template='A114001';     // A114001 VL-A
        } else if ($occ01=='E181002') {
           $template='E181002';     // CDM
        } else if ($occ01=='E181003') {
           $template='E181003';     // CDM-C
        } else if ($occ01=='E181004') {
           $template='E181004';     // CDM-M
        } else if ($occ01=='E185001') {
            $template='E185001';     // UKDD
        } else if ($occ01=='T185003') {
            $template='T185003';     // TWDEW
        } else if ($occ01=='J11700') {
            $template='JD';     // JANA
        } else if ($occ01=='E204001') {
            $template='JD';     // GL
        } else if ($occ01=='U190001') {
            $template='JD';     // B
        } else if ($occ01>='U200000' && $occ01<='U200025') {
            $template='TA';     // JD
        } else if ($occ01>='U192000' && $occ01<='U192010') {
            $template='TA';     // JD
        } else if ($occ01=='U191001') {
            $template='TA';     // JD
        } else if ($occ01=='U201001') {
            $template='TA';     // JD
        } else if ($occ01=='U204001') {
            $template='U204001';     // 003
        } else {
          $template= $occud04.$occud05.$occud06;
        }
        if ($occ01=='U196001'){
            $template='F';
        }
        //HP集團的
        if (substr($occ01,0,4)=='U198'){
            $template='HP';
        }

        if ($occ01=="E201002") {
            $occtitle = '';
        } else {
            $occtitle = " (" . $occ02 . ")";
        }

        $filename="templates/invoice_$template.xls";
        $currency=$row2['OCC42'];
        error_reporting(E_ALL);

        require_once 'classes/PHPExcel.php';
        require_once 'classes/PHPExcel/IOFactory.php';
        $objReader = PHPExcel_IOFactory::createReader('Excel5');
        $objPHPExcel = $objReader->load("$filename");
        //$objPHPExcel = new PHPExcel();
        // Set properties

        $objPHPExcel ->getProperties()->setCreator( 'Invoice')
                     ->setLastModifiedBy('Invoice')
                     ->setTitle('Invoice')
                     ->setSubject('Invoice')
                     ->setDescription('Invoice')
                     ->setKeywords('Invoice')
                     ->setCategory('Invoice');

        $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);
        $s2="select tc_ofa01, tc_ofa23, tc_ofaud03, tc_ofb11, tc_ofbud05, tc_ofb08, tc_ofb04, tc_ofb06, tc_ofb12, tc_ofb05, tc_ofb13, tc_ofb14, tc_ofb31, tc_ofbud02, tc_ofbud01, tc_ofb31, imaud04, ima02 " .
            "from tc_ofa_file, tc_ofb_file, ima_file " .
            "where tc_ofaconf='Y' and  tc_ofa02=to_date('$bdate1','yy/mm/dd') and tc_ofa04='$occ01' and tc_ofa01=tc_ofb01 and tc_ofb04=ima01  $invoicefilter ".
            "order by tc_ofa01,tc_ofb11,tc_ofb08,tc_ofb04 ";


        if ($template=='U144002')
        {    //UC-R 要秀Pan No.

                $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
                $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣

                $objPHPExcel->setActiveSheetIndex(0);
                $osheet = $objPHPExcel->getActiveSheet();

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);

                $y=5;
                $osheet -> setCellValue('A'.$y, "INVOICE NO.:".$row2["TC_OFA01"] . $occtitle)
                        -> setCellValue('E'.$y, "INVOICE DATE:".$bdate);

                $s3="select occ18, occ231, occ232, occ261, occ271 from occ_file where occ01='$occ01'";
                $erp_sql3 = oci_parse($erp_conn2,$s3 );
                oci_execute($erp_sql3);
                $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);
                $tel='';
                if ($row3['occ261']!='') $tel="Tel: " . $row3['occ261'] . "    ";
                if ($row3['occ271']!='') $tel.="Fax:" . $row3['occ271'];

                $y=6;
                $osheet -> setCellValue('A'.$y, $row3["OCC18"]);
                $y=7;
                $osheet -> setCellValue('A'.$y, $row3["OCC231"]);
                $y=8;
                $osheet -> setCellValue('A'.$y, $row3["OCC232"]);
                $y=9;
                $osheet -> setCellValue('A'.$y, $tel);

                $y=12;
                //$currency=$row2["TC_OFA23"];
                $i=0;
                $oldrxno='';
                $total=0;

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
                  $total+=$row2["TC_OFB14"];
                  if ( $row2["TC_OFB11"]!=$oldrxno) {
                    $i++;
                    $ii=$i;
                    $rxno =$row2['TC_OFB11'];
                    $oldrxno=$row2['TC_OFB11'];
                  } else {
                    $rxno=$row2['TC_OFB11'];
                    $ii='';
                  }


                  if ($row2['TC_OFB05']=="G") {
                      $unit="G";
                  } else  {
                      $unit="Unit";
                      $unittotal+=$row2["TC_OFB12"];
                  }

                  $tcofbud02='';
                  if ($row2["TC_OFBUD02"]!=''){
                      if ($occud08==2 or $occud08==3) {
                          if (substr($row2["TC_OFBUD02"],0,1)!='U' and substr($row2["TC_OFBUD02"],0,1)!='L' ) {
                              $tcofbud02='#'.changeteethno($row2["TC_OFBUD02"],$occud08);
                          } else {
                              $tcofbud02=$row2['TC_OFBUD02'];
                          }
                      } else {
                          $tcofbud02=$row2['TC_OFBUD02'];
                      }
                  }

                  //取出Pan No. 放在 訂單的doctor 欄位中
                  $oga01=$row2['TC_OFB31'];
                  $soea="select ta_oea002 from oea_file,oga_file where oea01=oga16 and oga01='$oga01'";
                  $erp_sqloea = oci_parse($erp_conn2,$soea );
                  oci_execute($erp_sqloea);
                  $rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC);

                  $osheet ->setCellValue('A'. $y, $ii)
                          ->setCellValue('B'. $y, $rxno . ' ')
                          ->setCellValue('C'. $y, $rowoea["TA_OEA002"])
                          ->setCellValue('D'. $y, $row2["TC_OFB06"])
                          ->setCellValue('E'. $y, $row2["TC_OFB12"])
                          ->setCellValue('F'. $y, $unit)
                          ->setCellValue('G'. $y, number_format($row2["TC_OFB13"],2,'.',','))
                          ->setCellValue('H'. $y, number_format($row2["TC_OFB14"],2,'.',','))
                          ->setCellValue('I'. $y, $tcofbud02)
                          ->setCellValue('J'. $y, $row2["TC_OFBUD01"]);



                  $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                  $osheet ->getStyle('G'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                  $osheet ->getStyle('H'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                  $osheet ->getStyle('I'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                  $osheet ->getStyle('J'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                  if (($y%2)==0){
                      $osheet->getStyle('A'.$y.':J'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                  }
                  $y++;
                }
                $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('H'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('I'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('J'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                //total
                $osheet ->setCellValue('A'.$y, 'Total')
                        ->setCellValue('E'.$y, $unittotal)
                        ->setCellValue('F'.$y, 'Units')
                        ->setCellValue('G'.$y, $currency)
                        ->setCellValue('H'.$y, number_format($total,2,'.',','));
                $osheet ->getStyle('E'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('F'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                $osheet ->getStyle('G'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('H'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('H'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );

                $osheet ->getStyle('A'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('B'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('C'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('D'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('F'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('G'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('H'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('I'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('J'.$y) ->getFont()->setBold(true);

                // Rename sheet
                $osheet ->setTitle('DailyInvoice');
                // Set active sheet index to the first sheet, so Excel opens this as the first sheet
                $objPHPExcel->setActiveSheetIndex(0);

        }
        else if($template=='A114001')
        {   //VL-A

                $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
                $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣

                $objPHPExcel->setActiveSheetIndex(0);
                $osheet = $objPHPExcel->getActiveSheet();

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);

                $y=5;
                $osheet -> setCellValue('A'.$y, "INVOICE NO.:".$row2["TC_OFA01"]. " (" . $occ02 . ")")
                        -> setCellValue('F'.$y, "INVOICE DATE:".$bdate);

                $s3="select occ18, occ231, occ232, occ261, occ271 from occ_file where occ01='$occ01'";
                $erp_sql3 = oci_parse($erp_conn2,$s3 );
                oci_execute($erp_sql3);
                $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);
                $tel='';
                if ($row3['occ261']!='') $tel="Tel: " . $row3['occ261'] . "    ";
                if ($row3['occ271']!='') $tel.="Fax:" . $row3['occ271'];

                $y=6;
                $osheet -> setCellValue('A'.$y, $row3["OCC18"]);
                $y=7;
                $osheet -> setCellValue('A'.$y, $row3["OCC231"]);
                $y=8;
                $osheet -> setCellValue('A'.$y, $row3["OCC232"]);
                $y=9;
                $osheet -> setCellValue('A'.$y, $tel);

                $y=12;
                //$currency=$row2["TC_OFA23"];
                $i=0;
                $oldrxno='';
                $total=0;

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
                  $total+=$row2["TC_OFB14"];
                  if ( $row2["TC_OFB11"]!=$oldrxno) {
                    $i++;
                    $ii=$i;
                    $rxno =$row2['TC_OFB11'];
                    $oldrxno=$row2['TC_OFB11'];
                  } else {
                    $rxno=$row2['TC_OFB11'];
                    $ii='';
                  }


                  if ($row2['TC_OFB05']=="G") {
                      $unit="G";
                  } else  {
                      $unit="Unit";
                      $unittotal+=$row2["TC_OFB12"];
                  }

                  //invoice目前沒有抛轉patient過來 先改由訂單去抓patient
                  $tc_ofbud05=$row2['TC_OFBUD05'];
                  $tc_ofb31=$row2['TC_OFB31'];
                  $soea="select ta_oea002, ta_oea003 from oea_file, oga_file where oea01=oga16 and oga01='$tc_ofb31' ";
                  $erp_sqloea = oci_parse($erp_conn2,$soea );
                  oci_execute($erp_sqloea);
                  $rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC) ;
                  //$tc_ofbud05=$rowoea['TA_OEA003'];

                  // Case No. :           $rxno
                  // Patient Name:        TC_OFBUD05
                  // Product Description: TC_OFB06)
                  // Qty                  TC_OFB12
                  // Unit Price           TC_OFB13
                  // Total                TC_OFB14
                  // Tooth Number         TC_OFBUD02
                  // memo                 TC_OFBUD01
                  $tcofbud02='';
                  if ($row2["TC_OFBUD02"]!=''){
                      if ($occud08==2 or $occud08==3) {
                          if (substr($row2["TC_OFBUD02"],0,1)!='U' and substr($row2["TC_OFBUD02"],0,1)!='L' ) {
                              $tcofbud02='#'.changeteethno($row2["TC_OFBUD02"],$occud08);
                          } else {
                              $tcofbud02=$row2['TC_OFBUD02'];
                          }
                      } else {
                          $tcofbud02=$row2['TC_OFBUD02'];
                      }
                  }

                  $osheet ->setCellValue('A'. $y, $ii)
                          ->setCellValue('B'. $y, $rxno . ' ')
                          ->setCellValue('C'. $y, $rowoea['TA_OEA002'])
                          ->setCellValue('D'. $y, $tc_ofbud05)
                          ->setCellValue('E'. $y, $row2["TC_OFB06"])
                          ->setCellValue('F'. $y, $row2["TC_OFB12"])
                          ->setCellValue('G'. $y, $unit)
                          ->setCellValue('H'. $y, number_format($row2["TC_OFB13"],2,'.',','))
                          ->setCellValue('I'. $y, number_format($row2["TC_OFB14"],2,'.',','))
                          ->setCellValue('J'. $y, $tcofbud02)
                          ->setCellValue('K'. $y, $row2["TC_OFBUD01"]);



                  $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                  $osheet ->getStyle('J'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                  $osheet ->getStyle('K'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                  if (($y%2)==0){
                      $osheet->getStyle('A'.$y.':K'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                  }
                  $y++;
                }
                $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('H'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('I'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('J'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('K'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                //total
                $osheet ->setCellValue('A'.$y, 'Total')
                        ->setCellValue('F'.$y, $unittotal)
                        ->setCellValue('G'.$y, 'Units')
                        ->setCellValue('H'.$y, $currency)
                        ->setCellValue('I'.$y, number_format($total,2,'.',','));
                $osheet ->getStyle('G'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                $osheet ->getStyle('H'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('I'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('I'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );

                $osheet ->getStyle('A'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('B'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('C'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('D'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('F'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('G'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('H'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('I'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('J'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('K'.$y) ->getFont()->setBold(true);

                // Rename sheet
                $osheet ->setTitle('DailyInvoice');
                // Set active sheet index to the first sheet, so Excel opens this as the first sheet
                $objPHPExcel->setActiveSheetIndex(0);

        }
        else if($template=='E181002')
        {  // CDM
                $s2="select tc_ofa01, tc_ofa23, tc_ofb11, tc_ofbud05, tc_ofb08, tc_ofb04, tc_ofb06, tc_ofb12, tc_ofb05, tc_ofb13, tc_ofb14, tc_ofb31, tc_ofbud02, tc_ofbud01, tc_ofb31, imaud04, ta_oea001 " .
                    "from tc_ofa_file, tc_ofb_file, ima_file, oga_file, oea_file " .
                    "where tc_ofaconf='Y' and upper(substr(trim(ta_oea001),1,1)) in ('A','B','C') and tc_ofb31=oga01 and oga16=oea01 and tc_ofa02=to_date('$bdate1','yy/mm/dd') and tc_ofa04='$occ01' and tc_ofa01=tc_ofb01 and tc_ofb04=ima01  $invoicefilter ".
                    "order by tc_ofa01, upper(trim(ta_oea001)), tc_ofb11,tc_ofb08,tc_ofb04 ";

                $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
                $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣

                $objPHPExcel->setActiveSheetIndex(0);
                $osheet = $objPHPExcel->getActiveSheet();

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);

                $y=5;
                $osheet -> setCellValue('A'.$y, "INVOICE NO.:".$row2["TC_OFA01"]. " (" . $occ02 . ")")
                        -> setCellValue('F'.$y, "INVOICE DATE:".$bdate);

                $s3="select occ18, occ231, occ232, occ261, occ271 from occ_file where occ01='$occ01'";
                $erp_sql3 = oci_parse($erp_conn2,$s3 );
                oci_execute($erp_sql3);
                $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);
                $tel='';
                if ($row3['occ261']!='') $tel="Tel: " . $row3['occ261'] . "    ";
                if ($row3['occ271']!='') $tel.="Fax:" . $row3['occ271'];

                $y=6;
                $osheet -> setCellValue('A'.$y, $row3["OCC18"]);
                $y=7;
                $osheet -> setCellValue('A'.$y, $row3["OCC231"]);
                $y=8;
                $osheet -> setCellValue('A'.$y, $row3["OCC232"]);
                $y=9;
                $osheet -> setCellValue('A'.$y, $tel);

                $y=12;
                //$currency=$row2["TC_OFA23"];
                $i=0;
                $oldrxno='';
                $total=0;

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
                  $total+=$row2["TC_OFB14"];
                  if ( $row2["TC_OFB11"]!=$oldrxno) {
                    $i++;
                    $ii=$i;
                    $rxno =$row2['TC_OFB11'];
                    $oldrxno=$row2['TC_OFB11'];
                  } else {
                    $rxno=$row2['TC_OFB11'];
                    $ii='';
                  }


                  if ($row2['TC_OFB05']=="G") {
                      $unit="G";
                  } else  {
                      $unit="Unit";
                      $unittotal+=$row2["TC_OFB12"];
                  }

                  //invoice目前沒有抛轉patient過來 先改由訂單去抓patient
                  $tc_ofbud05=$row2['TC_OFBUD05'];
                  $tc_ofb31=$row2['TC_OFB31'];
                  $soea="select ta_oea003 from oea_file, oga_file where oea01=oga16 and oga01='$tc_ofb31' ";
                  $erp_sqloea = oci_parse($erp_conn2,$soea );
                  oci_execute($erp_sqloea);
                  $rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC) ;
                  //$tc_ofbud05=$rowoea['TA_OEA003'];

                  // Case No. :           $rxno
                  // Patient Name:        TC_OFBUD05
                  // Product Description: TC_OFB06)
                  // Qty                  TC_OFB12
                  // Unit Price           TC_OFB13
                  // Total                TC_OFB14
                  // Tooth Number         TC_OFBUD02
                  // memo                 TC_OFBUD01
                  $tcofbud02='';
                  if ($row2["TC_OFBUD02"]!=''){
                      if ($occud08==2 or $occud08==3) {
                          if (substr($row2["TC_OFBUD02"],0,1)!='U' and substr($row2["TC_OFBUD02"],0,1)!='L' ) {
                              $tcofbud02='#'.changeteethno($row2["TC_OFBUD02"],$occud08);
                          } else {
                              $tcofbud02=$row2['TC_OFBUD02'];
                          }
                      } else {
                          $tcofbud02=$row2['TC_OFBUD02'];
                      }
                  }

                  $osheet ->setCellValue('A'. $y, $ii)
                          ->setCellValue('B'. $y, $rxno . ' ')
                          ->setCellValue('C'. $y, $tc_ofbud05)
                          ->setCellValue('D'. $y, $row2["TA_OEA001"])
                          ->setCellValue('E'. $y, $row2["TC_OFB06"])
                          ->setCellValue('F'. $y, $row2["TC_OFB12"])
                          ->setCellValue('G'. $y, $unit)
                          ->setCellValue('H'. $y, number_format($row2["TC_OFB13"],2,'.',','))
                          ->setCellValue('I'. $y, number_format($row2["TC_OFB14"],2,'.',','))
                          ->setCellValue('J'. $y, $row2["TC_OFBUD01"]);



                  $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                  $osheet ->getStyle('J'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);

                  if (($y%2)==0){
                      $osheet->getStyle('A'.$y.':J'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                  }
                  $y++;
                }
                $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('H'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('I'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('J'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                //total
                $osheet ->setCellValue('A'.$y, 'Total')
                        ->setCellValue('F'.$y, $unittotal)
                        ->setCellValue('G'.$y, 'Units')
                        ->setCellValue('H'.$y, $currency)
                        ->setCellValue('I'.$y, number_format($total,2,'.',','));
                $osheet ->getStyle('G'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                $osheet ->getStyle('H'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('I'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('I'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );

                $osheet ->getStyle('A'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('B'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('C'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('D'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('F'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('G'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('H'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('I'.$y) ->getFont()->setBold(true);

                //第二個sheet 放DEF
                $s2="select tc_ofa01, tc_ofa23, tc_ofb11, tc_ofbud05, tc_ofb08, tc_ofb04, tc_ofb06, tc_ofb12, tc_ofb05, tc_ofb13, tc_ofb14, tc_ofb31, tc_ofbud02, tc_ofbud01, tc_ofb31, imaud04, ta_oea001 " .
                    "from tc_ofa_file, tc_ofb_file, ima_file, oga_file, oea_file " .
                    "where tc_ofaconf='Y' and upper(substr(trim(ta_oea001),1,1)) in ('D','E','F') and tc_ofb31=oga01 and oga16=oea01 and tc_ofa02=to_date('$bdate1','yy/mm/dd') and tc_ofa04='$occ01' and tc_ofa01=tc_ofb01 and tc_ofb04=ima01  $invoicefilter ".
                    "order by tc_ofa01, upper(trim(ta_oea001)), tc_ofb11,tc_ofb08,tc_ofb04 ";

                $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
                $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣

                $objPHPExcel->setActiveSheetIndex(1);
                $osheet = $objPHPExcel->getActiveSheet();

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);

                $y=5;
                $osheet -> setCellValue('A'.$y, "INVOICE NO.:".$row2["TC_OFA01"]. " (" . $occ02 . ")")
                        -> setCellValue('F'.$y, "INVOICE DATE:".$bdate);

                $s3="select occ18, occ231, occ232, occ261, occ271 from occ_file where occ01='$occ01'";
                $erp_sql3 = oci_parse($erp_conn2,$s3 );
                oci_execute($erp_sql3);
                $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);
                $tel='';
                if ($row3['occ261']!='') $tel="Tel: " . $row3['occ261'] . "    ";
                if ($row3['occ271']!='') $tel.="Fax:" . $row3['occ271'];

                $y=6;
                $osheet -> setCellValue('A'.$y, $row3["OCC18"]);
                $y=7;
                $osheet -> setCellValue('A'.$y, $row3["OCC231"]);
                $y=8;
                $osheet -> setCellValue('A'.$y, $row3["OCC232"]);
                $y=9;
                $osheet -> setCellValue('A'.$y, $tel);

                $y=12;
                //$currency=$row2["TC_OFA23"];
                $i=0;
                $oldrxno='';
                $total=0;

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
                  $total+=$row2["TC_OFB14"];
                  if ( $row2["TC_OFB11"]!=$oldrxno) {
                    $i++;
                    $ii=$i;
                    $rxno =$row2['TC_OFB11'];
                    $oldrxno=$row2['TC_OFB11'];
                  } else {
                    $rxno=$row2['TC_OFB11'];
                    $ii='';
                  }


                  if ($row2['TC_OFB05']=="G") {
                      $unit="G";
                  } else  {
                      $unit="Unit";
                      $unittotal+=$row2["TC_OFB12"];
                  }

                  //invoice目前沒有抛轉patient過來 先改由訂單去抓patient
                  $tc_ofbud05=$row2['TC_OFBUD05'];
                  $tc_ofb31=$row2['TC_OFB31'];
                  $soea="select ta_oea003 from oea_file, oga_file where oea01=oga16 and oga01='$tc_ofb31' ";
                  $erp_sqloea = oci_parse($erp_conn2,$soea );
                  oci_execute($erp_sqloea);
                  $rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC) ;
                  //$tc_ofbud05=$rowoea['TA_OEA003'];

                  // Case No. :           $rxno
                  // Patient Name:        TC_OFBUD05
                  // Product Description: TC_OFB06)
                  // Qty                  TC_OFB12
                  // Unit Price           TC_OFB13
                  // Total                TC_OFB14
                  // Tooth Number         TC_OFBUD02
                  // memo                 TC_OFBUD01
                  $tcofbud02='';
                  if ($row2["TC_OFBUD02"]!=''){
                      if ($occud08==2 or $occud08==3) {
                          if (substr($row2["TC_OFBUD02"],0,1)!='U' and substr($row2["TC_OFBUD02"],0,1)!='L' ) {
                              $tcofbud02='#'.changeteethno($row2["TC_OFBUD02"],$occud08);
                          } else {
                              $tcofbud02=$row2['TC_OFBUD02'];
                          }
                      } else {
                          $tcofbud02=$row2['TC_OFBUD02'];
                      }
                  }

                  $osheet ->setCellValue('A'. $y, $ii)
                          ->setCellValue('B'. $y, $rxno . ' ')
                          ->setCellValue('C'. $y, $tc_ofbud05)
                          ->setCellValue('D'. $y, $row2["TA_OEA001"])
                          ->setCellValue('E'. $y, $row2["TC_OFB06"])
                          ->setCellValue('F'. $y, $row2["TC_OFB12"])
                          ->setCellValue('G'. $y, $unit)
                          ->setCellValue('H'. $y, number_format($row2["TC_OFB13"],2,'.',','))
                          ->setCellValue('I'. $y, number_format($row2["TC_OFB14"],2,'.',','))
                          ->setCellValue('J'. $y, $row2["TC_OFBUD01"]);



                  $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                  $osheet ->getStyle('J'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                  if (($y%2)==0){
                      $osheet->getStyle('A'.$y.':J'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                  }
                  $y++;
                }
                $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('H'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('I'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('J'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                //total
                $osheet ->setCellValue('A'.$y, 'Total')
                        ->setCellValue('F'.$y, $unittotal)
                        ->setCellValue('G'.$y, 'Units')
                        ->setCellValue('H'.$y, $currency)
                        ->setCellValue('I'.$y, number_format($total,2,'.',','));
                $osheet ->getStyle('G'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                $osheet ->getStyle('H'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('I'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('I'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );

                $osheet ->getStyle('A'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('B'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('C'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('D'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('F'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('G'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('H'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('I'.$y) ->getFont()->setBold(true);

                $objPHPExcel->setActiveSheetIndex(0);

        }
        else if($template=='E181003')
        {  // CDM-C
                $s2="select tc_ofa01, tc_ofa23, tc_ofb11, tc_ofbud05, tc_ofb08, tc_ofb04, tc_ofb06, tc_ofb12, tc_ofb05, tc_ofb13, tc_ofb14, tc_ofb31, tc_ofbud02, tc_ofbud01, tc_ofb31, imaud04, ta_oea001 " .
                    "from tc_ofa_file, tc_ofb_file, ima_file, oga_file, oea_file " .
                    "where tc_ofaconf='Y' and tc_ofb31=oga01 and oga16=oea01 and tc_ofa02=to_date('$bdate1','yy/mm/dd') and tc_ofa04='$occ01' and tc_ofa01=tc_ofb01 and tc_ofb04=ima01  $invoicefilter ".
                    "order by tc_ofa01, upper(trim(ta_oea001)), tc_ofb11,tc_ofb08,tc_ofb04 ";

                $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
                $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣

                $objPHPExcel->setActiveSheetIndex(0);
                $osheet = $objPHPExcel->getActiveSheet();

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);

                $y=5;
                $osheet -> setCellValue('A'.$y, "INVOICE NO.:".$row2["TC_OFA01"]. " (" . $occ02 . ")")
                        -> setCellValue('F'.$y, "INVOICE DATE:".$bdate);

                $s3="select occ18, occ231, occ232, occ261, occ271 from occ_file where occ01='$occ01'";
                $erp_sql3 = oci_parse($erp_conn2,$s3 );
                oci_execute($erp_sql3);
                $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);
                $tel='';
                if ($row3['occ261']!='') $tel="Tel: " . $row3['occ261'] . "    ";
                if ($row3['occ271']!='') $tel.="Fax:" . $row3['occ271'];

                $y=6;
                $osheet -> setCellValue('A'.$y, $row3["OCC18"]);
                $y=7;
                $osheet -> setCellValue('A'.$y, $row3["OCC231"]);
                $y=8;
                $osheet -> setCellValue('A'.$y, $row3["OCC232"]);
                $y=9;
                $osheet -> setCellValue('A'.$y, $tel);

                $y=12;
                //$currency=$row2["TC_OFA23"];
                $i=0;
                $oldrxno='';
                $total=0;

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
                  $total+=$row2["TC_OFB14"];
                  if ( $row2["TC_OFB11"]!=$oldrxno) {
                    $i++;
                    $ii=$i;
                    $rxno =$row2['TC_OFB11'];
                    $oldrxno=$row2['TC_OFB11'];
                  } else {
                    $rxno=$row2['TC_OFB11'];
                    $ii='';
                  }


                  if ($row2['TC_OFB05']=="G") {
                      $unit="G";
                  } else  {
                      $unit="Unit";
                      $unittotal+=$row2["TC_OFB12"];
                  }

                  //invoice目前沒有抛轉patient過來 先改由訂單去抓patient
                  $tc_ofbud05=$row2['TC_OFBUD05'];
                  $tc_ofb31=$row2['TC_OFB31'];
                  $soea="select ta_oea003 from oea_file, oga_file where oea01=oga16 and oga01='$tc_ofb31' ";
                  $erp_sqloea = oci_parse($erp_conn2,$soea );
                  oci_execute($erp_sqloea);
                  $rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC) ;
                  //$tc_ofbud05=$rowoea['TA_OEA003'];

                  // Case No. :           $rxno
                  // Patient Name:        TC_OFBUD05
                  // Product Description: TC_OFB06)
                  // Qty                  TC_OFB12
                  // Unit Price           TC_OFB13
                  // Total                TC_OFB14
                  // Tooth Number         TC_OFBUD02
                  // memo                 TC_OFBUD01
                  $tcofbud02='';
                  if ($row2["TC_OFBUD02"]!=''){
                      if ($occud08==2 or $occud08==3) {
                          if (substr($row2["TC_OFBUD02"],0,1)!='U' and substr($row2["TC_OFBUD02"],0,1)!='L' ) {
                              $tcofbud02='#'.changeteethno($row2["TC_OFBUD02"],$occud08);
                          } else {
                              $tcofbud02=$row2['TC_OFBUD02'];
                          }
                      } else {
                          $tcofbud02=$row2['TC_OFBUD02'];
                      }
                  }

                  $osheet ->setCellValue('A'. $y, $ii)
                          ->setCellValue('B'. $y, $rxno . ' ')
                          ->setCellValue('C'. $y, $tc_ofbud05)
                          ->setCellValue('D'. $y, $row2["TA_OEA001"])
                          ->setCellValue('E'. $y, $row2["TC_OFB06"])
                          ->setCellValue('F'. $y, $row2["TC_OFB12"])
                          ->setCellValue('G'. $y, $unit)
                          ->setCellValue('H'. $y, number_format($row2["TC_OFB13"],2,'.',','))
                          ->setCellValue('I'. $y, number_format($row2["TC_OFB14"],2,'.',','))
                          ->setCellValue('J'. $y, $row2["TC_OFBUD01"]);



                  $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                  $osheet ->getStyle('J'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);

                  if (($y%2)==0){
                      $osheet->getStyle('A'.$y.':J'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                  }
                  $y++;
                }
                $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('H'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('I'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('J'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                //total
                $osheet ->setCellValue('A'.$y, 'Total')
                        ->setCellValue('F'.$y, $unittotal)
                        ->setCellValue('G'.$y, 'Units')
                        ->setCellValue('H'.$y, $currency)
                        ->setCellValue('I'.$y, number_format($total,2,'.',','));
                $osheet ->getStyle('G'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                $osheet ->getStyle('H'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('I'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('I'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );

                $osheet ->getStyle('A'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('B'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('C'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('D'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('F'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('G'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('H'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('I'.$y) ->getFont()->setBold(true);

                $objPHPExcel->setActiveSheetIndex(0);

        }
        else if($template=='E181004')
        {  // CDM-M
                $s2="select tc_ofa01, tc_ofa23, tc_ofb11, tc_ofbud05, tc_ofb08, tc_ofb04, tc_ofb06, tc_ofb12, tc_ofb05, tc_ofb13, tc_ofb14, tc_ofb31, tc_ofbud02, tc_ofbud01, tc_ofb31, imaud04, ta_oea001 " .
                    "from tc_ofa_file, tc_ofb_file, ima_file, oga_file, oea_file " .
                    "where tc_ofaconf='Y' and tc_ofb31=oga01 and oga16=oea01 and tc_ofa02=to_date('$bdate1','yy/mm/dd') and tc_ofa04='$occ01' and tc_ofa01=tc_ofb01 and tc_ofb04=ima01  $invoicefilter ".
                    "order by tc_ofa01, upper(trim(ta_oea001)), tc_ofb11,tc_ofb08,tc_ofb04 ";

                $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
                $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣

                $objPHPExcel->setActiveSheetIndex(0);
                $osheet = $objPHPExcel->getActiveSheet();

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);

                $y=5;
                $osheet -> setCellValue('A'.$y, "INVOICE NO.:".$row2["TC_OFA01"]. " (" . $occ02 . ")")
                        -> setCellValue('F'.$y, "INVOICE DATE:".$bdate);

                $s3="select occ18, occ231, occ232, occ261, occ271 from occ_file where occ01='$occ01'";
                $erp_sql3 = oci_parse($erp_conn2,$s3 );
                oci_execute($erp_sql3);
                $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);
                $tel='';
                if ($row3['occ261']!='') $tel="Tel: " . $row3['occ261'] . "    ";
                if ($row3['occ271']!='') $tel.="Fax:" . $row3['occ271'];

                $y=6;
                $osheet -> setCellValue('A'.$y, $row3["OCC18"]);
                $y=7;
                $osheet -> setCellValue('A'.$y, $row3["OCC231"]);
                $y=8;
                $osheet -> setCellValue('A'.$y, $row3["OCC232"]);
                $y=9;
                $osheet -> setCellValue('A'.$y, $tel);

                $y=12;
                //$currency=$row2["TC_OFA23"];
                $i=0;
                $oldrxno='';
                $total=0;

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
                  $total+=$row2["TC_OFB14"];
                  if ( $row2["TC_OFB11"]!=$oldrxno) {
                    $i++;
                    $ii=$i;
                    $rxno =$row2['TC_OFB11'];
                    $oldrxno=$row2['TC_OFB11'];
                  } else {
                    $rxno=$row2['TC_OFB11'];
                    $ii='';
                  }


                  if ($row2['TC_OFB05']=="G") {
                      $unit="G";
                  } else  {
                      $unit="Unit";
                      $unittotal+=$row2["TC_OFB12"];
                  }

                  //invoice目前沒有抛轉patient過來 先改由訂單去抓patient
                  $tc_ofbud05=$row2['TC_OFBUD05'];
                  $tc_ofb31=$row2['TC_OFB31'];
                  $soea="select ta_oea003 from oea_file, oga_file where oea01=oga16 and oga01='$tc_ofb31' ";
                  $erp_sqloea = oci_parse($erp_conn2,$soea );
                  oci_execute($erp_sqloea);
                  $rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC) ;
                  //$tc_ofbud05=$rowoea['TA_OEA003'];

                  // Case No. :           $rxno
                  // Patient Name:        TC_OFBUD05
                  // Product Description: TC_OFB06)
                  // Qty                  TC_OFB12
                  // Unit Price           TC_OFB13
                  // Total                TC_OFB14
                  // Tooth Number         TC_OFBUD02
                  // memo                 TC_OFBUD01
                  $tcofbud02='';
                  if ($row2["TC_OFBUD02"]!=''){
                      if ($occud08==2 or $occud08==3) {
                          if (substr($row2["TC_OFBUD02"],0,1)!='U' and substr($row2["TC_OFBUD02"],0,1)!='L' ) {
                              $tcofbud02='#'.changeteethno($row2["TC_OFBUD02"],$occud08);
                          } else {
                              $tcofbud02=$row2['TC_OFBUD02'];
                          }
                      } else {
                          $tcofbud02=$row2['TC_OFBUD02'];
                      }
                  }

                  $osheet ->setCellValue('A'. $y, $ii)
                          ->setCellValue('B'. $y, $rxno . ' ')
                          ->setCellValue('C'. $y, $tc_ofbud05)
                          ->setCellValue('D'. $y, $row2["TA_OEA001"])
                          ->setCellValue('E'. $y, $row2["TC_OFB06"])
                          ->setCellValue('F'. $y, $row2["TC_OFB12"])
                          ->setCellValue('G'. $y, $unit)
                          ->setCellValue('H'. $y, number_format($row2["TC_OFB13"],2,'.',','))
                          ->setCellValue('I'. $y, number_format($row2["TC_OFB14"],2,'.',','))
                          ->setCellValue('J'. $y, $row2["TC_OFBUD01"]);



                  $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                  $osheet ->getStyle('J'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);

                  if (($y%2)==0){
                      $osheet->getStyle('A'.$y.':J'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                  }
                  $y++;
                }
                $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('H'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('I'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('J'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                //total
                $osheet ->setCellValue('A'.$y, 'Total')
                        ->setCellValue('F'.$y, $unittotal)
                        ->setCellValue('G'.$y, 'Units')
                        ->setCellValue('H'.$y, $currency)
                        ->setCellValue('I'.$y, number_format($total,2,'.',','));
                $osheet ->getStyle('G'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                $osheet ->getStyle('H'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('I'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('I'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );

                $osheet ->getStyle('A'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('B'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('C'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('D'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('F'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('G'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('H'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('I'.$y) ->getFont()->setBold(true);

                $objPHPExcel->setActiveSheetIndex(0);

        }
        else if($template=='T1850031')
        {  // TWDEW 2016/07/20 加印中文名稱

                $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
                $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣

                $objPHPExcel->setActiveSheetIndex(0);
                $osheet = $objPHPExcel->getActiveSheet();

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);

                $y=5;
                $osheet -> setCellValue('A'.$y, "INVOICE NO.:".$row2["TC_OFA01"]. " (" . $occ02 . ")")
                        -> setCellValue('E'.$y, "INVOICE DATE:".$bdate);

                $s3="select occ18, occ231, occ232, occ261, occ271 from occ_file where occ01='$occ01'";
                $erp_sql3 = oci_parse($erp_conn2,$s3 );
                oci_execute($erp_sql3);
                $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);
                $tel='';
                if ($row3['occ261']!='') $tel="Tel: " . $row3['occ261'] . "    ";
                if ($row3['occ271']!='') $tel.="Fax:" . $row3['occ271'];

                $y=6;
                $osheet -> setCellValue('A'.$y, $row3["OCC18"]);
                $y=7;
                $osheet -> setCellValue('A'.$y, $row3["OCC231"]);
                $y=8;
                $osheet -> setCellValue('A'.$y, $row3["OCC232"]);
                $y=9;
                $osheet -> setCellValue('A'.$y, $tel);

                $y=12;
                //$currency=$row2["TC_OFA23"];
                $i=0;
                $oldrxno='';
                $total=0;

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
                  $total+=$row2["TC_OFB14"];
                  if ( $row2["TC_OFB11"]!=$oldrxno) {
                    $i++;
                    $ii=$i;
                    $rxno =$row2['TC_OFB11'];
                    $oldrxno=$row2['TC_OFB11'];
                  } else {
                    $rxno=$row2['TC_OFB11'];
                    $ii='';
                  }


                  if ($row2['TC_OFB05']=="G") {
                      $unit="G";
                  } else  {
                      $unit="Unit";
                      $unittotal+=$row2["TC_OFB12"];
                  }

                  //invoice目前沒有抛轉patient過來 先改由訂單去抓patient
                  $tc_ofbud05=$row2['TC_OFBUD05'];
                  $tc_ofb31=$row2['TC_OFB31'];
                  $soea="select ta_oea003 from oea_file, oga_file where oea01=oga16 and oga01='$tc_ofb31' ";
                  $erp_sqloea = oci_parse($erp_conn2,$soea );
                  oci_execute($erp_sqloea);
                  $rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC) ;
                  //$tc_ofbud05=$rowoea['TA_OEA003'];
                  // Case No. :           $rxno
                  // Patient Name:        TC_OFBUD05
                  // Product Description: TC_OFB06)
                  // Qty                  TC_OFB12
                  // Unit Price           TC_OFB13
                  // Total                TC_OFB14
                  // Tooth Number         TC_OFBUD02
                  // memo                 TC_OFBUD01

                  $osheet ->setCellValue('A'. $y, $ii)
                          ->setCellValue('B'. $y, $rxno . ' ')
                          ->setCellValue('C'. $y, substr($tc_ofbud05,0,15))
                          ->setCellValue('D'. $y, $row2["TC_OFB06"])
                          ->setCellValue('E'. $y, $row2["TC_OFB12"])
                          ->setCellValue('F'. $y, $unit)
                          ->setCellValue('G'. $y, number_format($row2["TC_OFB13"],2,'.',','))
                          ->setCellValue('H'. $y, number_format($row2["TC_OFB14"],2,'.',','))
                          ->setCellValue('I'. $y, $row2["TC_OFBUD01"]);
                  $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                  $osheet ->getStyle('I'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                  //2017.07.13取消隔行變色
                  //if (($y%2)==0){
                  //    $osheet->getStyle('A'.$y.':I'.$y)->getFill() //->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                  //}
                  $y++;

                  //加印一行中文品代
                  $osheet ->setCellValue('D'. $y, $row2['IMA02']);
                  if (($y%2)==0){
                      $osheet->getStyle('A'.$y.':I'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                  }
                  $y++;
                }


                $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('H'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('I'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                //total
                $osheet ->setCellValue('A'.$y, 'Total')
                        ->setCellValue('E'.$y, $unittotal)
                        ->setCellValue('F'.$y, 'Units')
                        ->setCellValue('G'.$y, $currency)
                        ->setCellValue('H'.$y, number_format($total,2,'.',','));
                $osheet ->getStyle('F'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                $osheet ->getStyle('G'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('H'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('H'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );

                $osheet ->getStyle('A'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('B'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('C'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('D'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('F'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('G'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('H'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('I'.$y) ->getFont()->setBold(true);

                // Rename sheet
                $osheet ->setTitle('DailyInvoice');
                // Set active sheet index to the first sheet, so Excel opens this as the first sheet
                $objPHPExcel->setActiveSheetIndex(0);

        }
        else if($template=='T185003')
        { //無patient 有齒位

                $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
                $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣

                $objPHPExcel->setActiveSheetIndex(0);
                $osheet = $objPHPExcel->getActiveSheet();

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);

                $y=5;
                $osheet -> setCellValue('A'.$y, "INVOICE NO.:".$row2["TC_OFA01"]. " (" . $occ02 . ")")
                        -> setCellValue('E'.$y, "INVOICE DATE:".$bdate);

                $s3="select occ18, occ231, occ232, occ261, occ271 from occ_file where occ01='$occ01'";
                $erp_sql3 = oci_parse($erp_conn2,$s3 );
                oci_execute($erp_sql3);
                $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);
                $tel='';
                if ($row3['occ261']!='') $tel="Tel: " . $row3['occ261'] . "    ";
                if ($row3['occ271']!='') $tel.="Fax:" . $row3['occ271'];

                $y=6;
                $osheet -> setCellValue('A'.$y, $row3["OCC18"]);
                $y=7;
                $osheet -> setCellValue('A'.$y, $row3["OCC231"]);
                $y=8;
                $osheet -> setCellValue('A'.$y, $row3["OCC232"]);
                $y=9;
                $osheet -> setCellValue('A'.$y, $tel);

                $y=12;
                //$currency=$row2["TC_OFA23"];
                $i=0;
                $oldrxno='';
                $total=0;

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
                  $total+=$row2["TC_OFB14"];
                  if ( $row2["TC_OFB11"]!=$oldrxno) {
                    $i++;
                    $ii=$i;
                    $rxno =$row2['TC_OFB11'];
                    $oldrxno=$row2['TC_OFB11'];
                  } else {
                    $rxno=$row2['TC_OFB11'];
                    $ii='';
                  }


                  if ($row2['TC_OFB05']=="G") {
                      $unit="G";
                  } else  {
                      $unit="Unit";
                      $unittotal+=$row2["TC_OFB12"];
                  }

                  // Case No. :           $rxno
                  // Patient Name:        TC_OFBUD05
                  // Product Description: TC_OFB06)
                  // Qty                  TC_OFB12
                  // Unit Price           TC_OFB13
                  // Total                TC_OFB14
                  // Tooth Number         TC_OFBUD02
                  // memo                 TC_OFBUD01
                  $tcofbud02='';
                  if ($row2["TC_OFBUD02"]!=''){
                      if ($occud08==2 or $occud08==3) {
                          if (substr($row2["TC_OFBUD02"],0,1)!='U' and substr($row2["TC_OFBUD02"],0,1)!='L' ) {
                              $tcofbud02='#'.changeteethno($row2["TC_OFBUD02"],$occud08);
                          } else {
                              $tcofbud02=$row2['TC_OFBUD02'];
                          }
                      } else {
                          $tcofbud02=$row2['TC_OFBUD02'];
                      }
                  }
                  $osheet ->setCellValue('A'. $y, $ii)
                          ->setCellValue('B'. $y, $rxno . ' ')
                          ->setCellValue('C'. $y, $row2["TC_OFB06"])
                          ->setCellValue('D'. $y, $row2["TC_OFB12"])
                          ->setCellValue('E'. $y, $unit)
                          ->setCellValue('F'. $y, number_format($row2["TC_OFB13"],2,'.',','))
                          ->setCellValue('G'. $y, number_format($row2["TC_OFB14"],2,'.',','))
                          ->setCellValue('H'. $y, $tcofbud02)
                          ->setCellValue('I'. $y, $row2["TC_OFBUD01"]);
                  $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                  $osheet ->getStyle('H'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                  $osheet ->getStyle('I'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                  if (($y%2)==0){
                      $osheet->getStyle('A'.$y.':I'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                  }
                  $y++;

                  //加印一行中文品代
                  $osheet ->setCellValue('C'. $y, $row2['IMA02']);
                  if (($y%2)==0){
                      $osheet->getStyle('A'.$y.':I'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                  }
                  $y++;
                }
                $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('H'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('I'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                //total
                $osheet ->setCellValue('A'.$y, 'Total')
                        ->setCellValue('D'.$y, $unittotal)
                        ->setCellValue('E'.$y, 'Units')
                        ->setCellValue('F'.$y, $currency)
                        ->setCellValue('G'.$y, number_format($total,2,'.',','));
                $osheet ->getStyle('F'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('G'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );

                $osheet ->getStyle('A'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('B'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('C'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('D'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('F'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('G'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('H'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('I'.$y) ->getFont()->setBold(true);

                // Rename sheet
                $osheet ->setTitle('DailyInvoice');
                // Set active sheet index to the first sheet, so Excel opens this as the first sheet
                $objPHPExcel->setActiveSheetIndex(0);

        }
        else if ($template=="1YN")
        {  //DM
                $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(9);
                $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣

                $objPHPExcel->setActiveSheetIndex(0);
                $osheet = $objPHPExcel->getActiveSheet();

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);

                $y=6;
                $osheet -> setCellValue('A'.$y, "INVOICE NO.:".$row2["TC_OFA01"]. $occtitle)
                        -> setCellValue('E'.$y, "INVOICE DATE:".$bdate);

                $y=12;
                //$currency=$row2["TC_OFA23"];
                $i=0;
                $oldrxno='';
                $total=0;
                $unittotal=0;
                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
                  $total+=$row2["TC_OFB14"];
                  if ( $row2["TC_OFB11"]!=$oldrxno) {
                    $i++;
                    $ii=$i;
                    $rxno =$row2['TC_OFB11'];
                    $oldrxno=$row2['TC_OFB11'];
                  } else {
                    $rxno=$row2['TC_OFB11'];
                    $ii='';
                  }


                  if ($row2['TC_OFB05']=="G") {
                      $unit="G";
                  } else  {
                      $unit="Unit";
                      $unittotal+=$row2["TC_OFB12"];
                  }

                  //invoice目前沒有抛轉patient過來 先改由訂單去抓patient
                  $tc_ofbud05=$row2['TC_OFBUD05'];
                  $tc_ofb31=$row2['TC_OFB31'];
                  $soea="select ta_oea003 from oea_file, oga_file where oea01=oga16 and oga01='$tc_ofb31' ";
                  $erp_sqloea = oci_parse($erp_conn2,$soea );
                  oci_execute($erp_sqloea);
                  $rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC) ;
                  //$tc_ofbud05=$rowoea['TA_OEA003'];



                  // Case No. :           $rxno
                  // Patient Name:        TC_OFBUD05
                  // Product Description: TC_OFB06)
                  // Qty                  TC_OFB12
                  // Unit Price           TC_OFB13
                  // Total                TC_OFB14
                  // Tooth Number         TC_OFBUD02
                  // memo                 TC_OFBUD01

                  $osheet ->setCellValue('A'. $y, $ii)
                          ->setCellValue('B'. $y, $rxno . ' ')
                          ->setCellValue('C'. $y, $tc_ofbud05)
                          ->setCellValue('D'. $y, $row2["TC_OFB06"])
                          ->setCellValue('E'. $y, $row2["TC_OFB12"])
                          ->setCellValue('F'. $y, $unit)
                          ->setCellValue('G'. $y, number_format($row2["TC_OFB13"],2,'.',','))
                          ->setCellValue('H'. $y, number_format($row2["TC_OFB14"],2,'.',','))
                          ->setCellValue('I'. $y, $row2["TC_OFBUD01"]);
                  $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                  $osheet ->getStyle('I'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                  if (($y%2)==0){
                      $osheet->getStyle('A'.$y.':I'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                  }
                  $y++;
                }
                $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('H'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('I'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                //total
                $osheet ->setCellValue('A'.$y, 'Total')
                        ->setCellValue('E'.$y, $unittotal)
                        ->setCellValue('F'.$y, 'Units')
                        ->setCellValue('G'.$y, $currency)
                        ->setCellValue('H'.$y, number_format($total,2,'.',','));
                $osheet ->getStyle('F'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                $osheet ->getStyle('G'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('H'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('H'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );

                $osheet ->getStyle('A'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('B'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('C'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('D'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('F'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('G'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('H'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('I'.$y) ->getFont()->setBold(true);

                // Rename sheet
                $osheet ->setTitle('DailyInvoice');
                // Set active sheet index to the first sheet, so Excel opens this as the first sheet
                $objPHPExcel->setActiveSheetIndex(0);

        }
        else if ($template=="2NN")
        { //FDL
                $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
                $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
                //$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
                //$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(13);
                //$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(39);
                //$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(6);
                //$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(4.5);
                //$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(9.5);
                //$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(11);
                //$objPHPExcel->getActiveSheet()->getPageMargins()->setTop(0.5);
                //$objPHPExcel->getActiveSheet()->getPageMargins()->setRight(0.15);
                //$objPHPExcel->getActiveSheet()->getPageMargins()->setLeft(0.15);
                //$objPHPExcel->getActiveSheet()->getPageMargins()->setBottom(0.5);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 8); //每頁的1-5row都重複一樣

                //$objPHPExcel->getActiveSheet()->mergeCells('A18:E22'); 合併cell
                //$objPHPExcel->getActiveSheet()->getStyle('B2')->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_RED);
                //$objPHPExcel->getActiveSheet()->getStyle('B2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                //$objPHPExcel->getActiveSheet()->getStyle('B2')->getFill()->getStartColor()->setARGB('FFFF0000');
                //$objPHPExcel->getActiveSheet()->getStyle('B3:B7')->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFFF0000');
                //$objConditional1->getStyle()->getFont()->setBold(true);

                $objPHPExcel->setActiveSheetIndex(0);
                $osheet = $objPHPExcel->getActiveSheet();

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);

                $y=2;
                $osheet -> setCellValue('A'.$y, "INVOICE NO.:".$row2["TC_OFA01"].  $occtitle)
                        -> setCellValue('E'.$y, "INVOICE DATE:".$bdate);

                $y=9;
                //$currency=$row2["TC_OFA23"];
                $i=0;
                $oldrxno='';
                $total=0;
                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
                  $total+=$row2["TC_OFB14"];
                  if ( $row2["TC_OFB11"]!=$oldrxno) {
                    $i++;
                    $ii=$i;
                    $rxno =$row2['TC_OFB11'];
                    $oldrxno=$row2['TC_OFB11'];
                  } else {
                    $rxno=$row2['TC_OFB11'];
                    $ii='';
                  }

                  if ($row2['TC_OFB05']=="G") {
                      $unit="G";
                  } else  {
                      $unit="Unit";
                      $unittotal+=$row2["TC_OFB12"];
                  }
                  $osheet ->setCellValue('A'. $y, $ii)
                          ->setCellValue('B'. $y, $rxno . ' ')
                          ->setCellValue('C'. $y, $row2["TC_OFB06"])
                          ->setCellValue('D'. $y, $row2["TC_OFB12"])
                          ->setCellValue('E'. $y, $unit)
                          ->setCellValue('F'. $y, number_format($row2["TC_OFB13"],2,'.',','))
                          ->setCellValue('G'. $y, number_format($row2["TC_OFB14"],2,'.',','))
                          ->setCellValue('H'. $y, $row2["TC_OFBUD01"]);
                  $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                  $osheet ->getStyle('H'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                  if (($y%2)==0){
                      $osheet->getStyle('A'.$y.':H'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                  }
                  $y++;
                }
                $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('H'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                //total
                $osheet ->setCellValue('A'.$y, 'Total')
                        ->setCellValue('D'.$y, $unittotal)
                        ->setCellValue('E'.$y, 'Units')
                        ->setCellValue('F'.$y, $currency)
                        ->setCellValue('G'.$y, number_format($total,2,'.',','));
                $osheet ->getStyle('E'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                $osheet ->getStyle('F'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('G'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('G'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );

                $osheet ->getStyle('A'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('B'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('C'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('D'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('F'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('G'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('H'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('I'.$y) ->getFont()->setBold(true);

                // Rename sheet
                $osheet ->setTitle('DailyInvoice');
                // Set active sheet index to the first sheet, so Excel opens this as the first sheet
                $objPHPExcel->setActiveSheetIndex(0);

        }
        else if($template=='NYN')
        {

                $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
                $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣

                $objPHPExcel->setActiveSheetIndex(0);
                $osheet = $objPHPExcel->getActiveSheet();

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);

                $y=5;
                $osheet -> setCellValue('A'.$y, "INVOICE NO.:".$row2["TC_OFA01"].  $occtitle)
                        -> setCellValue('E'.$y, "INVOICE DATE:".$bdate);

                $s3="select occ18, occ231, occ232, occ261, occ271 from occ_file where occ01='$occ01'";
                $erp_sql3 = oci_parse($erp_conn2,$s3 );
                oci_execute($erp_sql3);
                $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);
                $tel='';
                if ($row3['occ261']!='') $tel="Tel: " . $row3['occ261'] . "    ";
                if ($row3['occ271']!='') $tel.="Fax:" . $row3['occ271'];

                $y=6;
                $osheet -> setCellValue('A'.$y, $row3["OCC18"]);
                $y=7;
                $osheet -> setCellValue('A'.$y, $row3["OCC231"]);
                $y=8;
                $osheet -> setCellValue('A'.$y, $row3["OCC232"]);
                $y=9;
                $osheet -> setCellValue('A'.$y, $tel);

                $y=12;
                //$currency=$row2["TC_OFA23"];
                $i=0;
                $oldrxno='';
                $total=0;

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
                  $total+=$row2["TC_OFB14"];
                  if ( $row2["TC_OFB11"]!=$oldrxno) {
                    $i++;
                    $ii=$i;
                    $rxno =$row2['TC_OFB11'];
                    $oldrxno=$row2['TC_OFB11'];
                  } else {
                    $rxno=$row2['TC_OFB11'];
                    $ii='';
                  }


                  if ($row2['TC_OFB05']=="G") {
                      $unit="G";
                  } else  {
                      $unit="Unit";
                      $unittotal+=$row2["TC_OFB12"];
                  }

                  //invoice目前沒有抛轉patient過來 先改由訂單去抓patient
                  $tc_ofbud05=$row2['TC_OFBUD05'];
                  $tc_ofb31=$row2['TC_OFB31'];
                  $soea="select ta_oea003 from oea_file, oga_file where oea01=oga16 and oga01='$tc_ofb31' ";
                  $erp_sqloea = oci_parse($erp_conn2,$soea );
                  oci_execute($erp_sqloea);
                  $rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC) ;
                  //$tc_ofbud05=$rowoea['TA_OEA003'];
                  // Case No. :           $rxno
                  // Patient Name:        TC_OFBUD05
                  // Product Description: TC_OFB06)
                  // Qty                  TC_OFB12
                  // Unit Price           TC_OFB13
                  // Total                TC_OFB14
                  // Tooth Number         TC_OFBUD02
                  // memo                 TC_OFBUD01

                  $osheet ->setCellValue('A'. $y, $ii)
                          ->setCellValue('B'. $y, $rxno . ' ')
                          ->setCellValue('C'. $y, substr($tc_ofbud05,0,15))
                          ->setCellValue('D'. $y, $row2["TC_OFB06"])
                          ->setCellValue('E'. $y, $row2["TC_OFB12"])
                          ->setCellValue('F'. $y, $unit)
                          ->setCellValue('G'. $y, number_format($row2["TC_OFB13"],2,'.',','))
                          ->setCellValue('H'. $y, number_format($row2["TC_OFB14"],2,'.',','))
                          ->setCellValue('I'. $y, $row2["TC_OFBUD01"]);
                  $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                  $osheet ->getStyle('I'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                  if (($y%2)==0){
                      $osheet->getStyle('A'.$y.':I'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                  }
                  $y++;
                }
                $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('H'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('I'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                //total
                $osheet ->setCellValue('A'.$y, 'Total')
                        ->setCellValue('E'.$y, $unittotal)
                        ->setCellValue('F'.$y, 'Units')
                        ->setCellValue('G'.$y, $currency)
                        ->setCellValue('H'.$y, number_format($total,2,'.',','));
                $osheet ->getStyle('F'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                $osheet ->getStyle('G'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('H'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('H'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );

                $osheet ->getStyle('A'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('B'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('C'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('D'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('F'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('G'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('H'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('I'.$y) ->getFont()->setBold(true);

                // Rename sheet
                $osheet ->setTitle('DailyInvoice');
                // Set active sheet index to the first sheet, so Excel opens this as the first sheet
                $objPHPExcel->setActiveSheetIndex(0);

        }
        else if($template=='NYY')
        {

                $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
                $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣

                $objPHPExcel->setActiveSheetIndex(0);
                $osheet = $objPHPExcel->getActiveSheet();

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);

                $y=5;
                $osheet -> setCellValue('A'.$y, "INVOICE NO.:".$row2["TC_OFA01"].  $occtitle)
                        -> setCellValue('E'.$y, "INVOICE DATE:".$bdate);

                $s3="select occ18, occ231, occ232, occ261, occ271 from occ_file where occ01='$occ01'";
                $erp_sql3 = oci_parse($erp_conn2,$s3 );
                oci_execute($erp_sql3);
                $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);
                $tel='';
                if ($row3['occ261']!='') $tel="Tel: " . $row3['occ261'] . "    ";
                if ($row3['occ271']!='') $tel.="Fax:" . $row3['occ271'];

                $y=6;
                $osheet -> setCellValue('A'.$y, $row3["OCC18"]);
                $y=7;
                $osheet -> setCellValue('A'.$y, $row3["OCC231"]);
                $y=8;
                $osheet -> setCellValue('A'.$y, $row3["OCC232"]);
                $y=9;
                $osheet -> setCellValue('A'.$y, $tel);

                $y=12;
                //$currency=$row2["TC_OFA23"];
                $i=0;
                $oldrxno='';
                $total=0;

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
                  $total+=$row2["TC_OFB14"];
                  if ( $row2["TC_OFB11"]!=$oldrxno) {
                    $i++;
                    $ii=$i;
                    $rxno =$row2['TC_OFB11'];
                    $oldrxno=$row2['TC_OFB11'];
                  } else {
                    $rxno=$row2['TC_OFB11'];
                    $ii='';
                  }


                  if ($row2['TC_OFB05']=="G") {
                      $unit="G";
                  } else  {
                      $unit="Unit";
                      $unittotal+=$row2["TC_OFB12"];
                  }

                  //invoice目前沒有抛轉patient過來 先改由訂單去抓patient
                  $tc_ofbud05=$row2['TC_OFBUD05'];
                  $tc_ofb31=$row2['TC_OFB31'];
                  $soea="select ta_oea003 from oea_file, oga_file where oea01=oga16 and oga01='$tc_ofb31' ";
                  $erp_sqloea = oci_parse($erp_conn2,$soea );
                  oci_execute($erp_sqloea);
                  $rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC) ;
                  //$tc_ofbud05=$rowoea['TA_OEA003'];

                  // Case No. :           $rxno
                  // Patient Name:        TC_OFBUD05
                  // Product Description: TC_OFB06)
                  // Qty                  TC_OFB12
                  // Unit Price           TC_OFB13
                  // Total                TC_OFB14
                  // Tooth Number         TC_OFBUD02
                  // memo                 TC_OFBUD01
                  $tcofbud02='';
                  if ($row2["TC_OFBUD02"]!=''){
                      if ($occud08==2 or $occud08==3) {
                          if (substr($row2["TC_OFBUD02"],0,1)!='U' and substr($row2["TC_OFBUD02"],0,1)!='L' ) {
                              $tcofbud02='#'.changeteethno($row2["TC_OFBUD02"],$occud08);
                          } else {
                              $tcofbud02=$row2['TC_OFBUD02'];
                          }
                      } else {
                          $tcofbud02=$row2['TC_OFBUD02'];
                      }
                  }

                  $osheet ->setCellValue('A'. $y, $ii)
                          ->setCellValue('B'. $y, $rxno . ' ')
                          ->setCellValue('C'. $y, $tc_ofbud05)
                          ->setCellValue('D'. $y, $row2["TC_OFB06"])
                          ->setCellValue('E'. $y, $row2["TC_OFB12"])
                          ->setCellValue('F'. $y, $unit)
                          ->setCellValue('G'. $y, number_format($row2["TC_OFB13"],2,'.',','))
                          ->setCellValue('H'. $y, number_format($row2["TC_OFB14"],2,'.',','))
                          ->setCellValue('I'. $y, $tcofbud02)
                          ->setCellValue('J'. $y, $row2["TC_OFBUD01"]);



                  $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                  $osheet ->getStyle('I'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                  $osheet ->getStyle('J'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                  if (($y%2)==0){
                      $osheet->getStyle('A'.$y.':J'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                  }
                  $y++;
                }
                $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('H'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('I'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('J'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                //total
                $osheet ->setCellValue('A'.$y, 'Total')
                        ->setCellValue('E'.$y, $unittotal)
                        ->setCellValue('F'.$y, 'Units')
                        ->setCellValue('G'.$y, $currency)
                        ->setCellValue('H'.$y, number_format($total,2,'.',','));
                $osheet ->getStyle('F'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                $osheet ->getStyle('G'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('H'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('H'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );

                $osheet ->getStyle('A'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('B'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('C'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('D'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('F'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('G'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('H'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('I'.$y) ->getFont()->setBold(true);

                // Rename sheet
                $osheet ->setTitle('DailyInvoice');
                // Set active sheet index to the first sheet, so Excel opens this as the first sheet
                $objPHPExcel->setActiveSheetIndex(0);

        }
        else if($template=='NNY')
        { //無patient 有齒位

                $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
                $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣

                $objPHPExcel->setActiveSheetIndex(0);
                $osheet = $objPHPExcel->getActiveSheet();

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);

                $y=5;
                $osheet -> setCellValue('A'.$y, "INVOICE NO.:".$row2["TC_OFA01"].  $occtitle)
                        -> setCellValue('E'.$y, "INVOICE DATE:".$bdate);

                $s3="select occ18, occ231, occ232, occ261, occ271 from occ_file where occ01='$occ01'";
                $erp_sql3 = oci_parse($erp_conn2,$s3 );
                oci_execute($erp_sql3);
                $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);
                $tel='';
                if ($row3['occ261']!='') $tel="Tel: " . $row3['occ261'] . "    ";
                if ($row3['occ271']!='') $tel.="Fax:" . $row3['occ271'];

                $y=6;
                $osheet -> setCellValue('A'.$y, $row3["OCC18"]);
                $y=7;
                $osheet -> setCellValue('A'.$y, $row3["OCC231"]);
                $y=8;
                $osheet -> setCellValue('A'.$y, $row3["OCC232"]);
                $y=9;
                $osheet -> setCellValue('A'.$y, $tel);

                $y=12;
                //$currency=$row2["TC_OFA23"];
                $i=0;
                $oldrxno='';
                $total=0;

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
                  $total+=$row2["TC_OFB14"];
                  if ( $row2["TC_OFB11"]!=$oldrxno) {
                    $i++;
                    $ii=$i;
                    $rxno =$row2['TC_OFB11'];
                    $oldrxno=$row2['TC_OFB11'];
                  } else {
                    $rxno=$row2['TC_OFB11'];
                    $ii='';
                  }


                  if ($row2['TC_OFB05']=="G") {
                      $unit="G";
                  } else  {
                      $unit="Unit";
                      $unittotal+=$row2["TC_OFB12"];
                  }

                  // Case No. :           $rxno
                  // Patient Name:        TC_OFBUD05
                  // Product Description: TC_OFB06)
                  // Qty                  TC_OFB12
                  // Unit Price           TC_OFB13
                  // Total                TC_OFB14
                  // Tooth Number         TC_OFBUD02
                  // memo                 TC_OFBUD01
                  $tcofbud02='';
                  if ($row2["TC_OFBUD02"]!=''){
                      if ($occud08==2 or $occud08==3) {
                          if (substr($row2["TC_OFBUD02"],0,1)!='U' and substr($row2["TC_OFBUD02"],0,1)!='L' ) {
                              $tcofbud02='#'.changeteethno($row2["TC_OFBUD02"],$occud08);
                          } else {
                              $tcofbud02=$row2['TC_OFBUD02'];
                          }
                      } else {
                          $tcofbud02=$row2['TC_OFBUD02'];
                      }
                  }
                  $osheet ->setCellValue('A'. $y, $ii)
                          ->setCellValue('B'. $y, $rxno . ' ')
                          ->setCellValue('C'. $y, $row2["TC_OFB06"])
                          ->setCellValue('D'. $y, $row2["TC_OFB12"])
                          ->setCellValue('E'. $y, $unit)
                          ->setCellValue('F'. $y, number_format($row2["TC_OFB13"],2,'.',','))
                          ->setCellValue('G'. $y, number_format($row2["TC_OFB14"],2,'.',','))
                          ->setCellValue('H'. $y, $tcofbud02)
                          ->setCellValue('I'. $y, $row2["TC_OFBUD01"]);
                  $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                  $osheet ->getStyle('H'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                  $osheet ->getStyle('I'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                  if (($y%2)==0){
                      $osheet->getStyle('A'.$y.':I'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                  }
                  $y++;
                }
                $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('H'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('I'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                //total
                $osheet ->setCellValue('A'.$y, 'Total')
                        ->setCellValue('D'.$y, $unittotal)
                        ->setCellValue('E'.$y, 'Units')
                        ->setCellValue('F'.$y, $currency)
                        ->setCellValue('G'.$y, number_format($total,2,'.',','));
                $osheet ->getStyle('F'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('G'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );

                $osheet ->getStyle('A'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('B'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('C'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('D'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('F'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('G'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('H'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('I'.$y) ->getFont()->setBold(true);

                // Rename sheet
                $osheet ->setTitle('DailyInvoice');
                // Set active sheet index to the first sheet, so Excel opens this as the first sheet
                $objPHPExcel->setActiveSheetIndex(0);

        }
        else if($template=='3YN')
        { //3YN 指的是 PLS的NAT和WAT invoice  自2015/01/07日起 加印patient

                $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
                $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣

                $objPHPExcel->setActiveSheetIndex(0);
                $osheet = $objPHPExcel->getActiveSheet();

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);

                $y=5;
                $osheet -> setCellValue('A'.$y, "INVOICE NO.:".$row2["TC_OFA01"].  $occtitle)
                        -> setCellValue('G'.$y, "INVOICE DATE:".$bdate);

                $s3="select occ18, occ231, occ232, occ261, occ271 from occ_file where occ01='$occ01'";
                $erp_sql3 = oci_parse($erp_conn2,$s3 );
                oci_execute($erp_sql3);
                $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);
                $tel='';
                if ($row3['occ261']!='') $tel="Tel: " . $row3['occ261'] . "    ";
                if ($row3['occ271']!='') $tel.="Fax:" . $row3['occ271'];

                $y=6;
                $osheet -> setCellValue('A'.$y, $row3["OCC18"]);
                $y=7;
                $osheet -> setCellValue('A'.$y, $row3["OCC231"]);
                $y=8;
                $osheet -> setCellValue('A'.$y, $row3["OCC232"]);
                $y=9;
                $osheet -> setCellValue('A'.$y, $tel);

                $y=12;
                //$currency=$row2["TC_OFA23"];
                $i=0;
                $oldrxno='';
                $total=0;

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
                  $total+=$row2["TC_OFB14"];
                  if ( $row2["TC_OFB11"]!=$oldrxno) {
                    $i++;
                    $ii=$i;
                    $rxno =$row2['TC_OFB11'];
                    $oldrxno=$row2['TC_OFB11'];
                  } else {
                    $rxno=$row2['TC_OFB11'];
                    $ii='';
                  }

                   //invoice目前沒有抛轉patient過來 先改由訂單去抓patient
                  $tc_ofbud05=$row2['TC_OFBUD05'];
                  if ($row2['TC_OFB05']=="G") {
                      $unit="G";
                  } else  {
                      $unit="Unit";
                      $unittotal+=$row2["TC_OFB12"];
                  }

                  // Case No. :           $rxno
                  // Patient Name:        TC_OFBUD05
                  // Product Description: TC_OFB06)
                  // Qty                  TC_OFB12
                  // Unit Price           TC_OFB13
                  // Total                TC_OFB14
                  // Tooth Number         TC_OFBUD02
                  // memo                 TC_OFBUD01

                  $osheet ->setCellValue('A'. $y, $ii)
                          ->setCellValue('B'. $y, $rxno . ' ')
                          ->setCellValue('C'. $y, $row2["TC_OFBUD05"])
                          ->setCellValue('D'. $y, $row2["IMAUD04"])
                          ->setCellValue('E'. $y, $row2["TC_OFB06"])
                          ->setCellValue('F'. $y, $row2["TC_OFB12"])
                          ->setCellValue('G'. $y, $unit)
                          ->setCellValue('H'. $y, number_format($row2["TC_OFB13"],2,'.',','))
                          ->setCellValue('I'. $y, number_format($row2["TC_OFB14"],2,'.',','))
                          ->setCellValue('J'. $y, $row2["TC_OFBUD01"]);
                  $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                  $osheet ->getStyle('J'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                  if (($y%2)==0){
                      $osheet->getStyle('A'.$y.':J'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                  }
                  $y++;
                }
                $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('H'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('I'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('J'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                //total
                $osheet ->setCellValue('A'.$y, 'Total')
                        ->setCellValue('F'.$y, $unittotal)
                        ->setCellValue('G'.$y, 'Units')
                        ->setCellValue('H'.$y, $currency)
                        ->setCellValue('I'.$y, number_format($total,2,'.',','));
                $osheet ->getStyle('H'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('I'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );

                $osheet ->getStyle('A'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('B'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('C'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('D'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('F'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('G'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('H'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('I'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('J'.$y) ->getFont()->setBold(true);

                // Rename sheet
                $osheet ->setTitle('DailyInvoice');
                // Set active sheet index to the first sheet, so Excel opens this as the first sheet
                $objPHPExcel->setActiveSheetIndex(0);

        }
        else if($template=='3NN')
        { //3NN 指的是 PLS的invoice
                // 自20160108起 PLS要加一個sheet 列印order和invoice不同的地方
                $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
                $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣

                $objPHPExcel->setActiveSheetIndex(0);
                $osheet = $objPHPExcel->getActiveSheet();

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);

                $y=5;
                $osheet -> setCellValue('A'.$y, "INVOICE NO.:".$row2["TC_OFA01"].  $occtitle)
                        -> setCellValue('F'.$y, "INVOICE DATE:".$bdate);

                $s3="select occ18, occ231, occ232, occ261, occ271 from occ_file where occ01='$occ01'";
                $erp_sql3 = oci_parse($erp_conn2,$s3 );
                oci_execute($erp_sql3);
                $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);
                $tel='';
                if ($row3['occ261']!='') $tel="Tel: " . $row3['occ261'] . "    ";
                if ($row3['occ271']!='') $tel.="Fax:" . $row3['occ271'];

                $y=6;
                $osheet -> setCellValue('A'.$y, $row3["OCC18"]);
                $y=7;
                $osheet -> setCellValue('A'.$y, $row3["OCC231"]);
                $y=8;
                $osheet -> setCellValue('A'.$y, $row3["OCC232"]);
                $y=9;
                $osheet -> setCellValue('A'.$y, $tel);

                $y=12;
                //$currency=$row2["TC_OFA23"];
                $i=0;
                $oldrxno='';
                $total=0;

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
                  $total+=$row2["TC_OFB14"];
                  if ( $row2["TC_OFB11"]!=$oldrxno) {
                    $i++;
                    $ii=$i;
                    $rxno =$row2['TC_OFB11'];
                    $oldrxno=$row2['TC_OFB11'];
                  } else {
                    $rxno=$row2['TC_OFB11'];
                    $ii='';
                  }


                  if ($row2['TC_OFB05']=="G") {
                      $unit="G";
                  } else  {
                      $unit="Unit";
                      $unittotal+=$row2["TC_OFB12"];
                  }

                  // Case No. :           $rxno
                  // Patient Name:        TC_OFBUD05
                  // Product Description: TC_OFB06)
                  // Qty                  TC_OFB12
                  // Unit Price           TC_OFB13
                  // Total                TC_OFB14
                  // Tooth Number         TC_OFBUD02
                  // memo                 TC_OFBUD01

                  $osheet ->setCellValue('A'. $y, $ii)
                          ->setCellValue('B'. $y, $rxno . ' ')
                          ->setCellValue('C'. $y, $row2["IMAUD04"])
                          ->setCellValue('D'. $y, $row2["TC_OFB06"])
                          ->setCellValue('E'. $y, $row2["TC_OFB12"])
                          ->setCellValue('F'. $y, $unit)
                          ->setCellValue('G'. $y, number_format($row2["TC_OFB13"],2,'.',','))
                          ->setCellValue('H'. $y, number_format($row2["TC_OFB14"],2,'.',','))
                          ->setCellValue('I'. $y, $row2["TC_OFBUD01"]);
                  $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                  $osheet ->getStyle('I'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                  if (($y%2)==0){
                      $osheet->getStyle('A'.$y.':I'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                  }
                  $y++;
                }
                $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('H'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('I'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                //total
                $osheet ->setCellValue('A'.$y, 'Total')
                        ->setCellValue('E'.$y, $unittotal)
                        ->setCellValue('F'.$y, 'Units')
                        ->setCellValue('G'.$y, $currency)
                        ->setCellValue('H'.$y, number_format($total,2,'.',','));
                $osheet ->getStyle('G'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('H'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );

                $osheet ->getStyle('A'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('B'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('C'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('D'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('F'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('G'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('H'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('I'.$y) ->getFont()->setBold(true);

                // Rename sheet
                $osheet ->setTitle('DailyInvoice');



                $objPHPExcel->setActiveSheetIndex(1);
                $osheet = $objPHPExcel->getActiveSheet();

                //新增第二個sheet

                $s3="select tc_ofa01, tc_ofb03, tc_ofb11, tc_ofb04, tc_ofb06, tc_ofb12,  tc_ofb31, tc_ofb32, tc_ofbud02,  imaud04 " .
                    "from tc_ofa_file, tc_ofb_file, ima_file " .
                    "where tc_ofaconf='Y' and  tc_ofa02=to_date('$bdate1','yy/mm/dd') and tc_ofa04='$occ01' and tc_ofa01=tc_ofb01 and tc_ofb04=ima01 $invoicefilter " .
                    "and ta_ima003!='Y' " .
                    "order by tc_ofa01,tc_ofb11,tc_ofb08,tc_ofb04" ;

                $i=1;
                $y=2;
                $erp_sql3 = oci_parse($erp_conn2,$s3 );
                oci_execute($erp_sql3);
                while ($row3 = oci_fetch_array($erp_sql3, OCI_ASSOC)) {
                    $tc_ofb31=$row3['TC_OFB31'];
                    $tc_ofb32=$row3['TC_OFB32'];
                    if (is_null($tc_ofb31)) {
                        $osheet ->setCellValue('A'. $y, $i)
                                ->setCellValue('B'. $y, $row3['TC_OFB11'] . " ")
                                ->setCellValue('C'. $y, $row3["IMAUD04"])
                                ->setCellValue('D'. $y, $row3["TC_OFB06"])
                                ->setCellValue('E'. $y, $row3["TC_OFB12"])
                                ->setCellValue('F'. $y, 0);
                        $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                        if (($y%2)==0){
                            $osheet->getStyle('A'.$y.':F'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                        }
                        $i++;
                        $y++;
                    } else {
                        $s4="select oeb12 from oeb_file, ogb_file where oeb01=ogb31 and oeb03=ogb32 and ogb01='$tc_ofb31' and ogb03='$tc_ofb32' ";
                        $erp_sql4 = oci_parse($erp_conn2,$s4 );
                        oci_execute($erp_sql4);
                        $row4 = oci_fetch_array($erp_sql4, OCI_ASSOC);
                        if ($row3['TC_OFB12']!=$row4['OEB12']) {
                            $osheet ->setCellValue('A'. $y, $i)
                                    ->setCellValue('B'. $y, $row3['TC_OFB11'] . " ")
                                    ->setCellValue('C'. $y, $row3["IMAUD04"])
                                    ->setCellValue('D'. $y, $row3["TC_OFB06"])
                                    ->setCellValue('E'. $y, $row3["TC_OFB12"])
                                    ->setCellValue('F'. $y, $row4['OEB12']);
                            $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                            if (($y%2)==0){
                                $osheet->getStyle('A'.$y.':F'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                            }
                            $i++;
                            $y++;
                        }
                    }
                }
                $osheet ->setCellValue('A'. $y, '--- end of record ---');

                // Set active sheet index to the first sheet, so Excel opens this as the first sheet
                $objPHPExcel->setActiveSheetIndex(0);

        }
        else if($template=='4YN')
        { //4YN 指的是 CNC-CEN 的 invoice  自2015/05/20日起 加印Reference no 放在 ta_oea001中

                $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
                $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣

                $objPHPExcel->setActiveSheetIndex(0);
                $osheet = $objPHPExcel->getActiveSheet();

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);

                $y=5;
                $osheet -> setCellValue('A'.$y, "INVOICE NO.:".$row2["TC_OFA01"].  $occtitle)
                        -> setCellValue('H'.$y, "INVOICE DATE:".$bdate);

                $s3="select occ18, occ231, occ232, occ261, occ271 from occ_file where occ01='$occ01'";
                $erp_sql3 = oci_parse($erp_conn2,$s3 );
                oci_execute($erp_sql3);
                $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);
                $tel='';
                if ($row3['occ261']!='') $tel="Tel: " . $row3['occ261'] . "    ";
                if ($row3['occ271']!='') $tel.="Fax:" . $row3['occ271'];

                $y=6;
                $osheet -> setCellValue('A'.$y, $row3["OCC18"]);
                $y=7;
                $osheet -> setCellValue('A'.$y, $row3["OCC231"]);
                $y=8;
                $osheet -> setCellValue('A'.$y, $row3["OCC232"]);
                $y=9;
                $osheet -> setCellValue('A'.$y, $tel);

                $y=12;
                //$currency=$row2["TC_OFA23"];
                $i=0;
                $oldrxno='';
                $total=0;

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
                  $total+=$row2["TC_OFB14"];
                  if ( $row2["TC_OFB11"]!=$oldrxno) {
                    $i++;
                    $ii=$i;
                    $rxno =$row2['TC_OFB11'];
                    $oldrxno=$row2['TC_OFB11'];

                    //取ta_oea001 Clinic欄位 放的是reference no.
                    $soea="select ta_oea001 from oea_file, oga_file where oea01=oga16 and oga01='". $row2['TC_OFB31']. "' ";
                    $erp_sqloea = oci_parse($erp_conn2,$soea );
                    oci_execute($erp_sqloea);
                    $rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC) ;
                    $clinic=$rowoea['TA_OEA001'];
                  } else {
                    $rxno=$row2['TC_OFB11'];
                    $ii='';
                    $clinic='';
                  }

                  $tc_ofbud05=$row2['TC_OFBUD05'];
                  if ($row2['TC_OFB05']=="G") {
                      $unit="G";
                  } else  {
                      $unit="Unit";
                      $unittotal+=$row2["TC_OFB12"];
                  }



                  // Case No. :           $rxno
                  // Patient Name:        TC_OFBUD05
                  // Product Description: TC_OFB06)
                  // Qty                  TC_OFB12
                  // Unit Price           TC_OFB13
                  // Total                TC_OFB14
                  // Tooth Number         TC_OFBUD02
                  // memo                 TC_OFBUD01

                  $osheet ->setCellValue('A'. $y, $ii)
                          ->setCellValue('B'. $y, $rxno . ' ')
                          ->setCellValue('C'. $y, $row2["TC_OFBUD05"])
                          ->setCellValue('D'. $y, $clinic)
                          ->setCellValue('E'. $y, $row2["IMAUD04"])
                          ->setCellValue('F'. $y, $row2["TC_OFB06"])
                          ->setCellValue('G'. $y, $row2["TC_OFB12"])
                          ->setCellValue('H'. $y, $unit)
                          ->setCellValue('I'. $y, number_format($row2["TC_OFB13"],2,'.',','))
                          ->setCellValue('J'. $y, number_format($row2["TC_OFB14"],2,'.',','))
                          ->setCellValue('K'. $y, $row2["TC_OFBUD01"]);
                  $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                  $osheet ->getStyle('K'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                  if (($y%2)==0){
                      $osheet->getStyle('A'.$y.':K'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                  }
                  $y++;
                }
                $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('H'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('I'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('J'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('K'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                //total
                $osheet ->setCellValue('A'.$y, 'Total')
                        ->setCellValue('G'.$y, $unittotal)
                        ->setCellValue('H'.$y, 'Units')
                        ->setCellValue('I'.$y, $currency)
                        ->setCellValue('J'.$y, number_format($total,2,'.',','));
                $osheet ->getStyle('I'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('J'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );

                $osheet ->getStyle('A'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('B'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('C'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('D'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('F'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('G'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('H'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('I'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('J'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('K'.$y) ->getFont()->setBold(true);

                // Rename sheet
                $osheet ->setTitle('DailyInvoice');
                // Set active sheet index to the first sheet, so Excel opens this as the first sheet
                $objPHPExcel->setActiveSheetIndex(0);

        }
        else if($template=='JD')
        { //有patient 有齒位, 有clinic F/HP要加others

            $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
            $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
            $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
            $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
            $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
            $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣

            $objPHPExcel->setActiveSheetIndex(0);
            $osheet = $objPHPExcel->getActiveSheet();

            $erp_sql2 = oci_parse($erp_conn2,$s2 );
            oci_execute($erp_sql2);
            $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);

            $y=5;
            $osheet -> setCellValue('A'.$y, "INVOICE NO.:".$row2["TC_OFA01"].  $occtitle)
                    -> setCellValue('G'.$y, "INVOICE DATE:".$bdate);

            $s3="select occ18, occ231, occ232, occ261, occ271 from occ_file where occ01='$occ01'";
            $erp_sql3 = oci_parse($erp_conn2,$s3 );
            oci_execute($erp_sql3);
            $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);
            $tel='';
            if ($row3['occ261']!='') $tel="Tel: " . $row3['occ261'] . "    ";
            if ($row3['occ271']!='') $tel.="Fax:" . $row3['occ271'];

            $y=6;
            $osheet -> setCellValue('A'.$y, $row3["OCC18"]);
            $y=7;
            $osheet -> setCellValue('A'.$y, $row3["OCC231"]);
            $y=8;
            $osheet -> setCellValue('A'.$y, $row3["OCC232"]);
            $y=9;
            $osheet -> setCellValue('A'.$y, $tel);

            $y=12;
            //$currency=$row2["TC_OFA23"];
            $i=0;
            $oldrxno='';
            $total=0;

            $erp_sql2 = oci_parse($erp_conn2,$s2 );
            oci_execute($erp_sql2);
            while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {


                $total+=$row2["TC_OFB14"];
                if ( $row2["TC_OFB11"]!=$oldrxno) {
                    $i++;
                    $ii=$i;
                    $rxno =$row2['TC_OFB11'];
                    $oldrxno=$row2['TC_OFB11'];
                } else {
                    $rxno=$row2['TC_OFB11'];
                    $ii='';
                }


                if ($row2['TC_OFB05']=="G") {
                    $unit="G";
                } else  {
                    $unit="Unit";
                    $unittotal+=$row2["TC_OFB12"];
                }

                //由訂單去抓clinic, others
                $tc_ofbud05=$row2['TC_OFBUD05'];
                $tc_ofb31=$row2['TC_OFB31'];
                $soea="select ta_oea001, ta_oea003  from oea_file, oga_file where oea01=oga16 and oga01='$tc_ofb31' ";
                $erp_sqloea = oci_parse($erp_conn2,$soea );
                oci_execute($erp_sqloea);
                $rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC) ;
                //$tc_ofbud05=$rowoea['TA_OEA003'];

                // Case No. :           $rxno
                // Patient Name:        TC_OFBUD05
                // Product Description: TC_OFB06)
                // Qty                  TC_OFB12
                // Unit Price           TC_OFB13
                // Total                TC_OFB14
                // Tooth Number         TC_OFBUD02
                // memo                 TC_OFBUD01
                $tcofbud02='';
                if ($row2["TC_OFBUD02"]!=''){
                    if ($occud08==2 or $occud08==3) {
                        if (substr($row2["TC_OFBUD02"],0,1)!='U' and substr($row2["TC_OFBUD02"],0,1)!='L' ) {
                            $tcofbud02='#'.changeteethno($row2["TC_OFBUD02"],$occud08);
                        } else {
                            $tcofbud02=$row2['TC_OFBUD02'];
                        }
                    } else {
                        $tcofbud02=$row2['TC_OFBUD02'];
                    }
                }

                $osheet ->setCellValue('A'. $y, $ii)
                    ->setCellValue('B'. $y, $rxno . ' ')
                    ->setCellValue('C'. $y, $rowoea["TA_OEA003"])
                    ->setCellValue('D'. $y, $rowoea["TA_OEA001"])
                    ->setCellValue('E'. $y, $row2["TC_OFB06"])
                    ->setCellValue('F'. $y, $row2["TC_OFB12"])
                    ->setCellValue('G'. $y, $unit)
                    ->setCellValue('H'. $y, number_format($row2["TC_OFB13"],2,'.',','))
                    ->setCellValue('I'. $y, number_format($row2["TC_OFB14"],2,'.',','))
                    ->setCellValue('J'. $y, $tcofbud02)
                    ->setCellValue('K'. $y, $row2["TC_OFBUD01"]);



                $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                $osheet ->getStyle('J'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                $osheet ->getStyle('K'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                if (($y%2)==0){
                    $osheet->getStyle('A'.$y.':K'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                }
                $y++;
            }
            $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('H'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('I'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('J'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('K'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            //total
            $osheet ->setCellValue('A'.$y, 'Total')
                ->setCellValue('F'.$y, $unittotal)
                ->setCellValue('G'.$y, 'Units')
                ->setCellValue('H'.$y, $currency)
                ->setCellValue('I'.$y, number_format($total,2,'.',','));
            $osheet ->getStyle('G'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            $osheet ->getStyle('H'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $osheet ->getStyle('I'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );

            $osheet ->getStyle('A'.$y) ->getFont()->setBold(true);
            $osheet ->getStyle('B'.$y) ->getFont()->setBold(true);
            $osheet ->getStyle('C'.$y) ->getFont()->setBold(true);
            $osheet ->getStyle('D'.$y) ->getFont()->setBold(true);
            $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);
            $osheet ->getStyle('F'.$y) ->getFont()->setBold(true);
            $osheet ->getStyle('G'.$y) ->getFont()->setBold(true);
            $osheet ->getStyle('H'.$y) ->getFont()->setBold(true);
            $osheet ->getStyle('I'.$y) ->getFont()->setBold(true);
            $osheet ->getStyle('J'.$y) ->getFont()->setBold(true);
            $osheet ->getStyle('K'.$y) ->getFont()->setBold(true);

            // Rename sheet
            $osheet ->setTitle('DailyInvoice');
            // Set active sheet index to the first sheet, so Excel opens this as the first sheet
            $objPHPExcel->setActiveSheetIndex(0);


        }
        else if($template=='TA')
        { //有patient 有齒位, 有clinic F/HP要加others

            $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
            $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
            $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
            $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
            $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
            $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣

            $objPHPExcel->setActiveSheetIndex(0);
            $osheet = $objPHPExcel->getActiveSheet();

            $s3="select occ18, occ231, occ232, occ261, occ271 from occ_file where occ01='$occ01'";
            $erp_sql3 = oci_parse($erp_conn2,$s3 );
            oci_execute($erp_sql3);
            $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);
            $osheet -> setCellValue('A1', "TO:".$row3["OCC18"]);
            $osheet -> setCellValue('A3', "DATE:".$bdate);

            $erp_sql2 = oci_parse($erp_conn2,$s2 );
            oci_execute($erp_sql2);
            $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);

            $osheet -> setCellValue('A5', "TRACKING NO.  " . $row2['TC_OFAUD03']);

            $y=7;
            $i=0;
            $oldrxno='';
            $total=0;

            $erp_sql2 = oci_parse($erp_conn2,$s2 );
            oci_execute($erp_sql2);

            while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {

                $total+=$row2["TC_OFB14"];
                if ( $row2["TC_OFB11"]!=$oldrxno) {
                    $i++;
                    $ii=$i;
                    $rxno =$row2['TC_OFB11'];
                    $oldrxno=$row2['TC_OFB11'];
                } else {
                    $rxno=$row2['TC_OFB11'];
                    $ii='';
                }


                if ($row2['TC_OFB05']=="G") {
                    $unit="G";
                } else  {
                    $unit="Unit";
                    $unittotal+=$row2["TC_OFB12"];
                }

                //由訂單去抓clinic, others
                $tc_ofbud05=$row2['TC_OFBUD05'];  // patient
                $tc_ofb31=$row2['TC_OFB31'];   // order no
                $soea="select ta_oea001, ta_oea003  from oea_file, oga_file where oea01=oga16 and oga01='$tc_ofb31' ";
                $erp_sqloea = oci_parse($erp_conn2,$soea );
                oci_execute($erp_sqloea);
                $rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC) ;
                //$tc_ofbud05=$rowoea['TA_OEA003'];

                // rx:                  TC_OFB11
                // Case No. :           $rxno
                // Patient Name:        TC_OFBUD05
                // Product Description: TC_OFB06)
                // Qty                  TC_OFB12
                // Unit Price           TC_OFB13
                // Total                TC_OFB14
                // Tooth Number         TC_OFBUD02
                // memo                 TC_OFBUD01
                $tcofbud02='';
                if ($row2["TC_OFBUD02"]!=''){  //teeth
                    if ($occud08==2 or $occud08==3) {
                        if (substr($row2["TC_OFBUD02"],0,1)!='U' and substr($row2["TC_OFBUD02"],0,1)!='L' ) {
                            $tcofbud02='#'.changeteethno($row2["TC_OFBUD02"],$occud08);
                        } else {
                            $tcofbud02=$row2['TC_OFBUD02'];
                        }
                    } else {
                        $tcofbud02=$row2['TC_OFBUD02'];
                    }
                }


                $osheet ->setCellValue('A'. $y, $rxno . ' ')
                        ->setCellValue('E'. $y, $row2["TC_OFBUD05"])
                        ->setCellValue('F'. $y, $tcofbud02)
                        ->setCellValue('P'. $y, $row2["TC_OFBUD01"]);


                if ($row2['TC_OFB05']=="G") {
                    $osheet ->setCellValue('K'. $y, $row2["TC_OFB06"])
                            ->setCellValue('L'. $y, number_format($row2["TC_OFB12"],2,'.',','))
                            ->setCellValue('M'. $y, number_format($row2["TC_OFB13"],2,'.',','))
                            ->setCellValue('N'. $y, number_format($row2["TC_OFB14"],2,'.',','))
                            ->setCellValue('Q'. $y, number_format($row2["TC_OFB14"],2,'.',','));
                } else {
                    $osheet ->setCellValue('G'. $y, $row2["TC_OFB06"])
                            ->setCellValue('H'. $y, number_format($row2["TC_OFB12"],2,'.',','))
                            ->setCellValue('I'. $y, number_format($row2["TC_OFB13"],2,'.',','))
                            ->setCellValue('J'. $y, number_format($row2["TC_OFB14"],2,'.',','))
                            ->setCellValue('Q'. $y, number_format($row2["TC_OFB14"],2,'.',','));
                }

                $osheet ->getStyle('I'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
                $osheet ->getStyle('J'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
                $osheet ->getStyle('M'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
                $osheet ->getStyle('N'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
                $osheet ->getStyle('Q'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );

                $osheet ->getStyle('A'.$y)  ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('A'.$y)  ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('A'.$y)  ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('A'.$y)  ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

                $osheet ->getStyle('B'.$y)  ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('B'.$y)  ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('B'.$y)  ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('B'.$y)  ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

                $osheet ->getStyle('C'.$y)  ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('C'.$y)  ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('C'.$y)  ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('C'.$y)  ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

                $osheet ->getStyle('D'.$y)  ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('D'.$y)  ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('D'.$y)  ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('D'.$y)  ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

                $osheet ->getStyle('E'.$y)  ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('E'.$y)  ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('E'.$y)  ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('E'.$y)  ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

                $osheet ->getStyle('F'.$y)  ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('F'.$y)  ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('F'.$y)  ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('F'.$y)  ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

                $osheet ->getStyle('G'.$y)  ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('G'.$y)  ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('G'.$y)  ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('G'.$y)  ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

                $osheet ->getStyle('H'.$y)  ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('H'.$y)  ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('H'.$y)  ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('H'.$y)  ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

                $osheet ->getStyle('I'.$y)  ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('I'.$y)  ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('I'.$y)  ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('I'.$y)  ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

                $osheet ->getStyle('J'.$y)  ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('J'.$y)  ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('J'.$y)  ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('J'.$y)  ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

                $osheet ->getStyle('K'.$y)  ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('K'.$y)  ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('K'.$y)  ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('K'.$y)  ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

                $osheet ->getStyle('L'.$y)  ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('L'.$y)  ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('L'.$y)  ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('L'.$y)  ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

                $osheet ->getStyle('M'.$y)  ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('M'.$y)  ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('M'.$y)  ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('M'.$y)  ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

                $osheet ->getStyle('N'.$y)  ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('N'.$y)  ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('N'.$y)  ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('N'.$y)  ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

                $osheet ->getStyle('O'.$y)  ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('O'.$y)  ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('O'.$y)  ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('O'.$y)  ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);


                $osheet ->getStyle('P'.$y)  ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('P'.$y)  ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('P'.$y)  ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('P'.$y)  ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);


                $osheet ->getStyle('Q'.$y)  ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('Q'.$y)  ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('Q'.$y)  ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('Q'.$y)  ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $y++;
            }

            //total
            $osheet ->setCellValue('Q'.$y, number_format($total,2,'.',','));
            $osheet ->getStyle('Q'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
            $osheet ->getStyle('Q'.$y) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('Q'.$y) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('Q'.$y) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('Q'.$y) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

            $osheet -> setCellValue('A2', "THE FOLLOWING " . $i .  " CASES HAD BEEN SENT ON");
            // Rename sheet
            $osheet ->setTitle('DailyInvoice');
            // Set active sheet index to the first sheet, so Excel opens this as the first sheet
            $objPHPExcel->setActiveSheetIndex(0);
        }
        else if($template=='U204001')
        { //有patient 有齒位, 有clinic F/HP要加others

            $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
            $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
            $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
            $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
            $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
            $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣

            $objPHPExcel->setActiveSheetIndex(0);
            $osheet = $objPHPExcel->getActiveSheet();

            $erp_sql2 = oci_parse($erp_conn2,$s2 );
            oci_execute($erp_sql2);
            $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);

            $y=5;
            $osheet -> setCellValue('H2', "Invoice No.:".$row2["TC_OFA01"]);
            //全名   發票地1  發票地2  TEL     FAX
            $s3="select occ18, occ231, occ232, occ261, occ271 from occ_file where occ01='$occ01'";
            $erp_sql3 = oci_parse($erp_conn2,$s3 );
            oci_execute($erp_sql3);
            $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);
            $tel='';


            $y=12;
            //$currency=$row2["TC_OFA23"];
            $i=0;
            $oldrxno='';
            $total=0;

            $erp_sql2 = oci_parse($erp_conn2,$s2 );
            oci_execute($erp_sql2);
            while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {

                //由訂單去抓clinic, others
                $tc_ofbud05=$row2['TC_OFBUD05'];
                $tc_ofb31=$row2['TC_OFB31'];
                $soea="select ta_oea001, ta_oea003  from oea_file, oga_file where oea01=oga16 and oga01='$tc_ofb31' ";
                $erp_sqloea = oci_parse($erp_conn2,$soea );
                oci_execute($erp_sqloea);
                $rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC) ;

                $total+=$row2["TC_OFB14"];
                if ( $row2["TC_OFB11"]!=$oldrxno) {
                    $i++;
                    $ii=$i;
                    $rxno =$row2['TC_OFB11'];
                    $oldrxno=$row2['TC_OFB11'];
                    $pan = $rowoea['TA_OEA001'];
                    $patient = $rowoea['TA_OEA003'];
                } else {
                    $rxno='';
                    $pan='';
                    $patient='';
                    $ii='';
                }


                if ($row2['TC_OFB05']=="G") {
                    $unit="G";
                } else  {
                    $unit="Unit";
                    $unittotal+=$row2["TC_OFB12"];
                }


                //$tc_ofbud05=$rowoea['TA_OEA003'];

                // Case No. :           $rxno
                // Patient Name:        TC_OFBUD05
                // Product Description: TC_OFB06)
                // Qty                  TC_OFB12
                // Unit Price           TC_OFB13
                // Total                TC_OFB14
                // Tooth Number         TC_OFBUD02
                // memo                 TC_OFBUD01
                $tcofbud02='';
                if ($row2["TC_OFBUD02"]!=''){
                    if ($occud08==2 or $occud08==3) {
                        if (substr($row2["TC_OFBUD02"],0,1)!='U' and substr($row2["TC_OFBUD02"],0,1)!='L' ) {
                            $tcofbud02='#'.changeteethno($row2["TC_OFBUD02"],$occud08);
                        } else {
                            $tcofbud02=$row2['TC_OFBUD02'];
                        }
                    } else {
                        $tcofbud02=$row2['TC_OFBUD02'];
                    }
                }

                $osheet ->setCellValue('A'. $y, $ii)
                    ->setCellValue('B'. $y, $rxno . ' ')
                    ->setCellValue('C'. $y, $pan . ' ')
                    ->setCellValue('D'. $y, $patient . ' ')
                    ->setCellValue('E'. $y, $row2["TC_OFB06"])
                    ->setCellValue('F'. $y, $tcofbud02)
                    ->setCellValue('G'. $y, $row2["TC_OFB12"])
                    ->setCellValue('H'. $y, $unit)
                    ->setCellValue('I'. $y, number_format($row2["TC_OFB13"],2,'.',','))
                    ->setCellValue('J'. $y, number_format($row2["TC_OFB14"],2,'.',','))
                    ->setCellValue('K'. $y, $row2["TC_OFBUD01"]);


                $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                $osheet ->getStyle('B'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                $osheet ->getStyle('C'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                $osheet ->getStyle('D'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                $osheet ->getStyle('E'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                if (($y%2)==0){
                    $osheet->getStyle('A'.$y.':K'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                }
                $y++;
            }
            $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('H'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('I'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('J'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('J'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            $osheet ->getStyle('K'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
            //total
            $osheet ->setCellValue('A'.$y, 'Total')
                ->setCellValue('G'.$y, $unittotal)
                ->setCellValue('H'.$y, 'Units')
                ->setCellValue('I'.$y, $currency)
                ->setCellValue('J'.$y, number_format($total,2,'.',','));
            $osheet ->getStyle('G'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $osheet ->getStyle('J'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );

            $osheet ->getStyle('A'.$y) ->getFont()->setBold(true);
            $osheet ->getStyle('B'.$y) ->getFont()->setBold(true);
            $osheet ->getStyle('C'.$y) ->getFont()->setBold(true);
            $osheet ->getStyle('D'.$y) ->getFont()->setBold(true);
            $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);
            $osheet ->getStyle('F'.$y) ->getFont()->setBold(true);
            $osheet ->getStyle('G'.$y) ->getFont()->setBold(true);
            $osheet ->getStyle('H'.$y) ->getFont()->setBold(true);
            $osheet ->getStyle('I'.$y) ->getFont()->setBold(true);
            $osheet ->getStyle('J'.$y) ->getFont()->setBold(true);
            $osheet ->getStyle('K'.$y) ->getFont()->setBold(true);

            // Rename sheet
            $osheet ->setTitle('DailyInvoice');
            // Set active sheet index to the first sheet, so Excel opens this as the first sheet
            $objPHPExcel->setActiveSheetIndex(0);

        }
        else if($template=='F' || $template=='HP')
        { //有patient 有齒位, 有clinic F/HP要加others

          $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
          $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
          $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
          $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
          $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
          $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣

          $objPHPExcel->setActiveSheetIndex(0);
          $osheet = $objPHPExcel->getActiveSheet();

          $erp_sql2 = oci_parse($erp_conn2,$s2 );
          oci_execute($erp_sql2);
          $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);

          $y=5;
          $osheet -> setCellValue('A'.$y, "INVOICE NO.:".$row2["TC_OFA01"].  $occtitle)
                  -> setCellValue('G'.$y, "INVOICE DATE:".$bdate);

          $s3="select occ18, occ231, occ232, occ261, occ271 from occ_file where occ01='$occ01'";
          $erp_sql3 = oci_parse($erp_conn2,$s3 );
          oci_execute($erp_sql3);
          $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);
          $tel='';
          if ($row3['occ261']!='') $tel="Tel: " . $row3['occ261'] . "    ";
          if ($row3['occ271']!='') $tel.="Fax:" . $row3['occ271'];

          $y=6;
          $osheet -> setCellValue('A'.$y, $row3["OCC18"]);
          $y=7;
          $osheet -> setCellValue('A'.$y, $row3["OCC231"]);
          $y=8;
          $osheet -> setCellValue('A'.$y, $row3["OCC232"]);
          $y=9;
          $osheet -> setCellValue('A'.$y, $tel);

          $y=12;
          //$currency=$row2["TC_OFA23"];
          $i=0;
          $oldrxno='';
          $total=0;

          $erp_sql2 = oci_parse($erp_conn2,$s2 );
          oci_execute($erp_sql2);
          while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {


            $total+=$row2["TC_OFB14"];
            if ( $row2["TC_OFB11"]!=$oldrxno) {
              $i++;
              $ii=$i;
              $rxno =$row2['TC_OFB11'];
              $oldrxno=$row2['TC_OFB11'];
            } else {
              $rxno=$row2['TC_OFB11'];
              $ii='';
            }


            if ($row2['TC_OFB05']=="G") {
              $unit="G";
            } else  {
              $unit="Unit";
              $unittotal+=$row2["TC_OFB12"];
            }

            //由訂單去抓clinic, others
            $tc_ofbud05=$row2['TC_OFBUD05'];
            $tc_ofb31=$row2['TC_OFB31'];
            $soea="select ta_oea001, ta_oea003, ta_oea061 from oea_file, oga_file where oea01=oga16 and oga01='$tc_ofb31' ";
            $erp_sqloea = oci_parse($erp_conn2,$soea );
            oci_execute($erp_sqloea);
            $rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC) ;
            //$tc_ofbud05=$rowoea['TA_OEA003'];

            // Case No. :           $rxno
            // Patient Name:        TC_OFBUD05
            // Product Description: TC_OFB06)
            // Qty                  TC_OFB12
            // Unit Price           TC_OFB13
            // Total                TC_OFB14
            // Tooth Number         TC_OFBUD02
            // memo                 TC_OFBUD01
            $tcofbud02='';
            if ($row2["TC_OFBUD02"]!=''){
              if ($occud08==2 or $occud08==3) {
                if (substr($row2["TC_OFBUD02"],0,1)!='U' and substr($row2["TC_OFBUD02"],0,1)!='L' ) {
                  $tcofbud02='#'.changeteethno($row2["TC_OFBUD02"],$occud08);
                } else {
                  $tcofbud02=$row2['TC_OFBUD02'];
                }
              } else {
                $tcofbud02=$row2['TC_OFBUD02'];
              }
            }

            $osheet ->setCellValue('A'. $y, $ii)
              ->setCellValue('B'. $y, $rxno . ' ')
              ->setCellValue('C'. $y, $rowoea["TA_OEA003"])
              ->setCellValue('D'. $y, $rowoea["TA_OEA001"])
              ->setCellValue('E'. $y, $row2["TC_OFB06"])
              ->setCellValue('F'. $y, $row2["TC_OFB12"])
              ->setCellValue('G'. $y, $unit)
              ->setCellValue('H'. $y, number_format($row2["TC_OFB13"],2,'.',','))
              ->setCellValue('I'. $y, number_format($row2["TC_OFB14"],2,'.',','))
              ->setCellValue('J'. $y, $tcofbud02)
              ->setCellValue('K'. $y, $row2["TC_OFBUD01"])
              ->setCellValue('L'. $y, $rowoea["TA_OEA061"]);



            $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            $osheet ->getStyle('J'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            $osheet ->getStyle('K'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            $osheet ->getStyle('L'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            if (($y%2)==0){
              $osheet->getStyle('A'.$y.':L'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
            }
            $y++;
          }
          $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          $osheet ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          $osheet ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          $osheet ->getStyle('H'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          $osheet ->getStyle('I'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          $osheet ->getStyle('J'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          $osheet ->getStyle('K'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          $osheet ->getStyle('L'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          //total
          $osheet ->setCellValue('A'.$y, 'Total')
            ->setCellValue('F'.$y, $unittotal)
            ->setCellValue('G'.$y, 'Units')
            ->setCellValue('H'.$y, $currency)
            ->setCellValue('I'.$y, number_format($total,2,'.',','));
          $osheet ->getStyle('G'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
          $osheet ->getStyle('H'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
          $osheet ->getStyle('I'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );

          $osheet ->getStyle('A'.$y) ->getFont()->setBold(true);
          $osheet ->getStyle('B'.$y) ->getFont()->setBold(true);
          $osheet ->getStyle('C'.$y) ->getFont()->setBold(true);
          $osheet ->getStyle('D'.$y) ->getFont()->setBold(true);
          $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);
          $osheet ->getStyle('F'.$y) ->getFont()->setBold(true);
          $osheet ->getStyle('G'.$y) ->getFont()->setBold(true);
          $osheet ->getStyle('H'.$y) ->getFont()->setBold(true);
          $osheet ->getStyle('I'.$y) ->getFont()->setBold(true);
          $osheet ->getStyle('J'.$y) ->getFont()->setBold(true);
          $osheet ->getStyle('K'.$y) ->getFont()->setBold(true);
          $osheet ->getStyle('L'.$y) ->getFont()->setBold(true);

          // Rename sheet
          $osheet ->setTitle('DailyInvoice');
          // Set active sheet index to the first sheet, so Excel opens this as the first sheet
          $objPHPExcel->setActiveSheetIndex(0);

        }
        else
        {  //沒有做任何設置 , 以不印齒位/patient

                $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
                $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣

                $objPHPExcel->setActiveSheetIndex(0);
                $osheet = $objPHPExcel->getActiveSheet();

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);

                $y=5;
                $osheet -> setCellValue('A'.$y, "INVOICE NO.:".$row2["TC_OFA01"].  $occtitle)
                        -> setCellValue('E'.$y, "INVOICE DATE:".$bdate);

                $s3="select occ18, occ231, occ232, occ261, occ271 from occ_file where occ01='$occ01'";
                $erp_sql3 = oci_parse($erp_conn2,$s3 );
                oci_execute($erp_sql3);
                $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);
                $tel='';
                if ($row3['occ261']!='') $tel="Tel: " . $row3['occ261'] . "    ";
                if ($row3['occ271']!='') $tel.="Fax:" . $row3['occ271'];

                $y=6;
                $osheet -> setCellValue('A'.$y, $row3["OCC18"]);
                $y=7;
                $osheet -> setCellValue('A'.$y, $row3["OCC231"]);
                $y=8;
                $osheet -> setCellValue('A'.$y, $row3["OCC232"]);
                $y=9;
                $osheet -> setCellValue('A'.$y, $tel);

                $y=12;
                //$currency=$row2["TC_OFA23"];
                $i=0;
                $oldrxno='';
                $total=0;

                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);
                while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
                  $total+=$row2["TC_OFB14"];
                  if ( $row2["TC_OFB11"]!=$oldrxno) {
                    $i++;
                    $ii=$i;
                    $rxno =$row2['TC_OFB11'];
                    $oldrxno=$row2['TC_OFB11'];
                  } else {
                    $rxno=$row2['TC_OFB11'];
                    $ii='';
                  }


                  if ($row2['TC_OFB05']=="G") {
                      $unit="G";
                  } else  {
                      $unit="Unit";
                      $unittotal+=$row2["TC_OFB12"];
                  }

                  // Case No. :           $rxno
                  // Patient Name:        TC_OFBUD05
                  // Product Description: TC_OFB06)
                  // Qty                  TC_OFB12
                  // Unit Price           TC_OFB13
                  // Total                TC_OFB14
                  // Tooth Number         TC_OFBUD02
                  // memo                 TC_OFBUD01

                  $osheet ->setCellValue('A'. $y, $ii)
                          ->setCellValue('B'. $y, $rxno . ' ')
                          ->setCellValue('C'. $y, $row2["TC_OFB06"])
                          ->setCellValue('D'. $y, $row2["TC_OFB12"])
                          ->setCellValue('E'. $y, $unit)
                          ->setCellValue('F'. $y, number_format($row2["TC_OFB13"],2,'.',','))
                          ->setCellValue('G'. $y, number_format($row2["TC_OFB14"],2,'.',','))
                          ->setCellValue('H'. $y, $row2["TC_OFBUD01"]);
                  $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                  $osheet ->getStyle('H'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                  if (($y%2)==0){
                      $osheet->getStyle('A'.$y.':H'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                  }
                  $y++;
                }
                $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('H'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                //total
                $osheet ->setCellValue('A'.$y, 'Total')
                        ->setCellValue('D'.$y, $unittotal)
                        ->setCellValue('E'.$y, 'Units')
                        ->setCellValue('F'.$y, $currency)
                        ->setCellValue('G'.$y, number_format($total,2,'.',','));
                $osheet ->getStyle('F'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                $osheet ->getStyle('G'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );

                $osheet ->getStyle('A'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('B'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('C'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('D'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('F'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('G'.$y) ->getFont()->setBold(true);
                $osheet ->getStyle('H'.$y) ->getFont()->setBold(true);

                // Rename sheet
                $osheet ->setTitle('DailyInvoice');
                // Set active sheet index to the first sheet, so Excel opens this as the first sheet
                $objPHPExcel->setActiveSheetIndex(0);
        }


        //if ($_GET['submit']=='Excel') {
            // Redirect output to a client’s web browser (Excel5)
            header('Content-Type: application/vnd.ms-excel');
            header('Content-Disposition: attachment;filename="'. 'DailyInvoice_' . $occ01 . '_' .$occ02 .'_' . $bdate . '.xls"');
            header('Cache-Control: max-age=0');
            $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        //} else {
        //    header('Content-Type: application/pdf');
        //    header('Content-Disposition: attachment;filename="'. 'DailyInvoice_' . $occ01 . '_' . $bdate . '.pdf"');
        //    header('Cache-Control: max-age=0');
        //    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'PDF');
        //}
        $objWriter->save('php://output');
        exit;
  }
/*
    //for xajax
  require ('xajax/xajax_core/xajax.inc.php');
  $xajax = new xajax();
  $xajax->configure('javascript URI', 'xajax/');

  $xajax->processRequest();

  echo '<?xml version="1.0" encoding="UTF-8"?>';
  */
  $IsAjax = false;

  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>客戶 Invoice 列印 </p>
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
         出貨日期:
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> &nbsp;&nbsp;
        送貨客戶:
        <select name="occ01" id="occ01">
            <?
              $s1= "select occ01,occ02 from occ_file order by occ01 ";
              $erp_sql1 = oci_parse($erp_conn1,$s1 );
              oci_execute($erp_sql1);
              while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                  echo "<option value=" . $row1["OCC01"];
                  if ($_GET["occ01"] == $row1["OCC01"]) echo " selected";
                  echo ">" . $row1['OCC01'] ."--" .$row1["OCC02"] . "</option>";
              }
            ?>
        </select>
        &nbsp;&nbsp;
        Invoice No.:
        <input type="text" id="invoiceno" name="invoiceno" value=<?=$invoiceno;?> >
        &nbsp;&nbsp;
        <input type="submit" name="submit" id="submit" value="Submit">  &nbsp;&nbsp;   &nbsp;&nbsp;
        <input type="submit" name="submit" id="submit" value="Excel">  &nbsp;&nbsp;   &nbsp;&nbsp;
      </td></tr>
    </table>
  </div>
</form>
<? if (is_null($_GET['submit'])) die ; ?>
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
        <th width="16">&nbsp;</th>
        <th>Invoice No.</th>
        <th>Case No.</th>
        <th>Patient</th>
        <th>Product Code</th>
        <th>Product Description</th>
        <th>Qty.</th>
        <th>Unit</th>
        <th>Price</th>
        <th>Total</th>
        <th>Currency</th>
        <th>Teeth No.</th>
        <th>Memo</th>
    </tr>
    <?
      $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);
      //檢查工單號有無出貨  , 配件不用
      $s2="select tc_ofa01, tc_ofa23, tc_ofb11, tc_ofbud05, tc_ofb08, tc_ofb04, tc_ofb06, tc_ofb12, tc_ofb05, tc_ofb13, tc_ofb14, tc_ofb31, tc_ofbud02, tc_ofbud01 " .
          "from tc_ofa_file, tc_ofb_file where tc_ofaconf='Y' and  tc_ofa02=to_date('$bdate1','yy/mm/dd')  and tc_ofa04='$occ01' and tc_ofa01=tc_ofb01 $invoicefilter order by tc_ofa01,tc_ofb11,tc_ofb08,tc_ofb04 ";
      $erp_sql2 = oci_parse($erp_conn2,$s2 );
      oci_execute($erp_sql2);
      $bgkleur = "ffffff";
      $i=0;
      $oldrxno='';
      $total=0;
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
          $total+=$row2["TC_OFB14"];
          if ( $row2["TC_OFB11"]!=$oldrxno) {
            $i++;
            $ii=$i;
            $rxno =$row2['TC_OFB11'];
            $oldrxno=$row2['TC_OFB11'];
          } else {
            $rxno=$row2['TC_OFB11'];
            $ii='';
          }

          if ($row2['TC_OFB05']=="G") {
              $unit="G";
          } else  {
              $unit="Unit";
          }

          //invoice目前沒有抛轉patient過來 先改由訂單去抓patient
          $tc_ofbud05=$row2['TC_OFBUD05'];
          $tc_ofb31=$row2['TC_OFB31'];
          $soea="select ta_oea003 from oea_file, oga_file where oea01=oga16 and oga01='$tc_ofb31' ";
          $erp_sqloea = oci_parse($erp_conn2,$soea );
          oci_execute($erp_sqloea);
          $rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC) ;
          //$tc_ofbud05=$rowoea['TA_OEA003'];

    ?>
          <tr bgcolor="#<?=$bgkleur;?>">
              <td><?=$ii;?></td>
              <td><?=$row2["TC_OFA01"];?></td>
              <td><?=$rxno;?></td>
              <td><?=$tc_ofbud05;?></td>
              <td><?=$row2["TC_OFB04"];?></td>
              <td><?=$row2["TC_OFB06"];?></td>
              <td style="text-align:right" ><?=number_format($row2["TC_OFB12"],2,'.',',');?></td>
              <td><?=$unit;?></td>
              <td style="text-align:right" ><?=number_format($row2["TC_OFB13"],2,'.',',');?></td>
              <td style="text-align:right" ><?=number_format($row2["TC_OFB14"],2,'.',',');?></td>
              <td><?=$row2["TC_OFA23"];?></td>
              <td><?=$row2["TC_OFBUD02"];?></td>
              <td><?=$row2["TC_OFBUD01"];?></td>
          </tr>
    <?
      }
    ?>
    <tr bgcolor="#<?=$bgkleur;?>">
        <td colspan="9"></td>
        <td  style="text-align:right" ><?=number_format($total,2,'.',',');?></td>
        <td colspan="3">&nbsp;</td>
    </tr>
</table>
