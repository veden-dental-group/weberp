<?php
  session_start();
  $pagetitle = "IT &raquo; 各產品工序報表";
  include("_data.php");

      error_reporting(E_NONE);  
      require_once 'classes/PHPExcel.php'; 
      require_once 'classes/PHPExcel/IOFactory.php';  
      $objPHPExcel = new PHPExcel();

      $objPHPExcel->setActiveSheetIndex(0);
      $o = $objPHPExcel->getActiveSheet();                                         

      $s1="select ecd01, ecd02 from ecd_file order by ecd01";
      $erp_sql1 = oci_parse($erp_conn1,$s1 );
      oci_execute($erp_sql1); 
      $i=0;
      while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {  
        $ecd01[$i]=$row1['ECD01'];
        $z=$i+67;
        if ( $z > 90){ //超過90以後的 都要加一個文字
           $a = chr(floor(($z-65)/26)+64);
           $b = chr((($z-65) % 26)+65) ;
        } else {
           $a="";
           $b=chr($z); 
        }
        $o -> setCellValue($a.$b.'1', $row1['ECD01']);
        $o -> setCellValue($a.$b.'2', $row1['ECD02']);
        $i++;
      }

      $s1="select ecb01, ima02, ecb03, ecb06, ecb17, ecbud07, ecbud08 from ecb_file, ima_file where ecb01=ima01 order by ecb01, ecb03";
      $erp_sql1 = oci_parse($erp_conn1,$s1 );
      oci_execute($erp_sql1); 
      $y=2;
      $ecb01='';
      while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) { 
        if ($ecb01!=$row1['ECB01']) {
           $y++;                           
           $ecb01 = $row1['ECB01'];
        }                      
        $o -> setCellValue('A'.$y, ' '.$row1['ECB01']);
        $o -> setCellValue('B'.$y, $row1['IMA02']);
        for ($j=0;$j<$i;$j++){
            if ($row1['ECB06']==$ecd01[$j]){
                $z= $j + 67 ;
                if ( $z > 90){ //超過90以後的 都要加一個文字
                   $a = 'A';
                   $b = chr((($z-67) % 26)+65) ;
                } else {
                   $a="";
                   $b=chr($z); 
                } 
                if ($row1['ECBUD07']==$row1['ECBUD08']) {
                    $o->setCellValue($a.$b.$y, $row1['ECBUD07']); 
                } else {
                    $o->setCellValue($a.$b.$y, '進:'.$row1['ECBUD07'].' 出:'.$row1['ECBUD08']); 
                }
            }
        }
        
      }
      $o ->setCellValue('A'.($y), '以下無資料');
                                               
      // Set active sheet index to the first sheet, so Excel opens this as the first sheet
      $objPHPExcel->setActiveSheetIndex(0);

      // Redirect output to a client’s web browser (Excel5)
      $objPHPExcel->getActiveSheet()->setTitle('產品工序單價表');
      $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');   
      $filename='report/products_steps_prices.xls';
      $objWriter->save($filename); 