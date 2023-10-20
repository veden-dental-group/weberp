<?php
// echo "here!!";
// echo $argv[1];
// echo $argv[2]; 
//每天計算當天出貨的顆數

// session_start();
// @ini_set('display_errors', 'on');
// error_reporting(E_ALL ^ E_NOTICE);
   
include("_data.php");  
date_default_timezone_set('Asia/Taipei');
$bdate = $argv[1] ? $argv[1] : date('Y-m-d');
$edate = $argv[2] ? $argv[2] : $bdate;
                   
    error_reporting(E_ALL);  
    require_once 'classes/PHPExcel.php'; 
    require_once 'classes/PHPExcel/IOFactory.php';  
    //$objPHPExcel = new PHPExcel();
    $objReader = PHPExcel_IOFactory::createReader('Excel5');  
    $objPHPExcel = $objReader->load("templates/erp_sentcases_4mv3.xls");  
    // Set properties
    $objPHPExcel ->getProperties()->setCreator("Frank")
               ->setLastModifiedBy("Frank")
               ->setTitle("$edate 出貨顆/床數統計")
               ->setSubject("$edate 出貨顆/床數統計")
               ->setDescription("$edate 出貨顆/床數統計")
               ->setKeywords("$edate 出貨顆/床數統計")
               ->setCategory("$edate 出貨顆/床數統計");

    // Add some data
    $objPHPExcel->setActiveSheetIndex(0)
                ->setCellValue('A1', $bdate. ' -- '. $edate. ' 出貨顆/床數統計'); 
    $objPHPExcel->setActiveSheetIndex(0) ->mergeCells('A1:G1');  
                  
    $month=substr($edate,0,7);                     

    //計算當日出貨顆/床數         
    $query = "select sum(s690000) s690000, sum(s6A1000) s6A1000, sum(s6A2000) s6A2000, sum(s6A3000) s6A3000, sum(s6A4100) s6A4100, sum(s6A4200) s6A4200, sum(s6A5000) s6A5000, sum(s6A6000) s6A6000 " .
              "from " .
              "(select if(s13='690000',round(s25*s27,2),0) s690000, if(s13='6A1000',round(s25*s27,2),0) s6A1000, if(s13='6A2000',round(s25*s27,2),0) s6A2000, if(s13='6A3000',round(s25*s27,2),0) s6A3000, " .
              "if(s13='6A5000',round(s25*s27,2),0) s6A5000, if(s13='6A6000',round(s25*s27,2),0) s6A6000, if(s13='6A4300',round(s25*s27,2),0) s6A4100, if(s13='6A4300',round(s25*s28,2),0) s6A4200 " .
              "from " .
              "(select s09, if(substr(s09,1,1)='1', s13, if(substr(s09,1,2)='27', s13, '6A4300')) s13 , s25, s26, s27, s28 from casesent where s01='$edate' ) a ) b ";

    $result= mysql_query($query);            
    $row   = mysql_fetch_array($result);
    $objPHPExcel->setActiveSheetIndex(0)     
                ->setCellValue('C4', $row['s690000'])
                ->setCellValue('C5', $row['s6A1000'])   
                ->setCellValue('C6', $row['s6A2000'])   
                ->setCellValue('C7', $row['s6A3000']) 
                ->setCellValue('C8', $row['s6A4100'])   
                ->setCellValue('C9', $row['s6A4200'])
                ->setCellValue('C10',$row['s6A5000'])
                ->setCellValue('C11',$row['s6A6000']) ;
                
    //計算當月合計出貨顆/床數         
    $query = "select sum(s690000) s690000, sum(s6A1000) s6A1000, sum(s6A2000) s6A2000, sum(s6A3000) s6A3000, sum(s6A4100) s6A4100, sum(s6A4200) s6A4200, sum(s6A5000) s6A5000, sum(s6A6000) s6A6000 " .
              "from " .
              "(select if(s13='690000',round(s25*s27,2),0) s690000, if(s13='6A1000',round(s25*s27,2),0) s6A1000, if(s13='6A2000',round(s25*s27,2),0) s6A2000, if(s13='6A3000',round(s25*s27,2),0) s6A3000, " .
              "if(s13='6A5000',round(s25*s27,2),0) s6A5000, if(s13='6A6000',round(s25*s27,2),0) s6A6000, if(s13='6A4300',round(s25*s27,2),0) s6A4100, if(s13='6A4300',round(s25*s28,2),0) s6A4200 " .
              "from " .
              //"(select s09, if(substr(s09,1,1)='1', s13, if(substr(s09,1,2)='27', s13, '6A4300')) s13 , s25, s26, s27, s28 from casesent where (date_format(s01,'%Y-%m')='$month') ) a ) b ";
              "(select s09, if(substr(s09,1,1)='1', s13, if(substr(s09,1,2)='27', s13, '6A4300')) s13 , s25, s26, s27, s28 from casesent where s01 between '$bdate' and '$edate' ) a ) b ";   
                                                                                                                                                                               
    $result= mysql_query($query);            
    $row   = mysql_fetch_array($result);
    $objPHPExcel->setActiveSheetIndex(0)     
                ->setCellValue('D4', $row['s690000'])
                ->setCellValue('D5', $row['s6A1000'])   
                ->setCellValue('D6', $row['s6A2000'])   
                ->setCellValue('D7', $row['s6A3000']) 
                ->setCellValue('D8', $row['s6A4100'])   
                ->setCellValue('D9', $row['s6A4200'])
                ->setCellValue('D10',$row['s6A5000'])
                ->setCellValue('D11',$row['s6A6000']) ;     
                
    //計算delay數
    $query = "select sum(d690000) d690000, sum(d6A1000) d6A1000, sum(d6A2000) d6A2000, sum(d6A3000) d6A3000, sum(d6A4100) d6A4100, sum(d6A4200) d6A4200, sum(d6A5000) d6A5000, sum(d6A6000) d6A6000 " .
              "from " .
              "(select if(makercode='690000', mplus, 0) d690000, if(makercode='6A1000', mplus, 0) d6A1000, if(makercode='6A2000', mplus, 0) d6A2000, if(makercode='6A3000', mplus, 0) d6A3000, " .
              "if(makercode='6A4300', mplus, 0) d6A4100, if(makercode='6A4300', rplus, 0) d6A4200, if(makercode='6A5000', mplus, 0) d6A5000, if(makercode='6A6000', mplus, 0) d6A6000 " .
              "from " .
              "(select if(substr(dd.productcode,1,1)='1', dd.makercode,if(substr(dd.productcode,1,2)='27','6A5000','6A4300')) makercode,  round((dd.qty*dd.mplus),2) mplus,  round((dd.qty*dd.rplus),2) rplus " .
              "from delay d, delaydetail dd " .
              "where d.orderno=dd.orderno and d.tdate=dd.tdate and instr( rx,'SAMPLE')=false " .
              "and d.tdate between '$bdate' and '$edate' and d.tdate>=d.duedate and d.status='' ".
              //"and (date_format(d.tdate,'%Y-%m')='$month') and d.tdate>=d.duedate and d.status='' ".   
              "and d.orderno not in (select orderno from delaydetail where productcode in ('1Z151','1Z152','1Z153')))a)b ";

    $result= mysql_query($query);            
    $row   = mysql_fetch_array($result);
    $objPHPExcel->setActiveSheetIndex(0)     
                ->setCellValue('E4', $row['d690000'])
                ->setCellValue('E5', $row['d6A1000'])   
                ->setCellValue('E6', $row['d6A2000'])   
                ->setCellValue('E7', $row['d6A3000']) 
                ->setCellValue('E8', $row['d6A4100'])   
                ->setCellValue('E9', $row['d6A4200'])
                ->setCellValue('E10',$row['d6A5000'])
                ->setCellValue('E11',$row['d6A6000']) ;             
                       

    //計算當月合計 內返/重修 顆/床數         
    $query = "select sum(rj690000) rj690000, sum(rd690000) rd690000, sum(rj6A1000) rj6A1000, sum(rd6A1000) rd6A1000, sum(rj6A2000) rj6A2000, sum(rd6A2000) rd6A2000, " .
              "sum(rj6A3000) rj6A3000, sum(rd6A3000) rd6A3000, sum(rj6A4100) rj6A4100, sum(rd6A4100) rd6A4100, sum(rj6A4200) rj6A4200, sum(rd6A4200) rd6A4200, " .
              "sum(rj6A5000) rj6A5000, sum(rd6A5000) rd6A5000, sum(rj6A6000) rj6A6000, sum(rd6A6000) rd6A6000 " .
              "from " .
              "(select if(s13='690000', reject, 0) rj690000, if(s13='690000', redo, 0) rd690000, if(s13='6A1000', reject, 0) rj6A1000, if(s13='6A1000', redo, 0) rd6A1000, " .
              "if(s13='6A2000', reject, 0) rj6A2000, if(s13='6A2000', redo, 0) rd6A2000, if(s13='6A3000', reject, 0) rj6A3000, if(s13='6A3000', redo, 0) rd6A3000, " .
              "if(s13='6A4100', reject, 0) rj6A4100, if(s13='6A4100', redo, 0) rd6A4100, if(s13='6A4200', reject, 0) rj6A4200, if(s13='6A4200', redo, 0) rd6A4200, " .
              "if(s13='6A5000', reject, 0) rj6A5000, if(s13='6A5000', redo, 0) rd6A5000, if(s13='6A6000', reject, 0) rj6A6000, if(s13='6A6000', redo, 0) rd6A6000 " .
              "from " .
              "(select rid s13, sum(rqty1) reject, sum(rqty2) redo from casereject where rdate between '$bdate' and '$edate' group by rid ) a ) b " ;
              //"(select rid s13, sum(rqty1) reject, sum(rqty2) redo from casereject where (date_format(rdate,'%Y-%m')='$month') group by rid ) a ) b " ;  

    $result= mysql_query($query);            
    $row   = mysql_fetch_array($result);
    $objPHPExcel->setActiveSheetIndex(0)     
                ->setCellValue('F4', $row['rj690000'])
                ->setCellValue('F5', $row['rj6A1000'])   
                ->setCellValue('F6', $row['rj6A2000'])   
                ->setCellValue('F7', $row['rj6A3000']) 
                ->setCellValue('F8', $row['rj6A4100'])   
                ->setCellValue('F9', $row['rj6A4200'])
                ->setCellValue('F10',$row['rj6A5000'])
                ->setCellValue('F11',$row['rj6A6000']) 
                
                ->setCellValue('G4', $row['rd690000'])
                ->setCellValue('G5', $row['rd6A1000'])   
                ->setCellValue('G6', $row['rd6A2000'])   
                ->setCellValue('G7', $row['rd6A3000']) 
                ->setCellValue('G8', $row['rd6A4100'])   
                ->setCellValue('G9', $row['rd6A4200'])
                ->setCellValue('G10',$row['rd6A5000'])
                ->setCellValue('G11',$row['rd6A6000']) ;  

    $objPHPExcel->getActiveSheet()->setTitle($edate . '出貨顆床數統計');
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');   
    $filename='email/test/' . $edate . '_DeliveriedCasesReportv3.xls';
    $objWriter->save($filename);    
   
   /*
    ////email
    require_once("_classes.php");
    $_SERVER['SERVER_NAME'] = 'www.vedenlabs.com';
    $regards = "Veden Dental Labs Inc.";
       
    $m = new mailer();
    $m->setMessage($regards);
    $m->setPriority( 'High' );
    $m->setFrom( "frankyu@vedenlabs.com", "Veden Dental Labs Inc"  );
    $m->setReplyTo( "frankyu@vedenlabs.com", "Veden Dental Labs Inc" );  
    $m->attachFile( $filename, $filename, "application/vnd.ms-excel"); 
    $m->send($email, $edate . ' Deliveried Cases Report v3') ;   
    msg('寄送完畢!!');
    */ 
    
?>
