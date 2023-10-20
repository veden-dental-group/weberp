<?
//每天計算當天出貨的顆數

session_start();   
include("_data.php");  
date_default_timezone_set('Asia/Taipei'); 
$edate=date('Y-m-d');    
//$edate='2012-08-21';
//先刪除同一日期的所有資料
$query2= "delete from casesent where s01='$edate' ";     
$result2= mysql_query($query2);

$s2= "select oga02, sfb82, gem02, sfbud02, sfb08, sfb22, sfb221, sfb01, to_char(oea02,'yyyy-mm-dd') oea02,  to_char(ta_oea005,'yyyy-mm-dd') ta_oea005, sfb05, ima02, imaud07, occ01, occ02 " .
         "from oga_file, ogb_file, sfb_file, ima_file, oea_file, gem_file, occ_file " .
         //"where sfb81<=to_date('$edate','yy/mm/dd') and sfb81>=to_date('120201','yy/mm/dd')   " . //某天以前  
         "where oga02=to_date('$edate','yy/mm/dd') " .     
         "and oga01=ogb01 " .
         "and ogb31=sfb22 " .   
         "and ogb32=sfb221 " .                                           
         "and sfb22=oea01 and oea04=occ01 " .
         "and sfb05=ima01 " .
         "and sfb82=gem01 " .
         "order by sfb82,sfb22,sfb221";

$erp_sql2 = oci_parse($erp_conn1,$s2 );  
oci_execute($erp_sql2);  
                                                   
while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {     
    $queryus = "insert into casesent ( s01,s02,s03,s031,s04,s05,s06,s09,s10,s11,s12,s13,s14,s25,s26 ) values (  
               '" . $edate                      . "',
               '" . safetext($row2['SFBUD02'])  . "',   
               '" . $row2['OGB31']              . "',  
               '" . $row2['OGB32']          . "',  
               '" . $row2['SFB01']          . "',  
               '" . $row2['OEA02']          . "',  
               '" . $row2['TA_OEA005']      . "',  
               '" . $row2['SFB05']          . "',  
               '" . $row2['IMA02']          . "',  
               '" . $row2['OCC01']          . "',  
               '" . safetext($row2['OCC02']). "',  
               '" . $row2['SFB82']          . "',  
               '" . safetext($row2['GEM02']). "',  
               " . $row2['SFB08']           . ",  
               " . $row2['IMAUD07']         . " )";     
    $resultus = mysql_query($queryus);       
    $msg=mysql_error();
    
}  
                  
error_reporting(E_ALL);  
require_once 'classes/PHPExcel.php'; 
require_once 'classes/PHPExcel/IOFactory.php';  
//$objPHPExcel = new PHPExcel();
$objReader = PHPExcel_IOFactory::createReader('Excel5');  
$objPHPExcel = $objReader->load("templates/erp_sentcases_4m.xls");  
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
            ->setCellValue('A1', $edate. ' 出貨顆/床數統計'); 
$objPHPExcel->setActiveSheetIndex(0) ->mergeCells('A1:G1');  
//$objPHPExcel->setActiveSheetIndex(0)   
//            ->setCellValue('A3', '序號') 
//            ->setCellValue('B3', '製處')   
//            ->setCellValue('C3', 'CASE號碼')   
//            ->setCellValue('D3', '訂單號碼')   
//            ->setCellValue('E3', '訂單序號')    
//            ->setCellValue('F3', '工單號碼')  
//            ->setCellValue('I3', '客戶應交貨日期') 
//            ->setCellValue('J3', 'Delay 天數') 
//            ->setCellValue('K3', '預計出貨日期') 
//            ->setCellValue('L3', '產品') 
//            ->setCellValue('M3', 'Delay 原因')   
//            ->setCellValue('P3', '客戶')
//            ->setCellValue('Q3', '回覆時間')
//            ->setCellValue('R3', '板色'); 
              
$month=substr($edate,0,7);                     
          
//$query1= " (select s13 , sum(s25*s26) sent, 0 tsent, 0 reject, 0 redo, 0 delay  from casesent where s01='$edate' group by s13) ";  
//$query2= " (select s13 , 0 sent, sum(s25*s26) tsent, 0 reject, 0 redo, 0 delay  from casesent where (date_format(s01,'%Y-%m')='$month') group by s13) ";  
//$query3= " (select rid s13, 0 sent, 0 tsent, sum(rqty1) reject, sum(rqty2) redo, 0 delay from casereject where (date_format(rdate,'%Y-%m')='$month') group by rid) ";   
//$query4= " (select d13 s13 , 0 sent, 0 tsent, 0 reject, 0 redo, sum(d25*d26) delay  from casedelay where (date_format(d21,'%Y-%m')='$month') and d16='' group by d13) ";    
//$query = "select mid, mname, sum(sent) sent, sum(tsent) tsent, sum(reject) reject, sum(redo) redo, sum(delay) delay from " . 
//         "(select s13 , sent, tsent, reject, redo, delay from " .
//         "( $query1 union all $query2 union all $query3 union $query4 ) a) b, maker where s13=mid and (sent+tsent+reject+redo+delay)!=0 and mflag1='Y' group by mid order by mid  ";
         
$query1= " (select s13, sum(sent) sent, 0 tsent, 0 reject, 0 redo, 0 delay from (select if(s13='6A4300', concat(s13, s26), s13) s13 , (s25*s26) sent from casesent where s01='$edate') aa group by s13) ";  

$query2= " (select s13, 0 sent, sum(tsent) tsent, 0 reject, 0 redo, 0 delay  from (select if(s13='6A4300', concat(s13, s26), s13) s13, (s25*s26) tsent from casesent where (date_format(s01,'%Y-%m')='$month')) bb group by s13) ";  

$query3= " (select rid s13, 0 sent, 0 tsent, sum(rqty1) reject, sum(rqty2) redo, 0 delay from casereject where (date_format(rdate,'%Y-%m')='$month') group by rid) ";   

//$query4= " (select s13, 0 sent, 0 tsent, 0 reject, 0 redo, sum(delay) delay  from (select if(d13='6A4300', concat(d13, d26), d13) s13 , (d25*d26) delay  from casedelay where (date_format(d21,'%Y-%m')='$month') and d16='') cc group by s13) ";    

$query4= " (select s13, 0 sent, 0 tsent, 0 reject, 0 redo, sum(delay) delay  from 
           (select if(dd.makercode='6A4300', concat(dd.makercode, dd.plus), dd.makercode) s13 , (dd.qty*dd.plus) delay  from delay d, delaydetail dd where d.orderno=dd.orderno and d.orderdate=dd.orderdate and instr( rx,'SAMPLE')=false and (date_format(d.tdate,'%Y-%m')='$month') and d.tdate>=d.duedate and d.status='') cc group by s13) ";    

$query = "select mid, mname, sum(sent) sent, sum(tsent) tsent, sum(reject) reject, sum(redo) redo, sum(delay) delay from " . 
         "(select s13 , sent, tsent, reject, redo, delay from " .
         "( $query1 union all $query2 union all $query3 union $query4 ) a) b, maker where s13=mid and (sent+tsent+reject+redo+delay)!=0 and mflag1='Y' group by mid order by mid  ";         
         
$result= mysql_query($query);

$i=0;
$y=4;                                                     
while ($row = mysql_fetch_array($result)){     
    $i++;
    $objPHPExcel->setActiveSheetIndex(0)     
                ->setCellValue('A'. $y, $i)             
                ->setCellValue('B'. $y, $row['mname'])
                ->setCellValue('C'. $y, $row['sent'])   
                ->setCellValue('D'. $y, $row['tsent'])   
                ->setCellValue('E'. $y, $row['delay']) 
                ->setCellValue('F'. $y, $row['reject'])   
                ->setCellValue('G'. $y, $row['redo'])
                ->setCellValue('H'. $y, $row['tsent']-$row['delay']-$row['reject']-$row['redo']) ;
    if (($y%2)==0){
        $objPHPExcel->setActiveSheetIndex(0) ->getStyle('A'.$y.':H'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');    
    }               
                
    $y++;       
}         
      
$objPHPExcel->setActiveSheetIndex(0) ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
$objPHPExcel->setActiveSheetIndex(0) ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
$objPHPExcel->setActiveSheetIndex(0) ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
$objPHPExcel->setActiveSheetIndex(0) ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
$objPHPExcel->setActiveSheetIndex(0) ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
$objPHPExcel->setActiveSheetIndex(0) ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);    
$objPHPExcel->setActiveSheetIndex(0) ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
$objPHPExcel->setActiveSheetIndex(0) ->getStyle('H'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
$objPHPExcel->setActiveSheetIndex(0)                 
            ->setCellValue('A'. $y, '合計');
$objPHPExcel->setActiveSheetIndex(0) ->mergeCells('A'.$y.':B'.$y); 
$objPHPExcel->setActiveSheetIndex(0) ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);   
$objPHPExcel->setActiveSheetIndex(0)                 
            ->setCellValue('C'. $y, '=SUM(C4:C' . ($y-1) . ')')
            ->setCellValue('D'. $y, '=SUM(D4:D' . ($y-1) . ')')
            ->setCellValue('E'. $y, '=SUM(E4:E' . ($y-1) . ')')
            ->setCellValue('F'. $y, '=SUM(F4:F' . ($y-1) . ')')
            ->setCellValue('G'. $y, '=SUM(G4:G' . ($y-1) . ')')
            ->setCellValue('H'. $y, '=SUM(H4:H' . ($y-1) . ')');
$y++;
//$objPHPExcel->setActiveSheetIndex(0)                 
//          ->setCellValue('A'. $y, '-- 以下無資料 --');
//$objPHPExcel->setActiveSheetIndex(0) ->mergeCells('A'.$y.':G'.$y);   
//$objPHPExcel->setActiveSheetIndex(0) ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);         
// Rename sheet
$objPHPExcel->getActiveSheet()->setTitle($edate . '出貨顆床數統計');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');   
$filename='report/' . $edate . '_DeliveriedCasesReport.xls';
$objWriter->save($filename);    

if ($i>0) {          
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
    $m->send("gdeliveriedcase@vedenlabs.com", $edate . ' Deliveried Cases Report') ;  
    //$m->send("cs@vedenlabs.com", $edate . ' Deliveried Cases Report') ;  
    //$m->send("m@vedenlabs.com", $edate . ' Deliveried Cases Report') ;   
    //$m->send("fa6@vedenlabs.com", $edate . ' Deliveried Cases Report') ;    
    //$m->send("fa3@vedenlabs.com", $edate . ' Deliveried Cases Report') ; 
    
    //gdeliveriedcase 
}
