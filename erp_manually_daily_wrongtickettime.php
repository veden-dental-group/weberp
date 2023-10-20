<?
//檢查當天報工的時間在10分鐘內的
session_start();   
include("_data.php");  
date_default_timezone_set('Asia/Taipei'); 
$edate=date('Y-m-d', strtotime("-1 days"));
error_reporting(E_ALL);  
require_once 'classes/PHPExcel.php'; 
require_once 'classes/PHPExcel/IOFactory.php';  
                                       
//前N天的合計數 看有無進步
$tdate=$edate;
$i=140;
while ( $i>=1){   
  $sdate[$i]=$tdate; 
  $i--; 
  $tdate=date('Y-m-d',strtotime("-1 day", strtotime($tdate)));
}  
//因為已經多減一天了  要加回來
$bdate = date('Y-m-d',strtotime("+1 day", strtotime($tdate))); 
        
$objPHPExcel = new PHPExcel();
// Set properties
$objPHPExcel ->getProperties()->setCreator("Frank")
           ->setLastModifiedBy("Frank")
           ->setTitle("報工有問題之工單 依工序合計")
           ->setSubject("報工有問題之工單 依工序合計")
           ->setDescription("報工有問題之工單 依工序合計")
           ->setKeywords("報工有問題之工單 依工序合計")
           ->setCategory("報工有問題之工單 依工序合計");

// Add some data
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A1', $bdate . ' -- ' .$edate . ' 同時刷進和刷出的合計數量'); 
$objPHPExcel->setActiveSheetIndex(0) ->mergeCells('A1:E1');  
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A3', '製處')              
            ->setCellValue('B3', '工序')    
            ->setCellValue('C3', $sdate[14] . "(" . date('D',strtotime($sdate[14])) . ")")  
            ->setCellValue('D3', $sdate[13] . "(" . date('D',strtotime($sdate[13])) . ")")  
            ->setCellValue('E3', $sdate[12] . "(" . date('D',strtotime($sdate[12])) . ")")  
            ->setCellValue('F3', $sdate[11] . "(" . date('D',strtotime($sdate[11])) . ")")  
            ->setCellValue('G3', $sdate[10] . "(" . date('D',strtotime($sdate[10])) . ")")  
            ->setCellValue('H3', $sdate[9] . "(" . date('D',strtotime($sdate[9])) . ")")  
            ->setCellValue('I3', $sdate[8] . "(" . date('D',strtotime($sdate[8])) . ")")  
            ->setCellValue('J3', $sdate[7] . "(" . date('D',strtotime($sdate[7])) . ")")     
            ->setCellValue('K3', $sdate[6] . "(" . date('D',strtotime($sdate[6])) . ")")    
            ->setCellValue('L3', $sdate[5] . "(" . date('D',strtotime($sdate[5])) . ")")   
            ->setCellValue('M3', $sdate[4] . "(" . date('D',strtotime($sdate[4])) . ")")   
            ->setCellValue('N3', $sdate[3] . "(" . date('D',strtotime($sdate[3])) . ")")   
            ->setCellValue('O3', $sdate[2] . "(" . date('D',strtotime($sdate[2])) . ")")   
            ->setCellValue('P3', $sdate[1] . "(" . date('D',strtotime($sdate[1])) . ")"); 
$i=14;     
$s2= "select gem02, ecb17, ";
for ($x=1; $x<=$i; $x++){
  $s2 .= " sum(t" . $x . ") t" . $x . ", " ;   
}      
$s2.= "'1' from ( select gem02, ecb17, " ;
for ($x=1; $x<=$i; $x++){
  $s2 .= " decode(gdate, '". $sdate[$x] . "', total ,0) t" . $x . ", " ;   
}          
$s2.="'1' from ( select gem02, tc_srg007 gdate, ecb17, count(*) total from " .
     "(select gem02,ecb17, to_char(tc_srg007,'yyyy-mm-dd') tc_srg007 " .
     "from tc_srg_file, sfb_file,gem_file,ecb_file " .
     "where tc_srg007 between to_date('$bdate','yy/mm/dd') and to_date('$edate','yy/mm/dd') and tc_srg007=tc_srg010 and substr(tc_srg008,1,4)=substr(tc_srg011,1,4) " .
     "and tc_srg001=sfb01 and sfb82=gem01 and sfb05=ecb01 and tc_srg004=ecb03 and ecb06!='9999' ) " .
     "group by gem02, tc_srg007, ecb17 )) group by gem02, ecb17 order by gem02,ecb17 ";
 
$erp_sql2 = oci_parse($erp_conn1,$s2 );  
oci_execute($erp_sql2);  
$y=4;  
$rec=0;                                                  
while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {        
    $objPHPExcel->setActiveSheetIndex(0)
                ->setCellValue('A'. $y, $row2['GEM02']) 
                ->setCellValue('B'. $y, $row2['ECB17'])
                ->setCellValue('C'. $y, $row2['T14'])   
                ->setCellValue('D'. $y, $row2['T13'])   
                ->setCellValue('E'. $y, $row2['T12'])       
                ->setCellValue('F'. $y, $row2['T11'])   
                ->setCellValue('G'. $y, $row2['T10'])   
                ->setCellValue('H'. $y, $row2['T9'])   
                ->setCellValue('I'. $y, $row2['T8'])   
                ->setCellValue('J'. $y, $row2['T7'])   
                ->setCellValue('K'. $y, $row2['T6']) 
                ->setCellValue('L'. $y, $row2['T5'])  
                ->setCellValue('M'. $y, $row2['T4'])  
                ->setCellValue('N'. $y, $row2['T3'])    
                ->setCellValue('O'. $y, $row2['T2']) 
                ->setCellValue('P'. $y, $row2['T1']);          
    $y++;
    $rec++;
}              

$objPHPExcel->getActiveSheet()->setTitle('最近十四日內報工有問題之工單合計');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');   
$filename3='email/' . $edate . '_wrongtickettimetotalweekly.xls';
$objWriter->save($filename3);   


if ($rec>0){     
    ////email 
    require_once("_classes.php");
    $_SERVER['SERVER_NAME'] = 'www.vedenlabs.com';
    $regards = "Veden Dental Labs Inc.";
       
    $m = new mailer();
    $m->setMessage($regards);
    $m->setPriority( 'High' );
    $m->setFrom( "frankyu@vedenlabs.com", "Veden Dental Labs Inc"  );
    $m->setReplyTo( "frankyu@vedenlabs.com", "Veden Dental Labs Inc" );   
    $m->attachFile( $filename3, $filename3, "application/vnd.ms-excel");   
    $m->send("frank@vedenlabs.com", $edate . ' Wrong Ticket Time') ;   
    //$m->send("m@vedenlabs.com", $edate . ' Wrong Ticket Time') ;
    //$m->send("csc18@vedenlabs.com", $edate . ' Ticket NO Weight') ;   
    //$m->send("sophia@vedenlabs.com", $edate . ' Wrong Ticket Time') ;
    //$m->send("annie@vedenlabs.com", $edate . ' Wrong Ticket Time') ;  
}
