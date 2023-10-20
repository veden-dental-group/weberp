<?
//每天重算各製處到貨顆數/Delay顆數/內返顆數
//星期五, 六 PTL不計算delay

session_start();   
include("_data.php");  
date_default_timezone_set('Asia/Taipei'); 
$tdate=date('Y-m-d');                       
$edate=date('Y-m-d', strtotime("-4 days"));   

//find the lastest fax date
function findfaxdatebysfb01($erp_conn1,$sfb01){
  $s3= "select to_char(tc_ohf008,'mm-dd-yyyy') tc_ohf008, tc_ohf009 from tc_ohf_file where tc_ohf001='$sfb01' order by tc_ohf008, tc_ohf009 desc";   
  $erp_sql3 = oci_parse($erp_conn1,$s3 );  
  oci_execute($erp_sql3);  
  $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);    
  if (is_null($row3['TC_OHF008'])){
      $faxdate='';
  } else {
      $faxdate= $row3['TC_OHF008']. ' ' . $row3['TC_OHF009'];
  }
  return $faxdate;  
}

//find the case stage
function findstagebysfb01($erp_conn1,$sfb01){
  $s3= "select tc_srg001, tc_srg006,tc_srg010, tc_srg018, tc_srg022, ecb17 from tc_srg_file, ecb_file where tc_srg001='$sfb01' and tc_srg005 is not null and tc_srg003=ecb01 and tc_srg004=ecb03 ";   
  $erp_sql3 = oci_parse($erp_conn1,$s3 );  
  oci_execute($erp_sql3);  
  $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);    
  $tc_srg001=$row3['TC_SRG001'];
  $tc_srg006=$row3['TC_SRG006']; 
  $tc_srg010=$row3['TC_SRG010'];
  $tc_srg018=$row3['TC_SRG018'];
  $tc_srg022=$row3['TC_SRG022'];
  $ecb17=$row3['ECB17'];  
  if (is_null($tc_srg001)){
      $nowstage='工序不明';
  } else {
      if ($tc_srg018=='Y') {  //返工    
          if ($tc_srg006=='Y') { //要QC                
              if (is_null($tc_srg022)) {
                  $nowstage=$ecb17;
              } else {
                  $nowstage=$ecb17 . " QC";  
              }      
          } else {
            $nowstage=$ecb17; 
          }        
      } else {
          if ($tc_srg006=='Y') { //要QC                
              if (is_null($tc_srg010)) {  //未出站
                  $nowstage=$ecb17;
              } else {
                  $nowstage=$ecb17. " QC";  
              }      
          } else {
            $nowstage=$ecb17; 
          }
      }
      
  }
  return $nowstage;  
}

//先刪除同一日期的所有資料
$query2= "delete from delay where tdate='$tdate' ";     
$result2= mysql_query($query2);

$query2= "delete from delaydetail where tdate='$tdate' ";     
$result2= mysql_query($query2);

$query2= "delete from delayfax where tdate='$tdate' ";     
$result2= mysql_query($query2);

$weekday=date('w',strtotime($tdate)); //0:日  1:一... 5:五  6:六)    
if ($weekday=5 or $weekday=6) {  //PTL, LP, F-1 星期五六 不算delay
    $s2= "select oea01, to_char(oea02,'yyyy-mm-dd') oea02, oea04, occ02, ta_oea006, sfb01,sfb22, sfb221, sfb82, gem02, sfb05, ima02, sfb08, imaud07, substr(imaud02,2,1) mday, substr(imaud02,3,1) qday, substr(imaud02,4,3) mplus, substr(imaud02,7,3) rplus  " .
         "from oea_file, sfb_file, occ_file, ima_file, gem_file " .
         "where oea01=sfb22 and oea04=occ01 and sfb05=ima01 and sfb82=gem01 and sfb28 is  null " .
         "and oea04 !='E129001' and oea04 !='E129002' and oea04 != 'E143001' and oea04 != 'E145001' " .   
         "and (ta_ima003!='Y' and ta_ima004!='Y' and ta_ima005!='Y' ) " .
         "and not exists  (select 1 from tc_ogb_file where tc_ogb003=oea01 ) " .
         "and not exists  (select 1 from tc_ohf_file where tc_ohf001=sfb01 and tc_ohf008 is null) " .
         "and oea02 <= to_date('$edate','yy/mm/dd') " .
         "order by sfb22,sfb221 " ;
} else {
    $s2= "select oea01, to_char(oea02,'yyyy-mm-dd') oea02, oea04, occ02, ta_oea006, sfb01,sfb22, sfb221, sfb82, gem02, sfb05, ima02, sfb08, imaud07, substr(imaud02,2,1) mday, substr(imaud02,3,1) qday, substr(imaud02,4,3) mplus, substr(imaud02,7,3) rplus   " .
         "from oea_file, sfb_file, occ_file, ima_file, gem_file " .
         "where oea01=sfb22 and oea04=occ01 and sfb05=ima01 and sfb82=gem01 and sfb28 is  null " .        
         "and (ta_ima003!='Y' and ta_ima004!='Y' and ta_ima005!='Y' ) " .
         "and not exists  (select 1 from tc_ogb_file where tc_ogb003=oea01 ) " .
         "and not exists  (select 1 from tc_ohf_file where tc_ohf001=sfb01 and tc_ohf008 is null) " .
         "and oea02 <= to_date('$edate','yy/mm/dd') " .
         "order by sfb22, sfb221 " ;
}
 
$erp_sql2 = oci_parse($erp_conn1,$s2 );  
oci_execute($erp_sql2);  
$ooea01='';                                                   
while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {       
    //寫單頭
    if ($ooea01!=$row2['OEA01']) { //新的工單號 在單頭加一筆
        $queryd= "insert into delay ( tdate, orderdate, orderno, rx, clientid, clientname, duedate, stage, makeday, recday, faxday, status) values (
                 '" . $tdate                        . "',
                 '" . $row2['OEA02']                . "', 
                 '" . $row2['OEA01']                . "',   
                 '" . $row2['TA_OEA006']            . "',  
                 '" . $row2['OEA04']                . "',  
                 '" . $row2['OCC02']                . "', 
                 '" . $tdate                        . "',
                 '', 0, 0, 0, '')";
        $resultd = mysql_query($queryd); 
        $ooea01=$row2['OEA01'];
    }
      
    //寫單身  
    $querydd = "insert into delaydetail ( tdate, orderdate, orderno, seq, ticketno, productcode, productname, makercode, makername, qty, mplus, rplus, makeday, qtyday, faxday, stage ) values (  
               '" . $tdate                        . "',                
               '" . $row2['OEA02']                . "', 
               '" . $row2['OEA01']                . "',
               '" . $row2['SFB221']               . "',   
               '" . $row2['SFB01']                . "',  
               '" . $row2['SFB05']                . "',  
               '" . $row2['IMA02']                . "',  
               '" . $row2['SFB82']                . "',  
               '" . $row2['GEM02']                . "',  
               '" . $row2['SFB08']                . "',  
               '" . $row2['MPLUS']                . "', 
               '" . $row2['RPLUS']                . "',    
               '" . $row2['MDAY']                 . "',  
               '" . $row2['QDAY']                 . "',0,
               '" . findstagebysfb01($erp_conn1,$row2['SFB01'])  . "')";     
    $resultdd = mysql_query($querydd);     
    
    //寫傳真
    $s3="select to_char(tc_ohf004,'yyyy-mm-dd') tc_ohf004, to_char(tc_ohf008,'yyyy-mm-dd') tc_ohf008 from tc_ohf_file where tc_ohf001='" . $row2['SFB01'] . "' ";
    $erp_sql3 = oci_parse($erp_conn1,$s3 );  
    oci_execute($erp_sql3);   
    while ($row3 = oci_fetch_array($erp_sql3, OCI_ASSOC)) { 
        $faxday = ((strtotime($row3['TC_OHF008']) - strtotime($row3['TC_OHF004']))/86400)+1;
        $querydf = "insert into delayfax ( tdate, ticketno, indate, outdate, faxday ) values (  
                   '" . $tdate                        . "',
                   '" . $row2['SFB01']                . "',
                   '" . $row3['TC_OHF004']            . "',   
                   '" . $row3['TC_OHF008']            . "',$faxday )";     
        $resultdF = mysql_query($querydf);        
    }           
}  

//程式到此 已取出三天前到貨但未出貨的case 開始計算應工作的天數


//計算fax的天數 寫到 delaydetail中
$query="select ticketno from delaydetail where tdate='$tdate'";
$result=mysql_query($query);
while ($row=mysql_fetch_array($result)){
  $ticketno=$row['ticketno'];
  $querydf="select sum(faxday) faxday from delayfax where tdate='$tdate' and  ticketno='$ticketno'";
  $resultdf=mysql_query($querydf);
  $rowdf=mysql_fetch_array($resultdf);
  $faxday=$rowdf['faxday'];
  
  $querydu="update delaydetail set faxday=$faxday where tdate='$tdate' and ticketno='$ticketno'";
  $resultdu=mysql_query($querydu);    
}

//計算每個CASE的各品代的天數相加 - (品代個數-1) + 最大的傳真天數


$queryd="select orderno, orderdate from delay where tdate='$tdate' ";
$resultd=mysql_query($queryd);
while ($rowd=mysql_fetch_array($resultd)){
    $orderno=$rowd['orderno'];
    
    //取本工單各品代的合計天數, 要判斷顆數是否>加一天的顆數 
    //  oracle可用 sign()函数根据某个值是0、正数还是负数，分别返回0、1、-1 , 因此判斷 qty-qtyday 是否是正數, 如果是, 則加一天, 否則不用加天數
    $querydd1="select sum(makeday+(if(qty>qtyday,if(qtyday=0,0,1),0))) makeday from delaydetail where tdate='$tdate' and orderno='$orderno' ";
    $resultdd1=mysql_query($querydd1);
    $rowdd1=mysql_fetch_array($resultdd1);
    $makeday=min($rowdd1['makeday'],10);
    
    //取本工單各幾個品代
    $querydd2="select count(*) recday from delaydetail where tdate='$tdate' and orderno='$orderno' ";
    $resultdd2=mysql_query($querydd2);
    $rowdd2=mysql_fetch_array($resultdd2);
    $recday=$rowdd2['recday']-1;
    
    //取本工單各品代的傳真的最大天數
    $querydd3="select max(faxday) faxday from delaydetail where tdate='$tdate' and orderno='$orderno' ";
    $resultdd3=mysql_query($querydd3);
    $rowdd3=mysql_fetch_array($resultdd3);
    $faxday=$rowdd3['faxday'];
  
    $orderdate=$rowd['orderdate'];
    $totalseconds= (min($makeday - $recday,10) + $faxday) * 86400 ;
    $duedate=date('Y-m-d',strtotime($orderdate)+$totalseconds);   
    
    $querydu="update delay set makeday=$makeday, recday=$recday, faxday=$faxday, duedate='$duedate' where tdate='$tdate' and orderno='$orderno'";
    $resultdu=mysql_query($querydu);   
}

//寄email了


error_reporting(E_ALL);  
require_once 'classes/PHPExcel.php'; 
require_once 'classes/PHPExcel/IOFactory.php';  
//$objPHPExcel = new PHPExcel();
$objReader = PHPExcel_IOFactory::createReader('Excel5');  
$objPHPExcel = $objReader->load("templates/erp_delaycases_4mv3.xls");  
// Set properties
$objPHPExcel ->getProperties()->setCreator("Frank")
           ->setLastModifiedBy("Frank")
           ->setTitle("$tdate 應出貨但未出貨的RX#")
           ->setSubject("$tdate 應出貨但未出貨的RX#")
           ->setDescription("$tdate 應出貨但未出貨的RX#")
           ->setKeywords("$tdate 應出貨但未出貨的RX#")
           ->setCategory("$tdate 應出貨但未出貨的RX#");

// Add some data
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A1', $tdate . ' delay cases report (應出貨但未出貨的RX#)'); 
$objPHPExcel->setActiveSheetIndex(0) ->mergeCells('A1:N1');       
          
$query2= "select dd.makercode, dd.makername makername, d.clientid clientid, d.clientname clientname, d.rx rxno, d.orderno orderno, dd.seq seq,  dd.ticketno ticketno, d.orderdate orderdate,  " .
         "d.duedate duedate, (d.makeday-d.recday) makeday, d.faxday faxday, dd.productcode productcode, dd.productname productname, dd.qty qty, dd.mplus mplus, dd.rplus rplus, concat(qtyday,'/',if(qty>qtyday,if(qtyday=0,0,1),0)) qtyday, dd.stage stage " .
         "from delay d, delaydetail dd " .
         "where d.tdate='$tdate' and d.tdate >= d.duedate and d.orderno=dd.orderno and dd.tdate='$tdate' " .
         "and d.orderno not in (select orderno from delaydetail where productcode in ('1Z151','1Z152','1Z153')) " .  //toronto不計算delay
         "order by dd.makercode, d.clientid, d.rx";     
$result2= mysql_query($query2);

$y=2;   
$i=0;   
$j=0;
$oldrxno=''; 
$oldmakercode='';                                             
while ($row2 = mysql_fetch_array($result2)){  
    if ($oldmakercode!=$row2['makercode']) {
        $oldmakercode =$row2['makercode'];
        $i=0;  //不同製處就由1開始算
        $y+=2; //列印位置
    } 
    
    if ($oldrxno!=$row2['rxno']){
        $rxno=$row2['rxno'];
        $orderno=$row2['orderno'];
        $seq=$row2['seq'];
        $oldrxno=$row2['rxno'];
        $i++;
        $j++;
        $di=$i;
        $makeday=$row2['makeday'];
        $faxday=$row2['faxday'];
        
    } else {
        $rxno='';
        $orderno='';      
        $seq='';
        $di='';
        $makeday='';
        $faxday='';
    }
    
    $objPHPExcel->setActiveSheetIndex(0)     
                ->setCellValue('A'. $y, $di)             
                ->setCellValue('B'. $y, $row2['makername'])   
                ->setCellValue('C'. $y, $row2['clientid'].'--'.$row2['clientname'])  
                ->setCellValue('D'. $y, $rxno)   
                ->setCellValue('E'. $y, $orderno.' ' . $seq)   
                ->setCellValue('F'. $y, $row2['ticketno'])     
                ->setCellValue('G'. $y, $row2['orderdate']) 
                ->setCellValue('H'. $y, $row2['duedate'])    
                ->setCellValue('I'. $y, $makeday)      
                ->setCellValue('J'. $y, $faxday) 
                ->setCellValue('K'. $y, $row2['qtyday'])    
                ->setCellValue('L'. $y, $row2['productcode'].'--'.$row2['productname'])     
                ->setCellValue('M'. $y, $row2['stage'])                       
                ->setCellValue('N'. $y, $row2['qty'])     
                ->setCellValue('O'. $y, $row2['mplus'])  
                ->setCellValue('P'. $y, $row2['rplus']); 
    if (($y%2)==0){
        $objPHPExcel->setActiveSheetIndex(0) ->getStyle('A'.$y.':P'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');    
    }               
                
    $y++;       
}         

$y++;     

$objPHPExcel->setActiveSheetIndex(0)                 
          ->setCellValue('A'. $y, '合計: '. $j. ' 組CASES');
$objPHPExcel->setActiveSheetIndex(0) ->mergeCells('A'.$y.':P'.$y); 
$objPHPExcel->setActiveSheetIndex(0) ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);   
$y++;
$objPHPExcel->setActiveSheetIndex(0)                 
          ->setCellValue('A'. $y, '-- 以下無資料 --');
$objPHPExcel->setActiveSheetIndex(0) ->mergeCells('A'.$y.':P'.$y);   
$objPHPExcel->setActiveSheetIndex(0) ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);         
// Rename sheet
$objPHPExcel->getActiveSheet()->setTitle($edate. ' 應出貨但未出貨的RX#');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');   
$filename='email/' . $tdate . '_DelayCasesReportv3.xls';
$objWriter->save($filename);    

if ($j>0) {          
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
    //$m->send("gcheckinout@vedenlabs.com", $tdate . ' delay cases report v2') ;  
    $m->send("frank@vedenlabs.com", $tdate . ' delay cases report v2') ;  
    //$m->send("cs@vedenlabs.com", $tdate . ' delay cases report') ;  
    //$m->send("m@vedenlabs.com", $tdate . ' delay cases report') ;   
    //$m->send("fa3@vedenlabs.com", $tdate . ' delay cases report') ;    
    //$m->send("fa6@vedenlabs.com", $tdate . ' delay cases report') ;    
}

