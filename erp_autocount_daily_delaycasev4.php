<?php
//計算未出貨的工單 在每道工序的停留時間是否超過

session_start();

include("_data.php");
date_default_timezone_set('Asia/Taipei');

//傳入 迄日 迄時 起日 起時  傳回兩者相差的分鐘數
function timeDiff( $aDate, $aTime , $bDate, $bTime ) {       // 2014-10-10, 15:03:03, 2014-10-10, 10:13:03
   // 分割第一个时间
   $ayear = substr ( $aDate , 0 , 4 );
   $amonth = substr ( $aDate , 5 , 2 );
   $aday = substr ( $aDate , 8 , 2 );
   $ahour = substr ( $aTime , 0 , 2 );
   $aminute = substr ( $aTime , 3 , 2 );
   $asecond = substr ( $aTime , 6 , 2 );
  // 分割第二个时间
   $byear = substr ( $bDate , 0 , 4 );
   $bmonth = substr ( $bDate , 5 , 2 );
   $bday = substr ( $bDate , 8 , 2 );
   $bhour = substr ( $bTime , 0 , 2 );
   $bminute = substr ( $bTime , 3 , 2 );
   $bsecond = substr ( $bTime , 6 , 2 );
  // 生成时间戳
   $a = mktime ( $ahour , $aminute , $asecond , $amonth , $aday , $ayear );
   $b = mktime ( $bhour , $bminute , $bsecond , $bmonth , $bday , $byear );
   $timeDiff = round(($a - $b) / 60)  ;
   return $timeDiff ;
}

$ydate=date('Y-m-d', strtotime("-1 days"));

$tdate = date('Y-m-d');
$ttime = Date('H:i:s');

//先刪除同一日期的所有資料
$query2= "delete from delayv4 where tdate='$tdate' ";
$result2= mysql_query($query2);

//第一次報工
$s2 ="select to_char(oea02,'yyyy-mm-dd') oea02, oea01, oeb03, occ01, occ02,  tc_srg001, tc_srg002, tc_srg003, ima02, tc_srg004, ecb06, ecb17, tc_srg005, to_char(tc_srg007, 'yyyy-mm-dd') tc_srg007, tc_srg008, tc_srg009, gen02, gen03, gem02
      from tc_srg_file, ima_file, ecb_file, gen_file, sfb_file, oeb_file, oea_file, occ_file, gem_file
      where tc_srg005 is not null and  ( tc_srg007 is not null and tc_srg010 is null )
      and not exists  (select 1 from tc_ogb_file where tc_ogb002=tc_srg001 )
      and not exists  (select 1 from tc_ohf_file where tc_ohf001=tc_srg001 and tc_ohf008 is null)
      and not exists  (select 1 from sfb_file where sfb01=tc_srg001 and sfb28 is not null )
      and tc_srg001=sfb01 and sfb22=oeb01 and sfb221=oeb03 and oeb01=oea01
      and tc_srg003=ima01
      and tc_srg003=ecb01 and tc_srg004=ecb03
      and tc_srg009=gen01
      and oea04=occ01
      and gen03=gem01
      and ecb06!='9999'
      order by tc_srg001 " ;


$erp_sql2 = oci_parse($erp_conn1,$s2 );
oci_execute($erp_sql2);
while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
    //寫單身
    $querydd = "insert into delayv4 ( tdate, ttime, orderdate, orderno, seq, ticketno, rx, clientid, clientname, productcode, productname, qty, stepseq, stepid, stepname, checkinid, checkinname, checkindate, checkintime, checkinmakerid, checkinmakername, ttype ) values (
               '" . $tdate                        . "',
               '" . $ttime                        . "',
               '" . $row2['OEA02']                . "',
               '" . $row2['OEA01']                . "',
               '" . $row2['OEB03']                . "',
               '" . $row2['TC_SRG001']            . "',
               '" . $row2['TC_SRG002']            . "',
               '" . $row2['OCC01']                . "',
               '" . $row2['OCC02']                . "',
               '" . $row2['TC_SRG003']            . "',
               '" . $row2['IMA02']                . "',
               '" . $row2['TC_SRG005']            . "',
               '" . $row2['TC_SRG004']            . "',
               '" . $row2['ECB06']                . "',
               '" . $row2['ECB17']                . "',
               '" . $row2['TC_SRG009']            . "',
               '" . $row2['GEN02']                . "',
               '" . $row2['TC_SRG007']            . "',
               '" . $row2['TC_SRG008']            . "',
               '" . $row2['GEN03']                . "',
               '" . $row2['GEM02']                . "','1' ) ";

    $resultdd = mysql_query($querydd) or die(mysql_error());

}


//返工報工
$s2 ="select to_char(oea02,'yyyy-mm-dd') oea02, oea01, oeb03, tc_srg001, tc_srg002, tc_srg003, ima02, tc_srg004, ecb06, ecb17, tc_srg005, to_char(tc_srg019, 'yyyy-mm-dd') tc_srg019, tc_srg020, tc_srg021, gen02, gen03, gem02
      from tc_srg_file, ima_file, ecb_file, gen_file, sfb_file, oeb_file, oea_file, occ_file, gem_file
      where tc_srg005 is not null and  ( tc_srg010 is not null and  tc_srg019 is not null and tc_srg022 is null )
      and not exists  (select 1 from tc_ogb_file where tc_ogb002=tc_srg001 )
      and not exists  (select 1 from tc_ohf_file where tc_ohf001=tc_srg001 and tc_ohf008 is null)
      and not exists  (select 1 from sfb_file where sfb01=tc_srg001 and sfb28 is not null )
      and tc_srg001=sfb01 and sfb22=oeb01 and sfb221=oeb03 and oeb01=oea01
      and tc_srg003=ima01
      and tc_srg003=ecb01 and tc_srg004=ecb03
      and tc_srg021=gen01
      and oea04=occ01
      and gen03=gem01
      and ecb06!='9999'
      order by tc_srg001 " ;

$erp_sql2 = oci_parse($erp_conn1,$s2 );
oci_execute($erp_sql2);
while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
    //寫單身
    $querydd = "insert into delayv4 ( tdate, ttime, orderdate, orderno, seq, ticketno, rx, occ01, occ02, productcode, productname, qty, stepseq, stepid, stepname, checkinid, checkinname, checkindate, checkintime, checkinmakerid, checkinmakername, ttype ) values (
               '" . $tdate                        . "',
               '" . $ttime                        . "',
               '" . $row2['OEA02']                . "',
               '" . $row2['OEA01']                . "',
               '" . $row2['OEB03']                . "',
               '" . $row2['TC_SRG001']            . "',
               '" . $row2['TC_SRG002']            . "',
               '" . $row2['OCC01']                . "',
               '" . $row2['OCC02']                . "',
               '" . $row2['TC_SRG003']            . "',
               '" . $row2['IMA02']                . "',
               '" . $row2['TC_SRG005']            . "',
               '" . $row2['TC_SRG004']            . "',
               '" . $row2['ECB06']                . "',
               '" . $row2['ECB17']                . "',
               '" . $row2['TC_SRG021']            . "',
               '" . $row2['GEN02']                . "',
               '" . $row2['TC_SRG019']            . "',
               '" . $row2['TC_SRG020']            . "',
               '" . $row2['GEN03']                . "',
               '" . $row2['GEM02']                . "', '1') ";

    $resultdd = mysql_query($querydd) or die('142 delayv4 insert error!! '. mysql_error());

}

commit;

//到此 有進站未出站 且有在製的全取出來了 開始計算應做時間 及 負責人

$query="select * from delayv4 where tdate='$tdate' and ttype='1' order by pkey";
$result=mysql_query($query) or die ('125 Delay read error!! '.mysql_error());
while ($row=mysql_fetch_array($result)) {
  $pkey=$row['pkey'];
  $stepid=$row['stepid'];
  $qty=$row['qty'];
  $ttdate=$row['tdate'];
  $tttime=$row['ttime'];
  $checkindate=$row['checkindate'];
  $checkintime=$row['checkintime'];
  $checkinmakerid=$row['checkinmakerid'];
  $checkinmakerid=$row['checkinmakerid'];
  $checkinmakername=$row['checkinmakername'];
  $checkinid=$row['checkinid'];
  $checkinname=$row['checkinname'];

  $querys="select time1, qty2, time2 from step where id='$stepid'";
  $results = mysql_query($querys) or die ('138 Step read error!! ' .  mysql_error());
  $rows=mysql_fetch_array($results);
  if ($qty >=$rows['qty2'])  {
      $steptime=$rows['time2'];
  } else {
      $steptime=$rows['time1'];
  }

  $querym="select makerid, makername, managerid, managername from stepmanager where makerid='$checkinmakerid' and stepid='$stepid' ";
  $resultm = mysql_query($querym) or die ('147 StepManager read error!! ' .  mysql_error());
  $rowm=mysql_fetch_array($resultm);
  if(is_null($rowm['managerid'])) {  //沒有指定負責人 則為本人
    $managermakerid=$checkinmakerid;
    $managermakername=$checkinmakername;
    $managerid=$checkinid;
    $managername=$checkinname;
  } else {
    $managermakerid=$rowm['makerid'];
    $managermakername=$rowm['makername'];
    $managerid=$rowm['managerid'];
    $managername=$rowm['managername'];
  }
  $makingtime=timeDiff($ttdate, $tttime, $checkindate, $checkintime);

  $overtime=$makingtime - $steptime;


  $queryu = "update delayv4 set
                  managermakerid    = '" . $managermakerid    . "',
                  managermakername  = '" . $managermakername  . "',
                  managerid         = '" . $managerid         . "',
                  managername       = '" . $managername       . "',
                  makingtime        = '" . $makingtime        . "',
                  steptime          = '" . $steptime          . "',
                  overtime          = '" . $overtime          . "'
                  where pkey        = '" . $pkey              . "' limit 1";

  $resultu = mysql_query($queryu) or die ('195 Delayv4 Update error!! ' . $queryu . '  ' . mysql_error());
}



//昨天第一次出站報工
$s2 ="select to_char(oea02,'yyyy-mm-dd') oea02, oea01, oeb03, occ01, occ02, tc_srg001, tc_srg002, tc_srg003, ima02, tc_srg004, ecb06, ecb17, sfb08, to_char(tc_srg007, 'yyyy-mm-dd') tc_srg007, tc_srg008, tc_srg009, a.gen02 gen021, a.gen03 gen031, c.gem02 gem021, to_char(tc_srg010, 'yyyy-mm-dd') tc_srg010, tc_srg011, tc_srg012, b.gen02 gen022, b.gen03 gen032, d.gem02 gem022
      from tc_srg_file, ima_file, ecb_file, gen_file a , gen_file b, sfb_file, oeb_file, oea_file, occ_file, gem_file c, gem_file d
      where tc_srg010 = to_date('$ydate','yy/mm/dd')
      and tc_srg001=sfb01 and sfb22=oeb01 and sfb221=oeb03 and oeb01=oea01
      and tc_srg003=ima01
      and tc_srg003=ecb01 and tc_srg004=ecb03
      and tc_srg009=a.gen01
      and a.gen03=c.gem01
      and tc_srg012=b.gen01
      and b.gen03=d.gem01
      and oea04=occ01
      and ecb06!='9999'
      order by tc_srg001 " ;


$erp_sql2 = oci_parse($erp_conn1,$s2 );
oci_execute($erp_sql2);
while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
    //寫單身
    $querydd = "insert into delayv4 ( tdate, ttime, orderdate, orderno, seq, ticketno, rx, clientid, clientname, productcode, productname, qty, stepseq, stepid, stepname, checkinid, checkinname,
                                      checkindate, checkintime, checkinmakerid, checkinmakername, checkoutid, checkoutname, checkoutdate, checkouttime, checkoutmakerid, checkoutmakername, ttype ) values (
               '" . $tdate                        . "',
               '" . $ttime                        . "',
               '" . $row2['OEA02']                . "',
               '" . $row2['OEA01']                . "',
               '" . $row2['OEB03']                . "',
               '" . $row2['TC_SRG001']            . "',
               '" . $row2['TC_SRG002']            . "',
               '" . $row2['OCC01']                . "',
               '" . $row2['OCC02']                . "',
               '" . $row2['TC_SRG003']            . "',
               '" . $row2['IMA02']                . "',
               '" . $row2['SFB08']                . "',
               '" . $row2['TC_SRG004']            . "',
               '" . $row2['ECB06']                . "',
               '" . $row2['ECB17']                . "',
               '" . $row2['TC_SRG009']            . "',
               '" . $row2['GEN021']               . "',
               '" . $row2['TC_SRG007']            . "',
               '" . $row2['TC_SRG008']            . "',
               '" . $row2['GEN031']               . "',
               '" . $row2['GEM021']               . "',
               '" . $row2['TC_SRG012']            . "',
               '" . $row2['GEN022']               . "',
               '" . $row2['TC_SRG010']            . "',
               '" . $row2['TC_SRG011']            . "',
               '" . $row2['GEN032']               . "',
               '" . $row2['GEM022']               . "','2' ) ";

    $resultdd = mysql_query($querydd) or die('226 delayv4 insert error!! ' . mysql_error());

}


//昨天返工報工
$s2 ="select to_char(oea02,'yyyy-mm-dd') oea02, oea01, oeb03, occ01, occ02, tc_srg001, tc_srg002, tc_srg003, ima02, tc_srg004, ecb06, ecb17, sfb08, to_char(tc_srg019, 'yyyy-mm-dd') tc_srg019, tc_srg020, tc_srg021, a.gen02 gen021, a.gen03 gen031, c.gem02 gem021, to_char(tc_srg022, 'yyyy-mm-dd') tc_srg022, tc_srg023, tc_srg024, b.gen02 gen022, b.gen03 gen032, d.gem02 gem022
      from tc_srg_file, ima_file, ecb_file, gen_file a , gen_file b, sfb_file, oeb_file, oea_file, occ_file, gem_file c, gem_file d
      where tc_srg022 = to_date('$ydate','yy/mm/dd')
      and tc_srg001=sfb01 and sfb22=oeb01 and sfb221=oeb03 and oeb01=oea01
      and tc_srg003=ima01
      and tc_srg003=ecb01 and tc_srg004=ecb03
      and tc_srg021=a.gen01
      and a.gen03=c.gem01
      and tc_srg024=b.gen01
      and b.gen03=d.gem01
      and oea04=occ01
      and ecb06!='9999'
      order by tc_srg001 " ;


$erp_sql2 = oci_parse($erp_conn1,$s2 );
oci_execute($erp_sql2);
while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
    //寫單身
    $querydd = "insert into delayv4 ( tdate, ttime, orderdate, orderno, seq, ticketno, rx, clientid, clientname, productcode, productname, qty, stepseq, stepid, stepname, checkinid, checkinname, checkindate, checkintime, checkinmakerid, checkinmakername, checkoutid, checkoutname, checkoutdate, checkouttime, checkoutmakerid, checkoutmakername, ttype ) values (
               '" . $tdate                        . "',
               '" . $ttime                        . "',
               '" . $row2['OEA02']                . "',
               '" . $row2['OEA01']                . "',
               '" . $row2['OEB03']                . "',
               '" . $row2['TC_SRG001']            . "',
               '" . $row2['TC_SRG002']            . "',
               '" . $row2['OCC01']                . "',
               '" . $row2['OCC02']                . "',
               '" . $row2['TC_SRG003']            . "',
               '" . $row2['IMA02']                . "',
               '" . $row2['SFB08']                . "',
               '" . $row2['TC_SRG004']            . "',
               '" . $row2['ECB06']                . "',
               '" . $row2['ECB17']                . "',
               '" . $row2['TC_SRG021']            . "',
               '" . $row2['GEN021']               . "',
               '" . $row2['TC_SRG019']            . "',
               '" . $row2['TC_SRG020']            . "',
               '" . $row2['GEN031']               . "',
               '" . $row2['GEM021']               . "',
               '" . $row2['TC_SRG024']            . "',
               '" . $row2['GEN022']               . "',
               '" . $row2['TC_SRG022']            . "',
               '" . $row2['TC_SRG023']            . "',
               '" . $row2['GEN032']               . "',
               '" . $row2['GEM022']               . "','2' ) ";

    $resultdd = mysql_query($querydd) or die('273 delayv4 insert error!! ' . mysql_error());

}

//到此 昨天報工出站的資料都有了 開始計算應做時間 及 負責人

$query="select * from delayv4 where tdate='$tdate' and ttype='2' order by pkey";
$result=mysql_query($query) or die ('125 Delay read error!! '.mysql_error());
while ($row=mysql_fetch_array($result)) {
  $pkey=$row['pkey'];
  $stepid=$row['stepid'];
  $qty=$row['qty'];
  $ttdate=$row['checkoutdate'];
  $tttime=$row['checkouttime'];
  $checkindate=$row['checkindate'];
  $checkintime=$row['checkintime'];
  $checkoutmakerid=$row['checkoutmakerid'];
  $checkoutmakername=$row['checkoutmakername'];
  $checkinid=$row['checkinid'];
  $checkinname=$row['checkinname'];

  $querys="select time1, qty2, time2 from step where id='$stepid'";
  $results = mysql_query($querys) or die ('138 Step read error!! ' .  mysql_error());
  $rows=mysql_fetch_array($results);
  if ($qty >=$rows['qty2'])  {
      $steptime=$rows['time2'];
  } else {
      $steptime=$rows['time1'];
  }

  $querym="select makerid, makername, managerid, managername from stepmanager where makerid='$checkoutmakerid' and stepid='$stepid' ";
  $resultm = mysql_query($querym) or die ('147 StepManager read error!! ' .  mysql_error());
  $rowm=mysql_fetch_array($resultm);
  if(is_null($rowm['managerid'])) {  //沒有指定負責人 則為本人   ]
    $managermakerid=$checkoutmakerid;
    $managermakername=$checkoutmakername;
    $managerid=$checkinid;
    $managername=$checkinname;
  } else {
    $managermakerid=$rowm['makerid'];
    $managermakername=$rowm['makername'];
    $managerid=$rowm['managerid'];
    $managername=$rowm['managername'];
  }
  $makingtime=timeDiff($ttdate, $tttime, $checkindate, $checkintime);

  $overtime=$makingtime - $steptime;


  $queryu = "update delayv4 set
                  managermakerid    = '" . $managermakerid    . "',
                  managermakername  = '" . $managermakername  . "',
                  managerid         = '" . $managerid    . "',
                  managername       = '" . $managername  . "',
                  makingtime        = '" . $makingtime   . "',
                  steptime          = '" . $steptime     . "',
                  overtime          = '" . $overtime     . "'
                  where pkey        = '" . $pkey         . "' limit 1";

  $resultu = mysql_query($queryu) or die ('195 Delayv4 Update error!! ' . $queryu . '  ' . mysql_error());
}



//寄email了
error_reporting(E_ALL);
require_once 'classes/PHPExcel.php';
require_once 'classes/PHPExcel/IOFactory.php';
//$objPHPExcel = new PHPExcel();
$objReader = PHPExcel_IOFactory::createReader('Excel5');
$objPHPExcel = $objReader->load("templates/erp_delaycasesv4.xls");
// Set properties
$objPHPExcel ->getProperties()->setCreator("Frank")
           ->setLastModifiedBy("Frank")
           ->setTitle("$tdate 製作超時CASES")
           ->setSubject("$tdate 製作超時CASES#")
           ->setDescription("$tdate 製作超時CASES#")
           ->setKeywords("$tdate 應出貨但未出貨的RX#")
           ->setCategory("$tdate 製作超時CASES#");

// Add some data
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A1', $tdate . ' 各工序製作超時CASES ');
$objPHPExcel->setActiveSheetIndex(0) ->mergeCells('A1:Q1');

$query2= "select * from delayv4 where managermakerid!='680000' and  tdate='$tdate' and overtime > 0 order by managermakerid, stepid, clientid, rx ";
$result2= mysql_query($query2);

$y=2;
$i=0;
$j=0;
$oldrxno='';
$oldmanagermakerid='';
while ($row2 = mysql_fetch_array($result2)){
    if ($oldmanagermakerid!=$row2['managermakerid']) {
        $oldmanagermakerid =$row2['managermakerid'];
        $i=0;  //不同製處就由1開始算
        $y+=2; //列印位置
    }

    if ($oldrxno!=$row2['rx']){
        $rxno=$row2['rx'];
        $orderno=$row2['orderno'];
        $seq=$row2['seq'];
        $oldrxno=$row2['rx'];
        $i++;
        $j++;
        $di=$i;
    } else {
        $rxno='';
        $orderno='';
        $seq='';
        $di='';
    }

    if ($row2['ttype']=='1') {
        $tttime= $row2['tdate'].' '.$row2['ttime'] ; //還沒刷出
    } else {
        $tttime='';
    }

    $objPHPExcel->setActiveSheetIndex(0)
                ->setCellValue('A'. $y, $di)
                ->setCellValue('B'. $y, $row2['managermakername'])
                ->setCellValue('C'. $y, $row2['stepname'])
                ->setCellValue('D'. $y, $row2['managername'])
                ->setCellValue('E'. $y, $row2['clientid'].'--'.$row2['clientname'])
                ->setCellValue('F'. $y, $rxno)
                ->setCellValue('G'. $y, $orderno.' ' . $seq)
                ->setCellValue('H'. $y, $row2['ticketno'] . ' ')
                ->setCellValue('I'. $y, $row2['orderdate'])
                ->setCellValue('J'. $y, $row2['productcode'].'--'.$row2['productname'])
                ->setCellValue('K'. $y, $row2['qty'])
                ->setCellValue('L'. $y, $row2['checkindate'].' '.$row2['checkintime'])
                ->setCellValue('M'. $y, $row2['checkoutdate'].' '.$row2['checkouttime'])
                ->setCellValue('N'. $y, $tttime)
                ->setCellValue('O'. $y, $row2['makingtime'])
                ->setCellValue('P'. $y, $row2['steptime'])
                ->setCellValue('Q'. $y, $row2['overtime']) ;
    if (($y%2)==0){
        $objPHPExcel->setActiveSheetIndex(0) ->getStyle('A'.$y.':Q'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
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
$objPHPExcel->setActiveSheetIndex(0)->setTitle($tdate. ' 製作超時CASES');




// 各工序製作分鐘數
$objPHPExcel->setActiveSheetIndex(1)
            ->setCellValue('A1', '各工序製作分鐘數 ');

$query2= "select * from step order by id ";
$result2= mysql_query($query2);
$y=4;
while ($row2 = mysql_fetch_array($result2)){
  $objPHPExcel->setActiveSheetIndex(1)
                  ->setCellValue('A'. $y, $row2['maker'])
                  ->setCellValue('B'. $y, $row2['id'])
                  ->setCellValue('C'. $y, $row2['name'])
                  ->setCellValue('D'. $y, $row2['time1'])
                  ->setCellValue('E'. $y, $row2['qty2'])
                  ->setCellValue('F'. $y, $row2['time2']);
  $y++;

}
$y++;
$objPHPExcel->setActiveSheetIndex(0)
          ->setCellValue('A'. $y, '-- 以下無資料 --');
$objPHPExcel->setActiveSheetIndex(1)->setTitle('各工序製作分鐘數');

// 各工序製作分鐘數
$objPHPExcel->setActiveSheetIndex(2)
            ->setCellValue('A1', '各工序負責人 ');

$query2= "select * from stepmanager order by makerid, stepid ";
$result2= mysql_query($query2);
$y=4;
while ($row2 = mysql_fetch_array($result2)){
  $objPHPExcel->setActiveSheetIndex(2)
                  ->setCellValue('A'. $y, $row2['makername'])
                  ->setCellValue('B'. $y, $row2['stepid'])
                  ->setCellValue('C'. $y, $row2['stepname'])
                  ->setCellValue('D'. $y, $row2['managerid']. ' ')
                  ->setCellValue('E'. $y, $row2['managername']);
  $y++;

}
$y++;
$objPHPExcel->setActiveSheetIndex(0)
          ->setCellValue('A'. $y, '-- 以下無資料 --');
$objPHPExcel->setActiveSheetIndex(2)->setTitle('各工序負責人');


// Rename sheet
$objPHPExcel->setActiveSheetIndex(0)->setTitle($tdate. ' 製作超時CASES');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
// $filename='email/' . $tdate . '_DelayCasesReportv3.xls';
// $objWriter->save($filename);
$filename='report/' . $tdate . '_DelayCasesReportv4.xls';
$objWriter->save($filename);


$report_name = $tdate . ' 各工序製作超時CASES';

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
$mail->addAddress('it@veden.dental', 'Veden');
$mail->addAddress('m@veden.dental', 'Veden');
$mail->addAddress('sales@veden.dental', 'Veden');
$mail->addAddress('cs@veden.dental', 'Veden');
$mail->Subject = "Veden $report_name ";
$mail->Body    = "Veden $report_name ";
$mail->addAttachment($filename, $filename);
$mail->send();


?>
