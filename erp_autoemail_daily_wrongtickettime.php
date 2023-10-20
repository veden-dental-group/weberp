<?
//檢查當天報工的時間在10分鐘內的
session_start();
include("_data.php");
date_default_timezone_set('Asia/Taipei');

// $edate = date('Y-m-d', strtotime("-1 days"));
$edate = $argv[1] ? $argv[1] : date('Y-m-d', strtotime("-1 days"));

$check_interval = 10 * 60;

error_reporting(E_ALL);
require_once 'classes/PHPExcel.php';
require_once 'classes/PHPExcel/IOFactory.php';

// begin wrong ticket time list
//報工清單
$objPHPExcel = new PHPExcel();
// Set properties
$objPHPExcel ->getProperties()->setCreator("Report")
           ->setLastModifiedBy("Report")
           ->setTitle("報工有問題之工單")
           ->setSubject("報工有問題之工單")
           ->setDescription("報工有問題之工單")
           ->setKeywords("報工有問題之工單")
           ->setCategory("報工有問題之工單");

// Add some data
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A1', $edate . ' 報工時, 同時刷進和刷出的工單 ');
$objPHPExcel->setActiveSheetIndex(0) ->mergeCells('A1:E1');
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A3', '製處')
            ->setCellValue('B3', '客戶')
            ->setCellValue('C3', 'RX #')
            ->setCellValue('D3', '工單號')
            ->setCellValue('E3', '工序')
            ->setCellValue('F3', '進站時間')
            ->setCellValue('G3', '出站時間')
            ->setCellValue('H3', '品名');

// modified by Mao 20130806
$s2 = "select sfb82,gem02, oea04,occ02, sfbud02, tc_srg001,  ecb06,ecb17, tc_srg008, tc_srg011, sfb05, ima02 " .
     "from
	tc_srg_file, sfb_file,gem_file,ima_file,ecb_file,oea_file, occ_file " .
     "where
	tc_srg007=to_date('$edate','yy/mm/dd')
	and tc_srg010 is not null
	and
		(24 * 60 *
		( to_date (to_char(tc_srg010,'YYYY-MM-DD') || ' ' || tc_srg011,  'YYYY-MM-DD hh24:mi:ss') -
		to_date (to_char(tc_srg007,'YYYY-MM-DD') || ' ' || tc_srg008,  'YYYY-MM-DD hh24:mi:ss'))) < 10 " .
     "and tc_srg001=sfb01 " .
     "and sfb82=gem01 " .
     "and ecb06!='9999' " .
     "and sfb05=ima01 " .
     "and sfb05=ecb01 and tc_srg004=ecb03 " .
     "and sfb22=oea01 " .
     "and oea04=occ01 " .
     "order by sfb82, oea04, tc_srg001,tc_srg004 ";

/*
$s2= "select sfb82,gem02, oea04,occ02, sfbud02, tc_srg001,  ecb06,ecb17, tc_srg008, tc_srg011, sfb05,ima02 " .
     "from tc_srg_file, sfb_file,gem_file,ima_file,ecb_file,oea_file, occ_file " .
     "where tc_srg007=to_date('$edate','yy/mm/dd') and tc_srg010 =to_date('$edate','yy/mm/dd') and substr(tc_srg008,1,4)=substr(tc_srg011,1,4) " .
     "and tc_srg001=sfb01 " .
     "and sfb82=gem01 " .
     "and ecb06!='9999' " .
     "and sfb05=ima01 " .
     "and sfb05=ecb01 and tc_srg004=ecb03 " .
     "and sfb22=oea01 " .
     "and oea04=occ01 " .
     "order by sfb82, oea04, tc_srg001,tc_srg004 ";
*/

$erp_sql2 = oci_parse($erp_conn1,$s2 );
oci_execute($erp_sql2);
$y=4;
while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
	$objPHPExcel->setActiveSheetIndex(0)
				->setCellValue('A'. $y, $row2['GEM02'])
				->setCellValue('B'. $y, $row2['OEA04'].' ' . $row2['OCC02'])
				->setCellValue('C'. $y, $row2['SFBUD02'])
				->setCellValue('D'. $y, $row2['TC_SRG001'])
				->setCellValue('E'. $y, $row2['ECB17'])
				->setCellValue('F'. $y, $row2['TC_SRG008'])
				->setCellValue('G'. $y, $row2['TC_SRG011'])
				->setCellValue('H'. $y, $row2['SFB05'].' ' . $row2['IMA02']);
	$y++;
}

$objPHPExcel->setActiveSheetIndex(0)
          ->setCellValue('A'. $y, '-- 以下無資料 --');
// Rename sheet
$objPHPExcel->getActiveSheet(0)->setTitle('報工有問題之工單 ');
// $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');

// $filename1='email/' . $edate . '_wrongtickettimelist.xls';
// $objWriter->save($filename1);
// $filename1='report/' . $edate . '_wrongtickettimelist.xls';
// $objWriter->save($filename1);
// end wrong ticket time list

// begin wrong ticket time total
//合計數
// $objPHPExcel = new PHPExcel();
// Set properties
/*
$objPHPExcel ->getProperties()->setCreator("Frank")
           ->setLastModifiedBy("Frank")
           ->setTitle("報工有問題之工單合計")
           ->setSubject("報工有問題之工單合計")
           ->setDescription("報工有問題之工單合計")
           ->setKeywords("報工有問題之工單合計")
           ->setCategory("報工有問題之工單合計");
*/

$sheet = 1;
// Add some data
$objPHPExcel->createSheet();
// $objPHPExcel->setActiveSheetIndex(1);

$objPHPExcel->setActiveSheetIndex($sheet)
            ->setCellValue('A1', $edate . ' 同時刷進和刷出的合計數量');
$objPHPExcel->setActiveSheetIndex($sheet) ->mergeCells('A1:E1');
$objPHPExcel->setActiveSheetIndex($sheet)
            ->setCellValue('A3', '製處')
            ->setCellValue('B3', '工序')
            ->setCellValue('C3', '數量');

$s2 = "
select
    sfb82, gem02, ecb06, ecb17, count(*) total
from
    (
    select sfb82 , gem02,  sfbud02 , tc_srg001 ,  tc_srg004, ecb06 ,ecb17, to_char(tc_srg007,'mm-dd') , tc_srg008 , to_char(tc_srg010,'mm-dd') , tc_srg011, sfb05, ima02
    from tc_srg_file, sfb_file,gem_file,ima_file,ecb_file
    where tc_srg007=to_date('$edate','yy/mm/dd')
    and tc_srg010 is not null
        and
        (24 * 60 *
        ( to_date (to_char(tc_srg010,'YYYY-MM-DD') || ' ' || tc_srg011,
'YYYY-MM-DD hh24:mi:ss') -
        to_date (to_char(tc_srg007,'YYYY-MM-DD') || ' ' || tc_srg008,  '
YYYY-MM-DD hh24:mi:ss'))) <= 10
    and tc_srg001=sfb01
    and sfb82=gem01
    and ecb06!='9999'
    and sfb05=ima01
    and sfb05=ecb01
    and tc_srg004=ecb03
    order by sfb82, tc_srg001,tc_srg004
)
group by sfb82, gem02, ecb06, ecb17
order by sfb82, gem02, ecb06, ecb17
";

/*
$s2= "select sfb82, gem02, ecb06, ecb17, count(*) total from " .
     "(select sfb82 , gem02,  sfbud02 , tc_srg001 ,  tc_srg004, ecb06 ,ecb17, to_char(tc_srg007,'mm-dd') , tc_srg008 , to_char(tc_srg010,'mm-dd') , tc_srg011, sfb05, ima02 " .
     "from tc_srg_file, sfb_file,gem_file,ima_file,ecb_file " .
     "where tc_srg007=to_date('$edate','yy/mm/dd') and tc_srg010 =to_date('$edate','yy/mm/dd') and substr(tc_srg008,1,4)=substr(tc_srg011,1,4) " .
     "and tc_srg001=sfb01 and sfb82=gem01 and ecb06!='9999' and sfb05=ima01 and sfb05=ecb01 and tc_srg004=ecb03 order by sfb82, tc_srg001,tc_srg004) " .
     "group by sfb82, gem02, ecb06, ecb17 order by sfb82, gem02, ecb06, ecb17 ";
*/

$erp_sql2 = oci_parse($erp_conn1,$s2 );
oci_execute($erp_sql2);
$y=4;
while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
    $objPHPExcel->setActiveSheetIndex($sheet)
                ->setCellValue('A'. $y, $row2['GEM02'])
                ->setCellValue('B'. $y, $row2['ECB17'])
                ->setCellValue('C'. $y, $row2['TOTAL']);
    $y++;
}


// Rename sheet
$objPHPExcel->getActiveSheet($sheet)->setTitle('報工有問題之工單合計');

// $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
// $filename2='email/' . $edate . '_wrongtickettimetotal.xls';
// $objWriter->save($filename2);
// $filename2='report/' . $edate . '_wrongtickettimetotal.xls';
// $objWriter->save($filename2);
// end wrong ticket time total

// begin wrong ticket time weekly
// 前10天的合計數 看有無進步
$tdate=$edate;
$i=14;
while ( $i>=1){
  $sdate[$i]=$tdate;
  $i--;
  $tdate=date('Y-m-d',strtotime("-1 day", strtotime($tdate)));
}
//因為已經多減一天了  要加回來
$bdate = date('Y-m-d',strtotime("+1 day", strtotime($tdate)));

/*
$objPHPExcel = new PHPExcel();
// Set properties
$objPHPExcel ->getProperties()->setCreator("Frank")
           ->setLastModifiedBy("Frank")
           ->setTitle("報工有問題之工單 依工序合計")
           ->setSubject("報工有問題之工單 依工序合計")
           ->setDescription("報工有問題之工單 依工序合計")
           ->setKeywords("報工有問題之工單 依工序合計")
           ->setCategory("報工有問題之工單 依工序合計");
*/

$sheet = 2;
// Add some data
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex($sheet)
            ->setCellValue('A1', $bdate . ' -- ' .$edate . ' 同時刷進和刷出的合計數量');
$objPHPExcel->setActiveSheetIndex($sheet) ->mergeCells('A1:E1');
$objPHPExcel->setActiveSheetIndex($sheet)
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
     "where tc_srg007 between to_date('$bdate','yy/mm/dd') and to_date('$edate','yy/mm/dd')
    and tc_srg010 is not null
    and
	(24 * 60 *
	( to_date (to_char(tc_srg010,'YYYY-MM-DD') || ' ' || tc_srg011, 'YYYY-MM-DD hh24:mi:ss') -
	to_date (to_char(tc_srg007,'YYYY-MM-DD') || ' ' || tc_srg008,  'YYYY-MM-DD hh24:mi:ss'))) <= 10 " .
     "and tc_srg001=sfb01 and sfb82=gem01 and sfb05=ecb01 and tc_srg004=ecb03 and ecb06!='9999' ) " .
     "group by gem02, tc_srg007, ecb17 )) group by gem02, ecb17 order by gem02,ecb17 ";

//echo $s2;
/*
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
     "where tc_srg007 between to_date('$bdate','yy/mm/dd')
	and to_date('$edate','yy/mm/dd')
	and tc_srg007=tc_srg010
	and substr(tc_srg008,1,4)=substr(tc_srg011,1,4) " .
     "and tc_srg001=sfb01 and sfb82=gem01 and sfb05=ecb01 and tc_srg004=ecb03 and ecb06!='9999' ) " .
     "group by gem02, tc_srg007, ecb17 )) group by gem02, ecb17 order by gem02,ecb17 ";
*/

$erp_sql2 = oci_parse($erp_conn1,$s2 );
oci_execute($erp_sql2);
$y=4;
$rec=0;
while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
    $objPHPExcel->setActiveSheetIndex($sheet)
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

$objPHPExcel->getActiveSheet($sheet)->setTitle('最近十四日內報工有問題之工單合計');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
// $filename3='email/' . $edate . '_wrongtickettimetotalweekly.xls';
// $objWriter->save($filename3);
$filename = 'report/' . $edate . '_wrongtickettime_all.xls';
$objWriter->save($filename);

require 'PHPMailer/PHPMailerAutoload.php';
$report_name = $edate . ' 同時刷進出的工單';

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
$mail->Subject = "Veden $report_name ";
$mail->Body = "Veden $report_name ";
$mail->addAttachment($filename, $filename);

$mail->send();

?>
