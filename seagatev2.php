<?php
    // 海關每日物料數字導出
    session_start();
    $pagetitle = "報關組 &raquo; 海關每日物料數字導出V2";
    include("_data.php");
    //include("_erp.php");
    //auth("seagate.php");

    if (is_null($_GET['bdate'])) {
      $bdate =  date('Y-m-d');
    } else {
      $bdate=$_GET['bdate'];
    }

    if (is_null($_GET['edate'])) {
      $edate =  date('Y-m-d');
    } else {
      $edate=$_GET['edate'];
    }


    if ($_GET["submit"]=="Export") {

        //計算數字
        $query="delete from seagate_dailys";
        $result=mysql_query($query); //清空所有資料


        $query1="select * from seagate_exports where tdate>='$bdate' and tdate<='$edate'";
        $result1= mysql_query($query1) or die ('48 Exports load error');
        while ($row1 = mysql_fetch_array($result1)) {
            $tdate    = $row1['tdate'];
            $no       = $row1['no'];
            $weight   = $row1['weight']*1000;
            $contract = $row1['contract'];

            //取出成份
            $query2   = "select * from seagate_items where no='$no' and contact='$contract'";
            $result2  = mysql_query($query2) or die ('56 Items load error');
            $row2     = mysql_fetch_array($result2);

            //計算重量
            $query3   = "insert into seagate_dailys ( no, contract, tdate, weight, s011, s012, s013, s021, s022, s023, s031, s032, s033, s041, s042, s043, s051, s052, s053, s061, s062, s063,
                          s071, s072, s073, s081, s082, s083, s091, s092, s093, s101, s102, s103, s111, s112, s113, s121, s122, s123, s131, s132, s133, s141, s142, s143, s151, s152, s153, s161, s162, s163,
                          s171, s172, s173, s181, s182, s183, s191, s192, s193, s201, s202, s203, s211, s212, s213 ) values (
                  '" . $no              . "',
                  '" . $contract        . "',
                  '" . $tdate           . "',
                  " .  $weight          . ",
                  " .  $row2['s011']    . ",
                  " .  $row2['s012']    . ",
                  " . round( ($weight * $row2['s011'])/( 1-$row2['s012']/100 ) ,6)  . ",
                  " .  $row2['s021']    . ",
                  " .  $row2['s022']    . ",
                  " . round( ($weight * $row2['s021'])/( 1-$row2['s022']/100 ) ,6)  . ",
                  " .  $row2['s031']    . ",
                  " .  $row2['s032']    . ",
                  " . round( ($weight * $row2['s031'])/( 1-$row2['s032']/100 ) ,6)  . ",
                  " .  $row2['s041']    . ",
                  " .  $row2['s042']    . ",
                  " . round( ($weight * $row2['s041'])/( 1-$row2['s042']/100 ) ,6)  . ",
                  " .  $row2['s051']    . ",
                  " .  $row2['s052']    . ",
                  " . round( ($weight * $row2['s051'])/( 1-$row2['s052']/100 ) ,6)  . ",
                  " .  $row2['s061']    . ",
                  " .  $row2['s062']    . ",
                  " . round( ($weight * $row2['s061'])/( 1-$row2['s062']/100 ) ,6)  . ",
                  " .  $row2['s071']    . ",
                  " .  $row2['s072']    . ",
                  " . round( ($weight * $row2['s071'])/( 1-$row2['s072']/100 ) ,6)  . ",
                  " .  $row2['s081']    . ",
                  " .  $row2['s082']    . ",
                  " . round( ($weight * $row2['s081'])/( 1-$row2['s082']/100 ) ,6)  . ",
                  " .  $row2['s091']    . ",
                  " .  $row2['s092']    . ",
                  " . round( ($weight * $row2['s091'])/( 1-$row2['s092']/100 ) ,6)  . ",
                  " .  $row2['s101']    . ",
                  " .  $row2['s102']    . ",
                  " . round( ($weight * $row2['s101'])/( 1-$row2['s102']/100 ) ,6)  . ",
                  " .  $row2['s111']    . ",
                  " .  $row2['s112']    . ",
                  " . round( ($weight * $row2['s111'])/( 1-$row2['s112']/100 ) ,6)  . ",
                  " .  $row2['s121']    . ",
                  " .  $row2['s122']    . ",
                  " . round( ($weight * $row2['s121'])/( 1-$row2['s122']/100 ) ,6)  . ",
                  " .  $row2['s121']    . ",
                  " .  $row2['s132']    . ",
                  " . round( ($weight * $row2['s131'])/( 1-$row2['s132']/100 ) ,6)  . ",
                  " .  $row2['s141']    . ",
                  " .  $row2['s142']    . ",
                  " . round( ($weight * $row2['s141'])/( 1-$row2['s142']/100 ) ,6)  . ",
                  " .  $row2['s151']    . ",
                  " .  $row2['s152']    . ",
                  " . round( ($weight * $row2['s151'])/( 1-$row2['s152']/100 ) ,6)  . ",
                  " .  $row2['s161']    . ",
                  " .  $row2['s162']    . ",
                  " . round( ($weight * $row2['s161'])/( 1-$row2['s162']/100 ) ,6)  . ",
                  " .  $row2['s171']    . ",
                  " .  $row2['s172']    . ",
                  " . round( ($weight * $row2['s171'])/( 1-$row2['s172']/100 ) ,6)  . ",
                  " .  $row2['s181']    . ",
                  " .  $row2['s182']    . ",
                  " . round( ($weight * $row2['s181'])/( 1-$row2['s182']/100 ) ,6)  . ",
                  " .  $row2['s191']    . ",
                  " .  $row2['s192']    . ",
                  " . round( ($weight * $row2['s191'])/( 1-$row2['s192']/100 ) ,6)  . ",
                  " .  $row2['s201']    . ",
                  " .  $row2['s202']    . ",
                  " . round( ($weight * $row2['s201'])/( 1-$row2['s202']/100 ) ,6)  . ",
                  " .  $row2['s211']    . ",
                  " .  $row2['s212']    . ",
                  " . round( ($weight * $row2['s211'])/( 1-$row2['s212']/100 ) ,6)  . ")";
            $result3= mysql_query($query3) or die ('116 dailys added error.==>' .$query3 . '<=='.mysql_error());
        }
        commit;

        //匯出
        $filename="templates/seagatev2.xls";

        require_once 'classes/PHPExcel.php';
        require_once 'classes/PHPExcel/IOFactory.php';
        require_once 'classes/PHPExcel/Cell/AdvancedValueBinder.php';
        PHPExcel_Cell::setValueBinder( new PHPExcel_Cell_AdvancedValueBinder() );
        $objReader = PHPExcel_IOFactory::createReader('Excel5');
        $objPHPExcel = $objReader->load("$filename");
        //$objPHPExcel = new PHPExcel();
        // Set properties

        $objPHPExcel ->getProperties()->setCreator( 'seagate')
                     ->setLastModifiedBy('seagate')
                     ->setTitle('seagate')
                     ->setSubject('seagate')
                     ->setDescription('seagate')
                     ->setKeywords('seagate')
                     ->setCategory('seagate');
        $objPHPExcel->setActiveSheetIndex(0);
        $osheet = $objPHPExcel->getActiveSheet();
        $y=3;

        $s[01]='石膏牙模';
        $s[02]='硅橡膠牙模';
        $s[03]='牙科镍铬合金';
        $s[04]='牙科鈷铬合金';
        $s[05]='牙科鈷铬合金';
        $s[06]='锻造牙科钯合金1（含鈀60.6%）';
        $s[07]='锻造牙科钯合金2（含金75.1%）';
        $s[08]='锻造牙科钯合金3（含金40%）';
        $s[09]='锻造牙科钯合金4（含鈀39.4%）';
        $s[10]='锻造牙科钯合金5（含鈀25%）';
        $s[11]='瓷牙粉';
        $s[12]='瓷粉專用水';
        $s[13]='瓷牙底色膏';
        $s[14]='樹脂粒';
        $s[15]='锻造牙科金合金片(含金97%，银3%)CG+CBG';
        $s[16]='锻造牙科金合金粉(含金97%，银3%)Inflow';
        $s[17]='锻造牙科金合金片（含金63.1%）cp+cbp';
        $s[18]='锻造牙科金合金粉（含金95.13%）PONTIC';
        $s[19]='锻造牙科金合金粉（含金94.22%）Inconnect';
        $s[20]='锻造牙科金合金粉（含金92.72%）UCP';
        $s[21]='压铸瓷块';

        $query="select tdate, sum(s013) s013, sum(s023) s023, sum(s033) s033, sum(s043) s043, sum(s053) s053, sum(s063) s063,
                sum(s073) s073, sum(s083) s083, sum(s093) s093, sum(s103) s103, sum(s113) s113, sum(s123) s123,
                sum(s133) s133, sum(s143) s143, sum(s153) s153, sum(s163) s163, sum(s173) s173, sum(s183) s183,
                sum(s193) s193, sum(s203) s203, sum(s213) s213
                from seagate_dailys
                group by tdate";
                //order by tdate no";
        $result= mysql_query($query) or die ('152 dailys load error. ' .mysql_error());
        while ($row = mysql_fetch_array($result)) {
            $tdate=$row['tdate'];


            if ($row['s013'] > 0) {
                $osheet ->setCellValue('A'. $y, $row['tdate']);
                $osheet ->setCellValue('B'. $y, '01');
                $osheet ->setCellValue('C'. $y, $s[01]);
                $osheet ->setCellValue('D'. $y, $row['s013']);
                $osheet ->getStyle('A'.$y)
                        ->getNumberFormat()
                        ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);
                $y++;
            }
            if ($row['s023'] > 0) {
                $osheet ->setCellValue('A'. $y, $row['tdate']);
                $osheet ->setCellValue('B'. $y, '02');
                $osheet ->setCellValue('C'. $y, $s[02]);
                $osheet ->setCellValue('D'. $y, $row['s023']);
                $osheet ->getStyle('A'.$y)
                        ->getNumberFormat()
                        ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);
                $y++;
            }
            if ($row['s033'] > 0) {
                $osheet ->setCellValue('A'. $y, $row['tdate']);
                $osheet ->setCellValue('B'. $y, '03');
                $osheet ->setCellValue('C'. $y, $s[03]);
                $osheet ->setCellValue('D'. $y, $row['s033']);
                $osheet ->getStyle('A'.$y)
                        ->getNumberFormat()
                        ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);
                $y++;
            }
            if ($row['s043'] > 0) {
                $osheet ->setCellValue('A'. $y, $row['tdate']);
                $osheet ->setCellValue('B'. $y, '04');
                $osheet ->setCellValue('C'. $y, $s[04]);
                $osheet ->setCellValue('D'. $y, $row['s043']);
                $osheet ->getStyle('A'.$y)
                        ->getNumberFormat()
                        ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);
                $y++;
            }
            if ($row['s053'] > 0) {
                $osheet ->setCellValue('A'. $y, $row['tdate']);
                $osheet ->setCellValue('B'. $y, '05');
                $osheet ->setCellValue('C'. $y, $s[05]);
                $osheet ->setCellValue('D'. $y, $row['s053']);
                $osheet ->getStyle('A'.$y)
                        ->getNumberFormat()
                        ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);
                $y++;
            }
            if ($row['s063'] > 0) {
                $osheet ->setCellValue('A'. $y, $row['tdate']);
                $osheet ->setCellValue('B'. $y, '06');
                $osheet ->setCellValue('C'. $y, $s[06]);
                $osheet ->setCellValue('D'. $y, $row['s063']);
                $osheet ->getStyle('A'.$y)
                        ->getNumberFormat()
                        ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);
                $y++;
            }
            if ($row['s073'] > 0) {
                $osheet ->setCellValue('A'. $y, $row['tdate']);
                $osheet ->setCellValue('B'. $y, '07');
                $osheet ->setCellValue('C'. $y, $s[07]);
                $osheet ->setCellValue('D'. $y, $row['s073']);
                $osheet ->getStyle('A'.$y)
                        ->getNumberFormat()
                        ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);
                $y++;
            }
            if ($row['s083'] > 0) {
                $osheet ->setCellValue('A'. $y, $row['tdate']);
                $osheet ->setCellValue('B'. $y, '08');
                $osheet ->setCellValue('C'. $y, $s[08]);
                $osheet ->setCellValue('D'. $y, $row['s083']);
                $osheet ->getStyle('A'.$y)
                        ->getNumberFormat()
                        ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);
                $y++;
            }
            if ($row['s093'] > 0) {
                $osheet ->setCellValue('A'. $y, $row['tdate']);
                $osheet ->setCellValue('B'. $y, '09');
                $osheet ->setCellValue('C'. $y, $s[09]);
                $osheet ->setCellValue('D'. $y, $row['s093']);
                $osheet ->getStyle('A'.$y)
                        ->getNumberFormat()
                        ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);
                $y++;
            }
            if ($row['s103'] > 0) {
                $osheet ->setCellValue('A'. $y, $row['tdate']);
                $osheet ->setCellValue('B'. $y, '10');
                $osheet ->setCellValue('C'. $y, $s[10]);
                $osheet ->setCellValue('D'. $y, $row['s103']);
                $osheet ->getStyle('A'.$y)
                        ->getNumberFormat()
                        ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);
                $y++;
            }
            if ($row['s113'] > 0) {
                $osheet ->setCellValue('A'. $y, $row['tdate']);
                $osheet ->setCellValue('B'. $y, '11');
                $osheet ->setCellValue('C'. $y, $s[11]);
                $osheet ->setCellValue('D'. $y, $row['s113']);
                $osheet ->getStyle('A'.$y)
                        ->getNumberFormat()
                        ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);
                $y++;
            }
            if ($row['s123'] > 0) {
                $osheet ->setCellValue('A'. $y, $row['tdate']);
                $osheet ->setCellValue('B'. $y, '12');
                $osheet ->setCellValue('C'. $y, $s[12]);
                $osheet ->setCellValue('D'. $y, $row['s123']);
                $osheet ->getStyle('A'.$y)
                        ->getNumberFormat()
                        ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);
                $y++;
            }
            if ($row['s133'] > 0) {
                $osheet ->setCellValue('A'. $y, $row['tdate']);
                $osheet ->setCellValue('B'. $y, '13');
                $osheet ->setCellValue('C'. $y, $s[13]);
                $osheet ->setCellValue('D'. $y, $row['s133']);
                $osheet ->getStyle('A'.$y)
                        ->getNumberFormat()
                        ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);
                $y++;
            }
            if ($row['s143'] > 0) {
                $osheet ->setCellValue('A'. $y, $row['tdate']);
                $osheet ->setCellValue('B'. $y, '14');
                $osheet ->setCellValue('C'. $y, $s[14]);
                $osheet ->setCellValue('D'. $y, $row['s143']);
                $osheet ->getStyle('A'.$y)
                        ->getNumberFormat()
                        ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);
                $y++;
            }
            if ($row['s153'] > 0) {
                $osheet ->setCellValue('A'. $y, $row['tdate']);
                $osheet ->setCellValue('B'. $y, '15');
                $osheet ->setCellValue('C'. $y, $s[15]);
                $osheet ->setCellValue('D'. $y, $row['s153']);
                $osheet ->getStyle('A'.$y)
                        ->getNumberFormat()
                        ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);
                $y++;
            }
            if ($row['s163'] > 0) {
                $osheet ->setCellValue('A'. $y, $row['tdate']);
                $osheet ->setCellValue('B'. $y, '16');
                $osheet ->setCellValue('C'. $y, $s[16]);
                $osheet ->setCellValue('D'. $y, $row['s163']);
                $osheet ->getStyle('A'.$y)
                        ->getNumberFormat()
                        ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);
                $y++;
            }
            if ($row['s173'] > 0) {
                $osheet ->setCellValue('A'. $y, $row['tdate']);
                $osheet ->setCellValue('B'. $y, '17');
                $osheet ->setCellValue('C'. $y, $s[17]);
                $osheet ->setCellValue('D'. $y, $row['s173']);
                $osheet ->getStyle('A'.$y)
                        ->getNumberFormat()
                        ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);
                $y++;
            }
            if ($row['s183'] > 0) {
                $osheet ->setCellValue('A'. $y, $row['tdate']);
                $osheet ->setCellValue('B'. $y, '18');
                $osheet ->setCellValue('C'. $y, $s[18]);
                $osheet ->setCellValue('D'. $y, $row['s183']);
                $osheet ->getStyle('A'.$y)
                        ->getNumberFormat()
                        ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);
                $y++;
            }
            if ($row['s193'] > 0) {
                $osheet ->setCellValue('A'. $y, $row['tdate']);
                $osheet ->setCellValue('B'. $y, '19');
                $osheet ->setCellValue('C'. $y, $s[19]);
                $osheet ->setCellValue('D'. $y, $row['s193']);
                $osheet ->getStyle('A'.$y)
                        ->getNumberFormat()
                        ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);
                $y++;
            }
            if ($row['s203'] > 0) {
                $osheet ->setCellValue('A'. $y, $row['tdate']);
                $osheet ->setCellValue('B'. $y, '20');
                $osheet ->setCellValue('C'. $y, $s[20]);
                $osheet ->setCellValue('D'. $y, $row['s203']);
                $osheet ->getStyle('A'.$y)
                        ->getNumberFormat()
                        ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);
                $y++;
            }
            if ($row['s213'] > 0) {
                $osheet ->setCellValue('A'. $y, $row['tdate']);
                $osheet ->setCellValue('B'. $y, '21');
                $osheet ->setCellValue('C'. $y, $s[21]);
                $osheet ->setCellValue('D'. $y, $row['s213']);
                $osheet ->getStyle('A'.$y)
                        ->getNumberFormat()
                        ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);
                $y++;
            }
        }

        // Rename sheet
        $osheet ->setTitle('Daily Export');
        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $objPHPExcel->setActiveSheetIndex(0);

        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="'. 'dailyexport_' . $bdate . '_' . $edate . '.xls"');
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
  $IsAjax = False;
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>海關每日物料數字導出 </p>
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
         出貨期間:
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> ~~
        <input name="edate" type="text" id="edate" size="12" maxlength="12" value=<?=$edate;?> onfocus="new WdatePicker()">
        &nbsp;&nbsp;   &nbsp;&nbsp;
        <input type="submit" name="submit" id="submit" value="Export">  &nbsp;&nbsp;   &nbsp;&nbsp;
      </td></tr>
    </table>
  </div>
</form>
