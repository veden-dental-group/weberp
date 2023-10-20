<?php
    // 海關每日物料數字導出
    session_start();
    $pagetitle = "報關組 &raquo; 海關每日物料數字導出";
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
            $weight   = $row1['weight'];
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
            $result3= mysql_query($query3) or die ('116 dailys added error. ' .mysql_error());
        }
        commit;

        //匯出
        if ($_GET['isdetails']=='Y') {
            $filename="templates/seagatey.xls";
        } else {
            $filename="templates/seagaten.xls";
        }
        error_reporting(E_ALL);

        require_once 'classes/PHPExcel.php';
        require_once 'classes/PHPExcel/IOFactory.php';
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

        $query="select tdate, no, sum(weight) weight, count(*) cc, sum(s011) s011, sum(s012) s012, sum(s013) s013, sum(s021) s021, sum(s022) s022, sum(s023) s023,  sum(s031) s031, sum(s032) s032, sum(s033) s033,
                sum(s041) s041, sum(s042) s042, sum(s043) s043, sum(s051) s051, sum(s052) s052,  sum(s053) s053, sum(s061) s061, sum(s062) s062, sum(s063) s063,
                sum(s071) s071, sum(s072) s072, sum(s073) s073, sum(s081) s081, sum(s082) s082,  sum(s083) s083, sum(s091) s091, sum(s092) s092, sum(s093) s093,
                sum(s101) s101, sum(s102) s102, sum(s103) s103, sum(s111) s111, sum(s112) s112,  sum(s113) s113, sum(s121) s121, sum(s122) s122, sum(s123) s123, sum(s131) s131, sum(s132) s132, sum(s133) s133,
                sum(s141) s141, sum(s142) s142, sum(s143) s143, sum(s151) s151, sum(s152) s152,  sum(s153) s153, sum(s161) s161, sum(s162) s162, sum(s163) s163,
                sum(s171) s171, sum(s172) s172, sum(s173) s173, sum(s181) s181, sum(s182) s182,  sum(s183) s183, sum(s191) s191, sum(s192) s192, sum(s193) s193,
                sum(s201) s201, sum(s202) s202, sum(s203) s203, sum(s211) s211, sum(s212) s212,  sum(s213) s213
                from seagate_dailys
                group by tdate, no";
                //order by tdate no";
        $result= mysql_query($query) or die ('152 dailys load error. ' .mysql_error());
        while ($row = mysql_fetch_array($result)) {
            if ($_GET['isdetails']=='Y') {
                $osheet ->setCellValue('A'. $y, $row['tdate'])
                      ->setCellValue('B'. $y, $row['no'])
                      ->setCellValue('D'. $y, number_format($row["weight"],1,'.',','))
                      ->setCellValue('E'. $y, $row["cc"])
                      ->setCellValue('F'. $y, number_format($row["s011"],6,'.',','))
                      ->setCellValue('G'. $y, number_format($row["s012"],6,'.',','))
                      ->setCellValue('H'. $y, number_format($row["s013"],6,'.',','))

                      ->setCellValue('I'. $y, number_format($row["s021"],6,'.',','))
                      ->setCellValue('J'. $y, number_format($row["s022"],6,'.',','))
                      ->setCellValue('K'. $y, number_format($row["s023"],6,'.',','))

                      ->setCellValue('L'. $y, number_format($row["s031"],6,'.',','))
                      ->setCellValue('M'. $y, number_format($row["s032"],6,'.',','))
                      ->setCellValue('N'. $y, number_format($row["s033"],6,'.',','))

                      ->setCellValue('O'. $y, number_format($row["s041"],6,'.',','))
                      ->setCellValue('P'. $y, number_format($row["s042"],6,'.',','))
                      ->setCellValue('Q'. $y, number_format($row["s043"],6,'.',','))

                      ->setCellValue('R'. $y, number_format($row["s051"],6,'.',','))
                      ->setCellValue('S'. $y, number_format($row["s052"],6,'.',','))
                      ->setCellValue('T'. $y, number_format($row["s053"],6,'.',','))

                      ->setCellValue('U'. $y, number_format($row["s061"],6,'.',','))
                      ->setCellValue('V'. $y, number_format($row["s062"],6,'.',','))
                      ->setCellValue('W'. $y, number_format($row["s063"],6,'.',','))

                      ->setCellValue('X'. $y, number_format($row["s071"],6,'.',','))
                      ->setCellValue('Y'. $y, number_format($row["s072"],6,'.',','))
                      ->setCellValue('Z'. $y, number_format($row["s073"],6,'.',','))

                      ->setCellValue('AA'. $y, number_format($row["s081"],6,'.',','))
                      ->setCellValue('AB'. $y, number_format($row["s082"],6,'.',','))
                      ->setCellValue('AC'. $y, number_format($row["s083"],6,'.',','))

                      ->setCellValue('AD'. $y, number_format($row["s091"],6,'.',','))
                      ->setCellValue('AE'. $y, number_format($row["s092"],6,'.',','))
                      ->setCellValue('AF'. $y, number_format($row["s093"],6,'.',','))

                      ->setCellValue('AG'. $y, number_format($row["s101"],6,'.',','))
                      ->setCellValue('AH'. $y, number_format($row["s102"],6,'.',','))
                      ->setCellValue('AI'. $y, number_format($row["s103"],6,'.',','))

                      ->setCellValue('AJ'. $y, number_format($row["s111"],6,'.',','))
                      ->setCellValue('AK'. $y, number_format($row["s112"],6,'.',','))
                      ->setCellValue('AL'. $y, number_format($row["s113"],6,'.',','))

                      ->setCellValue('AM'. $y, number_format($row["s121"],6,'.',','))
                      ->setCellValue('AN'. $y, number_format($row["s122"],6,'.',','))
                      ->setCellValue('AO'. $y, number_format($row["s123"],6,'.',','))

                      ->setCellValue('AP'. $y, number_format($row["s131"],6,'.',','))
                      ->setCellValue('AQ'. $y, number_format($row["s132"],6,'.',','))
                      ->setCellValue('AR'. $y, number_format($row["s133"],6,'.',','))

                      ->setCellValue('AS'. $y, number_format($row["s141"],6,'.',','))
                      ->setCellValue('AT'. $y, number_format($row["s142"],6,'.',','))
                      ->setCellValue('AU'. $y, number_format($row["s143"],6,'.',','))

                      ->setCellValue('AV'. $y, number_format($row["s151"],6,'.',','))
                      ->setCellValue('AW'. $y, number_format($row["s152"],6,'.',','))
                      ->setCellValue('AX'. $y, number_format($row["s153"],6,'.',','))

                      ->setCellValue('AY'. $y, number_format($row["s161"],6,'.',','))
                      ->setCellValue('AZ'. $y, number_format($row["s162"],6,'.',','))
                      ->setCellValue('BA'. $y, number_format($row["s163"],6,'.',','))

                      ->setCellValue('BB'. $y, number_format($row["s171"],6,'.',','))
                      ->setCellValue('BC'. $y, number_format($row["s172"],6,'.',','))
                      ->setCellValue('BD'. $y, number_format($row["s173"],6,'.',','))

                      ->setCellValue('BE'. $y, number_format($row["s181"],6,'.',','))
                      ->setCellValue('BF'. $y, number_format($row["s182"],6,'.',','))
                      ->setCellValue('BG'. $y, number_format($row["s183"],6,'.',','))

                      ->setCellValue('BH'. $y, number_format($row["s191"],6,'.',','))
                      ->setCellValue('BI'. $y, number_format($row["s192"],6,'.',','))
                      ->setCellValue('BJ'. $y, number_format($row["s193"],6,'.',','))

                      ->setCellValue('BK'. $y, number_format($row["s201"],6,'.',','))
                      ->setCellValue('BL'. $y, number_format($row["s202"],6,'.',','))
                      ->setCellValue('BM'. $y, number_format($row["s203"],6,'.',','))

                      ->setCellValue('BN'. $y, number_format($row["s211"],6,'.',','))
                      ->setCellValue('BO'. $y, number_format($row["s212"],6,'.',','))
                      ->setCellValue('BP'. $y, number_format($row["s213"],6,'.',','));

                $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                $osheet ->getStyle('B'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                $osheet ->getStyle('D'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                $osheet ->getStyle('E'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);

                if (($y%2)==0){
                    $osheet->getStyle('A'.$y.':BS'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                }
            } else {
                $osheet ->setCellValue('A'. $y, $row['tdate'])
                      ->setCellValue('B'. $y, $row['no'])
                      ->setCellValue('D'. $y, number_format($row["weight"],1,'.',','))
                      ->setCellValue('E'. $y, $row["cc"])
                      ->setCellValue('F'. $y, number_format($row["s013"],6,'.',','))
                      ->setCellValue('G'. $y, number_format($row["s023"],6,'.',','))
                      ->setCellValue('H'. $y, number_format($row["s033"],6,'.',','))
                      ->setCellValue('I'. $y, number_format($row["s043"],6,'.',','))
                      ->setCellValue('J'. $y, number_format($row["s053"],6,'.',','))
                      ->setCellValue('K'. $y, number_format($row["s063"],6,'.',','))
                      ->setCellValue('L'. $y, number_format($row["s073"],6,'.',','))
                      ->setCellValue('M'. $y, number_format($row["s083"],6,'.',','))
                      ->setCellValue('N'. $y, number_format($row["s093"],6,'.',','))
                      ->setCellValue('O'. $y, number_format($row["s103"],6,'.',','))
                      ->setCellValue('P'. $y, number_format($row["s113"],6,'.',','))
                      ->setCellValue('Q'. $y, number_format($row["s123"],6,'.',','))
                      ->setCellValue('R'. $y, number_format($row["s133"],6,'.',','))
                      ->setCellValue('S'. $y, number_format($row["s143"],6,'.',','))
                      ->setCellValue('T'. $y, number_format($row["s153"],6,'.',','))
                      ->setCellValue('U'. $y, number_format($row["s163"],6,'.',','))
                      ->setCellValue('V'. $y, number_format($row["s173"],6,'.',','))
                      ->setCellValue('W'. $y, number_format($row["s183"],6,'.',','))
                      ->setCellValue('X'. $y, number_format($row["s193"],6,'.',','))
                      ->setCellValue('Y'. $y, number_format($row["s203"],6,'.',','))
                      ->setCellValue('Z'. $y, number_format($row["s213"],6,'.',','));

                $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                $osheet ->getStyle('B'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                $osheet ->getStyle('D'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                $osheet ->getStyle('E'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);

                if (($y%2)==0){
                    $osheet->getStyle('A'.$y.':Z'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
                }
            }
            $y++;
                /*
                  $osheet ->getStyle('A'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('A'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('A'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('A'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('B'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('B'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('B'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('B'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('C'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('C'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('C'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('C'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('D'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('D'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('D'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('D'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('E'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('E'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('E'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('E'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('F'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('F'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('F'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('F'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('G'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('G'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('G'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('G'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                */
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
        <input name="isdetails" type="checkbox" value="Y" <? if ($_GET['isdetails']=='Y') echo " checked"; ?> >含細項
        &nbsp;&nbsp;   &nbsp;&nbsp;
        <input type="submit" name="submit" id="submit" value="Export">  &nbsp;&nbsp;   &nbsp;&nbsp;
      </td></tr>
    </table>
  </div>
</form>
