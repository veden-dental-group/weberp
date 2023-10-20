<?php
  session_start();
  $pagetitle = "財務部 &raquo; 庫存帳齡分析";
  include("_data.php");
  include("_erp.php");
//  auth("erp_stockage.php");

  if (is_null($_GET['edate'])) {
    $edate =  date('Y-m-d');
  } else {
    $edate=$_GET['edate'];
  }

  if (is_null($_GET['pmonth'])) {
    $pmonth = date("Ym",strtotime("-1 month")) ;
  } else {
    $pmonth = $_GET['pmonth'] ;
  }

  $pyy=substr($pmonth,0,4);
  $pmm=substr($pmonth,4,2);

  $yy=substr($edate,0,4);
  $mm=substr($edate,5,2);
  $dd=substr($edate,8,2);
  $edate1=date('Y-m-d', strtotime("-180 days", mktime(0,0,0,$mm,$dd,$yy)));
  $edate2=date('Y-m-d', strtotime("-1 years", mktime(0,0,0,$mm,$dd,$yy)));
  $edate3=date('Y-m-d', strtotime("-2 years", mktime(0,0,0,$mm,$dd,$yy)));
  $edate4=date('Y-m-d', strtotime("-3 years", mktime(0,0,0,$mm,$dd,$yy)));
  $edate5=date('Y-m-d', strtotime("-4 years", mktime(0,0,0,$mm,$dd,$yy)));
  $edate6=date('Y-m-d', strtotime("-4 years", mktime(0,0,0,$mm,$dd,$yy)));

  if (($_GET["submit"]=="Excel") or  ($_GET["submit"]=="PDF")) {
        error_reporting(E_ALL);

        require_once 'classes/PHPExcel.php';
        require_once 'classes/PHPExcel/IOFactory.php';
        $objReader = PHPExcel_IOFactory::createReader('Excel5');
        $objPHPExcel = $objReader->load("templates/stockage.xls");
        //$objPHPExcel = new PHPExcel();
        // Set properties

        $objPHPExcel ->getProperties()->setCreator( 'Stock Age')
                     ->setLastModifiedBy('Stock Age')
                     ->setTitle('Stock Age')
                     ->setSubject('Stock Age')
                     ->setDescription('Stock Age')
                     ->setKeywords('Stock Age')
                     ->setCategory('Stock Age');


        $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
        $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
        $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
        $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
        $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
        //$objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣

        $objPHPExcel->setActiveSheetIndex(0);
        $osheet = $objPHPExcel->getActiveSheet();
        $y=1;

        //$osheet -> setCellValue('A1', "項次")
        //        -> setCellValue('B1', "品代")
        //        -> setCellValue('C1', "品名")
        //        -> setCellValue('D1', "1天~180天 (X <=" . $edate1 . ")")
        //        -> setCellValue('E1', "181天~一年 (" .$edate1 . "< X <=" . $edate2 . ")")
        //        -> setCellValue('F1', "一年~二年 (" .$edate2 . "< X <=" . $edate3 . ")")
        //        -> setCellValue('G1', "二年~三年 (" .$edate3 . "< X <=" . $edate4 . ")")
        //       -> setCellValue('H1', "三年~四年 (" .$edate4 . "< X <=" . $edate5 . ")")
        //        -> setCellValue('I1', "四年以上 (" .$edate6 . "< X)")
        //        -> setCellValue('J1', "庫存不足");

        $osheet -> setCellValue('C2', $edate);

        //$osheet -> setCellValue('A1', "項次")
        //        -> setCellValue('B1', "品代")
        //        -> setCellValue('C1', "品名")
        //        -> setCellValue('D1', "1天~180天")
        //        -> setCellValue('E1', "181天~一年")
        //        -> setCellValue('F1', "一年~二年")
        //        -> setCellValue('G1', "二年~三年")
        //        -> setCellValue('H1', "三年~四年")
        //        -> setCellValue('I1', "四年以上")
        //        -> setCellValue('J1', "單位");
        //        -> setCellValue('J1', "庫存不足");
        $i=0;
        $y=5;
        $query7="select * from stock where countdate='$edate' and ( qty0!=0 or qty1>0 or qty2>0 or qty3>0 or qty4>0 or qty5>0 or qty6>0) order by code " ;
        $result7= mysql_query($query7) or die ('1355 Stock select error!!' . mysql_error());
        while ($row7 = mysql_fetch_array($result7)) {
            //要取出單價
            $code=$row7['code'];
            $s1="select ccc23a from ccc_file where ccc01='$code' and ccc02='$pyy' and ccc03='$pmm'";
            $erp_sql1 = oci_parse($erp_conn1,$s1 );
            oci_execute($erp_sql1);
            $row1 = oci_fetch_array($erp_sql1, OCI_ASSOC);
            if (is_null($row1['CCC23A'])) {
                $price=0;
            } else {
                $price=floatval($row1['CCC23A']);
            }

            $i++;
            $osheet ->setCellValue('A'. $y, $i)
                    ->setCellValue('B'. $y, $code)
                    ->setCellValue('C'. $y, $row7['name'])
                    ->setCellValue('D'. $y, $row7['unit'])
                    ->setCellValue('E'. $y, $price)
                    ->setCellValue('F'. $y, number_format($row7["qty1"],2,'.',','))
                    ->setCellValue('G'. $y, "=E".$y."*"."F".$y)
                    ->setCellValue('H'. $y, number_format($row7["qty2"],2,'.',','))
                    ->setCellValue('I'. $y, "=E".$y."*"."H".$y)
                    ->setCellValue('J'. $y, number_format($row7["qty3"],2,'.',','))
                    ->setCellValue('K'. $y, "=E".$y."*"."J".$y)
                    ->setCellValue('L'. $y, number_format($row7["qty4"],2,'.',','))
                    ->setCellValue('M'. $y, "=E".$y."*"."L".$y)
                    ->setCellValue('N'. $y, number_format($row7["qty5"],2,'.',','))
                    ->setCellValue('O'. $y, "=E".$y."*"."N".$y)
                    ->setCellValue('P'. $y, number_format($row7["qty6"],2,'.',','))
                    ->setCellValue('Q'. $y, "=E".$y."*"."P".$y)
                    ->setCellValue('R'. $y, "=F".$y."+"."H".$y."+"."J".$y ."+"."L".$y ."+"."N".$y ."+"."P".$y)
                    ->setCellValue('S'. $y, "=G".$y."+"."I".$y."+"."K".$y ."+"."M".$y ."+"."O".$y ."+"."Q".$y) ;
            $osheet ->getStyle('B'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            $osheet ->getStyle('C'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            $osheet ->getStyle('D'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $osheet ->getStyle('E'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $osheet ->getStyle('F'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $osheet ->getStyle('G'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $osheet ->getStyle('H'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $osheet ->getStyle('I'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $osheet ->getStyle('J'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $osheet ->getStyle('K'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $osheet ->getStyle('L'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $osheet ->getStyle('M'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $osheet ->getStyle('N'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $osheet ->getStyle('O'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $osheet ->getStyle('P'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $osheet ->getStyle('Q'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $osheet ->getStyle('R'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $osheet ->getStyle('S'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            if (($y%2)==0){
                $osheet->getStyle('A'.$y.':S'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
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
        $osheet ->getStyle('M'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
        $osheet ->getStyle('N'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
        $osheet ->getStyle('O'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
        $osheet ->getStyle('P'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
        $osheet ->getStyle('Q'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
        $osheet ->getStyle('R'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
        $osheet ->getStyle('S'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

        //合計
        $osheet ->setCellValue('C'. $y, '合計')
                    ->setCellValue('G'. $y, '=sum(G5:G' . ($y-1).')')
                    ->setCellValue('I'. $y, '=sum(I5:I' . ($y-1).')')
                    ->setCellValue('K'. $y, '=sum(K5:K' . ($y-1).')')
                    ->setCellValue('M'. $y, '=sum(M5:M' . ($y-1).')')
                    ->setCellValue('O'. $y, '=sum(O5:O' . ($y-1).')')
                    ->setCellValue('Q'. $y, '=sum(Q5:Q' . ($y-1).')')
                    ->setCellValue('S'. $y, '=sum(S5:S' . ($y-1).')');

        $osheet ->getStyle('G'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
        $osheet ->getStyle('I'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
        $osheet ->getStyle('K'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
        $osheet ->getStyle('M'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
        $osheet ->getStyle('O'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
        $osheet ->getStyle('Q'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
        $osheet ->getStyle('S'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
        // Rename sheet
        $osheet ->setTitle('Stock Age');
        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $objPHPExcel->setActiveSheetIndex(0);


        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="'. 'Stockage_'. $edate . '.xls"');
        header('Cache-Control: max-age=0');
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save('php://output');
        exit;
  }

  $IsAjax = false;
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>庫齡分析 </p>
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
         截止日期:
        <input name="edate" type="text" id="edate" size="12" maxlength="12" value=<?=$edate;?> onfocus="new WdatePicker()"> &nbsp;&nbsp;
        單價月份:
        <input name="pmonth" type="text" id="pmonth" onfocus="WdatePicker({dateFmt:'yyyyMM'})" value="<?=$pmonth;?>"> &nbsp;&nbsp;
        &nbsp;&nbsp;
        <input type="submit" name="submit" id="submit" value="計算">  &nbsp;&nbsp;   &nbsp;&nbsp;
        <input type="submit" name="submit" id="submit" value="Excel">  &nbsp;&nbsp;   &nbsp;&nbsp;
      </td></tr>
    </table>
  </div>
</form>
<? if (is_null($_GET['submit'])) die ; ?>
<? if ($_GET['submit']!="計算") die ; ?>
<?
//重算庫齡 原則上先重算庫齡後 才可以導出
//本重算會先將同一天的庫齡資料刪除後 再重算
//1 將sfe 入庫的資料寫到stockin中
//2 將sfe出庫的資料寫到stockout中
//3 將期間寫到stock中
//4 將stockin 加到 stock
//5 將stockout 減到 stock中
//6 顯示資料

$sdate1= "1天~180天 (X <= $edate1 ) ";
$sdate2= "181天~一年 ( $edate1 < X <= $edate2 )";
$sdate3= "一年~二年 ( $edate2 < X <= $edate3 )";
$sdate4= "二年~三年 ( $edate3 < X <= $edate4 )";
$sdate5= "三年~四年 ( $edate4 < X <= $edate5 )";
$sdate6= "四年以上  ( X < $edate6 )";


?>
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
        <th width="16">&nbsp;</th>
        <th>品代</th>
        <th>品名</th>
        <th> <?=$sdate1;?> </th>
        <th> <?=$sdate2;?> </th>
        <th> <?=$sdate3;?> </th>
        <th> <?=$sdate4;?> </th>
        <th> <?=$sdate5;?> </th>
        <th> <?=$sdate6;?> </th>
        <th>庫存不足</th>
        <th>單位</th>
    </tr>
    <?
      //取出$edate 之前的所有入庫資料: 入庫rvu/rvv, 退料sfe, 雜收ina/inb, 還料imr/imq 寫入stockin
      //先刪除同一個統計日
      $query2= "delete from stockin where countdate='$edate'";
      $result2= mysql_query($query2) or die ('1154 Stockin Deleted error. ' .mysql_error());
      //入庫 1
      $s1="select to_char(rvu03,'yyyy-mm-dd') rvu03, rvv31, ima02, rvvud02, sum(rvvud07) rvvud07 from rvu_file, rvv_file, ima_file " .
          "where rvuconf='Y' and rvu00='1' and rvu03<=to_date('$edate','yy/mm/dd') and rvu01=rvv01 and rvv31=ima01 and substr(ima06,1,1)!='9' group by rvu03,rvv31,ima02,rvvud02 ";
      $erp_sql1 = oci_parse($erp_conn1,$s1 );
      oci_execute($erp_sql1);
      while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
          $rvu03=$row1['RVU03'];   //日期
          $rvv31=$row1['RVV31'];   //料號
          $ima02=$row1['IMA02'];   //品名
          $rvvud02=$row1['RVVUD02'];   //入庫單位
          $rvvud07=is_null($row1['RVVUD07'])?0:$row1['RVVUD07'];   //入庫量
          $query1= "insert into stockin ( countdate, code, name, unit, indate, qty, source ) values (
                   '" . $edate      . "',
                   '" . $rvv31      . "',
                   '" . $ima02      . "',
                   '" . $rvvud02    . "',
                   '" . $rvu03      . "',
                   '" . $rvvud07    . "','1')";
          $result1= mysql_query($query1) or die ('1163 Stockin Added error. ' .mysql_error());
      }

      //退料 2
      $s2="select to_char(sfe04,'yyyy-mm-dd') sfe04, sfe07, ima02, sfe17, sum(sfe16) sfe16 from sfe_file, ima_file " .
          "where sfe06='4' and sfe04<=to_date('$edate','yy/mm/dd') and sfe07=ima01 and substr(ima06,1,1)!='9' group by sfe04,sfe07,ima02,sfe17 ";
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
          $sfe04=$row2['SFE04'];   //日期
          $sfe07=$row2['SFE07'];   //料號
          $ima02=$row2['IMA02'];   //品名
          $sfe17=$row2['SFE17'];   //入庫單位
          $sfe16=$row2['SFE16'];   //入庫量
          $query2= "insert into stockin ( countdate, code, name, unit, indate, qty, source ) values (
                   '" . $edate      . "',
                   '" . $sfe07      . "',
                   '" . $ima02      . "',
                   '" . $sfe17      . "',
                   '" . $sfe04      . "',
                   '" . $sfe16      . "','2')";
          $result2= mysql_query($query2) or die ("1180 Stockin Added error. " .mysql_error());
      }

      //雜收 3
      $s2="select to_char(ina02,'yyyy-mm-dd') ina02, inb04, ima02, inb08, sum(inb09*inb08_fac) inb09 from ina_file, inb_file, ima_file " .
          "where inaconf='Y' and ina00='3' and ina02<=to_date('$edate','yy/mm/dd') and ina01=inb01 and inb04=ima01 and substr(ima06,1,1)!='9' group by ina02,inb04,ima02,inb08 ";
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
          $ina02=$row2['INA02'];   //日期
          $inb04=$row2['INB04'];   //料號
          $ima02=$row2['IMA02'];   //單名
          $inb08=$row2['INB08'];   //入庫單位
          $inb09=$row2['INB09'];   //入庫量
          $query2= "insert into stockin ( countdate, code, name, unit, indate, qty, source ) values (
                   '" . $edate      . "',
                   '" . $inb04      . "',
                   '" . $ima02      . "',
                   '" . $inb08      . "',
                   '" . $ina02      . "',
                   '" . $inb09      . "','3')";
          $result2= mysql_query($query2) or die ('1180 Stockin Added error. ' .mysql_error());
      }

      //借料 4
      $s2="select to_char(imo02,'yyyy-mm-dd') imo02, imp03, ima02, imp05, sum(imp04) imp04 from imo_file, imp_file, ima_file " .
          "where imoconf='Y' and imo02<=to_date('$edate','yy/mm/dd') and imo01=imp01 and imp03=ima01 and substr(ima06,1,1)!='9' group by imo02,imp03,ima02,imp05 ";
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
          $imo02=$row2['IMO02'];   //日期
          $imp03=$row2['IMP03'];   //料號
          $ima02=$row2['IMA02'];   //品名
          $imp05=$row2['IMP05'];   //入庫單位
          $imp04=$row2['IMP04'];   //入庫量
          $query2= "insert into stockin ( countdate, code, name, unit, indate, qty, source ) values (
                   '" . $edate      . "',
                   '" . $imp03      . "',
                   '" . $ima02      . "',
                   '" . $imp05      . "',
                   '" . $imo02      . "',
                   '" . $imp04      . "','4')";
          $result2= mysql_query($query2) or die ('1210 Stockin Added error. ' .mysql_error());
      }



      //先刪除 stockout 同一個統計日
      $query2= "delete from stockout where countdate='$edate'";
      $result2= mysql_query($query2) or die ('1180 Stockout Deleted error. ' .mysql_error());

      //庫存減少的計算中不用理會庫存單位 因為一定一樣的
      //發料 2
      $s2="select to_char(sfe04,'yyyy-mm-dd') rvu03, sfe07 rvv31, sum(sfe16) rvv17 from sfe_file where sfe06='1' and sfe04<=to_date('$edate','yy/mm/dd') group by sfe04, sfe07";
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
          $rvu03=$row2['RVU03'];   //日期
          $rvv31=$row2['RVV31'];   //料號
          $rvv17=$row2['RVV17'];   //入庫量
          $query2= "insert into stockout ( countdate, code, outdate, qty, source ) values (
                   '" . $edate      . "',
                   '" . $rvv31      . "',
                   '" . $rvu03      . "',
                   '" . $rvv17      . "','2')";
          $result2= mysql_query($query2) or die ('1230 Stockout Added error. ' .mysql_error());
      }

      //雜發 3
      $s2="select to_char(ina02,'yyyy-mm-dd') rvv03, inb04 rvv31, sum(inb09*inb08_fac) rvv17 from ina_file, inb_file where inaconf='Y' and ina00='1' and ina02<=to_date('$edate','yy/mm/dd') and ina01=inb01 group by ina02,inb04 ";
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
          $rvu03=$row2['RVV03'];   //日期
          $rvv31=$row2['RVV31'];   //料號
          $rvv17=$row2['RVV17'];   //入庫量
          $query2= "insert into stockout ( countdate, code, outdate, qty, source ) values (
                   '" . $edate      . "',
                   '" . $rvv31      . "',
                   '" . $rvu03      . "',
                   '" . $rvv17      . "','3')";
          $result2= mysql_query($query2) or die ('1245 Stockout Added error. ' .mysql_error());
      }

      //還料 4
      $s2="select to_char(imr02,'yyyy-mm-dd') rvv03, imq05 rvv31, sum(imq07) rvv17 from imr_file, imq_file where imrconf='Y' and imr02<=to_date('$edate','yy/mm/dd') and imr01=imq01 group by imr02, imq05 ";
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
          $rvu03=$row2['RVV03'];   //日期
          $rvv31=$row2['RVV31'];   //料號
          $rvv17=$row2['RVV17'];   //入庫量
          $query2= "insert into stockout ( countdate, code, outdate, qty, source ) values (
                   '" . $edate      . "',
                   '" . $rvv31      . "',
                   '" . $rvu03      . "',
                   '" . $rvv17      . "','4')";
          $result2= mysql_query($query2) or die ('1260 Stockout Added error. ' .mysql_error());
      }
      //無訂單出貨
      $s2="select to_char(oga02,'yyyy-mm-dd') rvv03, ogb04 rvv31, sum(ogb12) rvv17 from oga_file, ogb_file " .
          "where oga905 is null and ogaconf='Y' and oga02<=to_date('$edate','yy/mm/dd') and oga01=ogb01 group by oga02, ogb04 ";
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
          $rvu03=$row2['RVV03'];   //日期
          $rvv31=$row2['RVV31'];   //料號
          $rvv17=$row2['RVV17'];   //入庫量
          $query2= "insert into stockout ( countdate, code, outdate, qty, source ) values (
                   '" . $edate      . "',
                   '" . $rvv31      . "',
                   '" . $rvu03      . "',
                   '" . $rvv17      . "','4')";
          $result2= mysql_query($query2) or die ('1260 Stockout Added error. ' .mysql_error());
      }

      //先刪除stockin 同一個統計日
      $query2= "delete from stock where countdate='$edate'";
      $result2= mysql_query($query2) or die ('1180 Stockout Deleted error. ' .mysql_error());

      //這裡應先把目前所有的料號的庫存單位寫到stock中以便換算
      $sql3="select ima01, ima02, ima25 from ima_file where substr(ima06,1,1)>'9' order by ima01 ";
      $erp_sql3 = oci_parse($erp_conn1,$sql3 );
      oci_execute($erp_sql3);
      while ($row3 = oci_fetch_array($erp_sql3, OCI_ASSOC)) {
          $code=$row3['IMA01'];
          $name=$row3['IMA02'];
          $unit=$row3['IMA25'];
          $query2= "insert into stock ( countdate, code, name, unit, date1, date2, date3, date4, date5, date6, qty0, qty1, qty2, qty3, qty4, qty5,qty6 ) values (
                   '" . $edate       . "',
                   '" . $code        . "',
                   '" . $name        . "',
                   '" . $unit        . "',
                   '" . $edate1      . "',
                   '" . $edate2      . "',
                   '" . $edate3      . "',
                   '" . $edate4      . "',
                   '" . $edate5      . "',
                   '" . $edate6      . "',0,0,0,0,0,0,0)";
          $result2= mysql_query($query2) or die ('1286 Stock added error. ' .mysql_error());
      }


      $query1="select code, name, unit, indate, sum(qty) qty from stockin where countdate='$edate' group by code, name, unit, indate order by code,indate ";
      $result1 = mysql_query($query1) or die ('1269 Stockin error!!' . mysql_error());
      while ($row1= mysql_fetch_array($result1)) {
          $code=$row1['code'];
          $indate=$row1['indate'];
          $qty=$row1['qty'];
          $name=$row1['name'];
          $inunit=$row1['unit'];
          //先檢查本料號有記錄否
          $query2="select * from stock where countdate='$edate' and code='$code' ";
          $result2= mysql_query($query2) or die ('1277 Stock select error!!' . mysql_error());
          $row2= mysql_fetch_array($result2);
          $stockunit=$row2['unit'];

          //庫存單位和入庫單位不同 要做單位換算
          if ( $inunit != $stockunit) {
                        // 來源單位 目的單位 來源數量 目的數量
            $sql5 = "select smd02,    smd03,   smd04, smd06 from smd_file where smd01='$code' ";
            $erp_sql5 = oci_parse($erp_conn1,$sql5 );
            oci_execute($erp_sql5);
            $row5 = oci_fetch_array($erp_sql5, OCI_ASSOC);
            if (is_null($row5['SMD02'])) {
              //do nothing
            } else {
              if ($inunit==$row5['SMD02']) {
                  $qty = ($qty * $row5['SMD06'] / $row5['SMD04'] );
              } else {
                  $qty = ($qty * $row5['SMD04'] / $row5['SMD06']);
              }
            }
          }

          /*
          //先檢查本料號有記錄否
          $query2="select * from stock where countdate='$edate' and code='$code' ";
          $result2= mysql_query($query2) or die ('1277 Stock select error!!' . mysql_error());
          if (mysql_num_rows($result2) == 0) { //新品代 加一筆
              $query2= "insert into stock ( countdate, code, name, unit, date1, date2, date3, date4, date5, date6, qty0, qty1, qty2, qty3, qty4, qty5,qty6 ) values (
                       '" . $edate       . "',
                       '" . $code        . "',
                       '" . $name        . "',
                       '" . $unit        . "',
                       '" . $edate1      . "',
                       '" . $edate2      . "',
                       '" . $edate3      . "',
                       '" . $edate4      . "',
                       '" . $edate5      . "',
                       '" . $edate6      . "',0,0,0,0,0,0,0)";
              $result2= mysql_query($query2) or die ('1286 Stock added error. ' .mysql_error());
              commit;
          }
          */


          if ($indate < $edate6) {
              $query3="update stock set qty6 = qty6 + $qty where countdate='$edate' and code='$code'  " ;
          } else if (($indate >= $edate5) && ($indate < $edate4)) {
              $query3="update stock set qty5 = qty5 + $qty where countdate='$edate' and code='$code'  " ;
          } else if (($indate >= $edate4) && ($indate < $edate3)) {
              $query3="update stock set qty4 = qty4 + $qty where countdate='$edate' and code='$code'  " ;
          } else if (($indate >= $edate3) && ($indate < $edate2)) {
              $query3="update stock set qty3 = qty3 + $qty where countdate='$edate' and code='$code'  " ;
          } else if (($indate >= $edate2) && ($indate < $edate1)) {
              $query3="update stock set qty2 = qty2 + $qty where countdate='$edate' and code='$code'  " ;
          } else {
              $query3="update stock set qty1 = qty1 + $qty where countdate='$edate' and code='$code'  " ;
          }
          $result3= mysql_query($query3) or die ('1304 Stock Updated error. ==>' .$sql5 .'<==>' . $query3  ."<==" . mysql_error());
          commit;
      }

      //自stockout中取出物料的合計出庫數
      $query4="select code, sum(qty) qty from stockout where countdate='$edate' group by code order by code ";
      $result4 = mysql_query($query4) or die ('1310 Stockout error!!' . mysql_error());
      while ($row4= mysql_fetch_array($result4)) {
          $code=$row4['code'];
          $qty=floatval($row4['qty']);

          $query5="select * from stock where countdate='$edate' and code='$code' limit 1";
          $result5= mysql_query($query5) or die ('1316 Stock error!!' . mysql_error());
          $row5=mysql_fetch_array($result5);
          $pkey=$row5['pkey'];
          $qty1=floatval($row5['qty1']);
          $qty2=floatval($row5['qty2']);
          $qty3=floatval($row5['qty3']);
          $qty4=floatval($row5['qty4']);
          $qty5=floatval($row5['qty5']);
          $qty6=floatval($row5['qty6']);

          if ($qty>=$qty6){
              $qty -= $qty6;
              $qty6=0;
          } else {
              $qty6-=$qty;
              $qty=0;
          }

          if ($qty>=$qty5){
              $qty -= $qty5;
              $qty5=0;
          } else {
              $qty5-=$qty;
              $qty=0;
          }

          if ($qty>=$qty4){
              $qty -= $qty4;
              $qty4=0;
          } else {
              $qty4-=$qty;
              $qty=0;
          }

          if ($qty>=$qty3){
              $qty -= $qty3;
              $qty3=0;
          } else {
              $qty3-=$qty;
              $qty=0;
          }

          if ($qty>=$qty2){
              $qty -= $qty2;
              $qty2=0;
          } else {
              $qty2-=$qty;
              $qty=0;
          }

          if ($qty>=$qty1){
              $qty -= $qty1;
              $qty1=0;
          } else {
              $qty1-=$qty;
              $qty=0;
          }

          $query6="update stock set qty0=$qty, qty1=$qty1,qty2=$qty2,qty3=$qty3,qty4=$qty4,qty5=$qty5,qty6=$qty6 where pkey='$pkey'";
          $result6= mysql_query($query6) or die ('1350 Stock update error!!' . mysql_error());
      }

      $bgkleur = "ffffff";
      $i=0;
      $query7="select * from stock where countdate='$edate' and ( qty0!=0 or qty1>0 or qty2>0 or qty3>0 or qty4>0 or qty5>0 or qty6>0) order by code " ;
      $result7= mysql_query($query7) or die ('1355 Stock select error!!' . mysql_error());
      while ($row7 = mysql_fetch_array($result7)) {
          $i++;
    ?>
          <tr bgcolor="#<?=$bgkleur;?>">
              <td><?=$i;?></td>
              <td><?=$row7["code"];?></td>
              <td><?=$row7["name"];?></td>
              <td style="text-align:right" ><?=number_format($row7["qty1"],2,'.',',');?></td>
              <td style="text-align:right" ><?=number_format($row7["qty2"],2,'.',',');?></td>
              <td style="text-align:right" ><?=number_format($row7["qty3"],2,'.',',');?></td>
              <td style="text-align:right" ><?=number_format($row7["qty4"],2,'.',',');?></td>
              <td style="text-align:right" ><?=number_format($row7["qty5"],2,'.',',');?></td>
              <td style="text-align:right" ><?=number_format($row7["qty6"],2,'.',',');?></td>
              <td style="text-align:right" ><?=number_format($row7["qty0"],2,'.',',');?></td>
              <td><?=$row7["unit"];?></td>
          </tr>
    <?
      }
    ?>
</table>
