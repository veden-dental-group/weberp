<?php
  session_start();
  $pagetitle = "資材部 &raquo; 採購交貨率";
  include("_data.php");
  include("_erp.php");
  //auth("erp_purchaserate.php");

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

  if (is_null($_GET['vd'])) {
    $vd='vd110';
  } else {
    $vd=$_GET['vd'];
  }

    if ($_GET["submit"]=="匯出") {
        error_reporting(E_ALL);
        require_once 'classes/PHPExcel.php';
        require_once 'classes/PHPExcel/IOFactory.php';
        $objPHPExcel = new PHPExcel();
        // Set properties
        $objPHPExcel ->getProperties()->setCreator( 'Frank')
                     ->setLastModifiedBy('Frank')
                     ->setTitle('採購交貨率')
                     ->setSubject('採購交貨率')
                     ->setDescription('採購交貨率')
                     ->setKeywords('採購交貨率')
                     ->setCategory('採購交貨率');

        // Add some data

        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A1', '採購交貨率');
        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A3', '廠商代號')
                    ->setCellValue('B3', '廠商名稱')
                    ->setCellValue('C3', '採購單號')
                    ->setCellValue('D3', '採購項次')
                    ->setCellValue('E3', '物料編號')
                    ->setCellValue('F3', '物料名稱')
                    ->setCellValue('G3', '採購量')
                    ->setCellValue('H3', '採購單位')
                    ->setCellValue('I3', '採購日期')
                    ->setCellValue('J3', '原始交貨日期')
                    ->setCellValue('K3', '入庫日期')
                    ->setCellValue('L3', '入庫數量')
                    ->setCellValue('M3', '入庫單位');

        $y=4;
        $i=0;

        $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);
        $edate1=substr($edate,2,2).substr($edate,5,2).substr($edate,8,2);
        $s2= "select pmm09,  pmc03, pmn01, pmn02, pmn04, pmn041, pmn20, pmn07,
              to_char(pmm04,'yy-mm-dd') pmm04, to_char(pmn33,'yy-mm-dd') pmn33, to_char(rvu03,'yy-mm-dd') rvu03,rvvud07, rvvud02
              from pmm_file, pmc_file, pmn_file, rvu_file, rvv_file where pmm09=pmc01 and pmm01=pmn01 and rvu01=rvv01 and rvv36=pmn01 and rvv37=pmn02
              and rvu03 between to_date($bdate1,'yy/mm/dd') and to_date($edate1,'yy/mm/dd')
              and pmm09!='V110000'
              order by pmm09, pmn01,pmn02";
        if ($vd=='vd110') {
          $erp_sql2 = oci_parse($erp_conn1,$s2 );
        } else {
          $erp_sql2 = oci_parse($erp_conn2,$s2 );
        }
        oci_execute($erp_sql2);
        while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
            $i++;
            $objPHPExcel->setActiveSheetIndex(0)
                        ->setCellValue('A'. $y, $row2["PMM09"])
                        ->setCellValue('B'. $y, $row2["PMC03"])
                        ->setCellValue('C'. $y, $row2["PMN01"])
                        ->setCellValue('D'. $y, $row2["PMN02"])
                        ->setCellValue('E'. $y, $row2["PMN04"])
                        ->setCellValue('F'. $y, $row2["PMN041"])
                        ->setCellValue('G'. $y, $row2["PMN20"])
                        ->setCellValue('H'. $y, $row2["PMN07"])
                        ->setCellValue('I'. $y, $row2["PMM04"])
                        ->setCellValue('J'. $y, $row2["PMN33"])
                        ->setCellValue('K'. $y, $row2["RVU03"])
                        ->setCellValue('L'. $y, $row2["RVVUD07"])
                        ->setCellValue('M'. $y, $row2["RVVUD02"]);
            $y++;
        }

        // Rename sheet
        $objPHPExcel->getActiveSheet()->setTitle('採購交貨率');

        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $objPHPExcel->setActiveSheetIndex(0);

        // Redirect output to a client’s web browser (Excel5)
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="PurchaseRate.xls"');
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
  $IsAjax = True;
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>客戶Delay Report </p>
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
         起訖日期:
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> ~~
        <input name="edate" type="text" id="edate" size="12" maxlength="12" value=<?=$edate;?> onfocus="new WdatePicker()">
        公司別:
        <select name="vd" id="vd">
          <option value='vd110' <? if($vd=='vd110') echo ' selected'?> >VD110</option>
          <option value='vd210' <? if($vd=='vd210') echo ' selected'?> >VD210</option>
        </select>
        <input type="submit" name="submit" id="submit" value="查詢">  &nbsp;&nbsp;   &nbsp;&nbsp;
        <input type="submit" name="submit" id="submit" value="匯出">  &nbsp;&nbsp;   &nbsp;&nbsp;
      </td></tr>
    </table>
  </div>
</form>
<? if (is_null($_GET['submit'])) die ; ?>
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
        <th width="16">&nbsp;</th>
        <th>廠商代號</th>
        <th>廠商名稱</th>
        <th>採購單號</th>
        <th>採購項次</th>
        <th>物料編號</th>
        <th>物料名稱</th>
        <th>採購量</th>
        <th>採購單位</th>
        <th>採購日期</th>
        <th>原始交貨日期</th>
        <th>入庫日期</th>
        <th>入庫量</th>
        <th>入庫單位</th>
    </tr>
    <?
      $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);
      $edate1=substr($edate,2,2).substr($edate,5,2).substr($edate,8,2);
      //檢查工單號有無出貨  , 配件不用
      $s2= "select pmm09,  pmc03, pmn01, pmn02, pmn04, pmn041, pmn20, pmn07,
            to_char(pmm04,'yy-mm-dd') pmm04, to_char(pmn33,'yy-mm-dd') pmn33, to_char(rvu03,'yy-mm-dd') rvu03,rvvud07, rvvud02
            from pmm_file, pmc_file, pmn_file, rvu_file, rvv_file where pmm09=pmc01 and pmm01=pmn01 and rvu01=rvv01 and rvv36=pmn01 and rvv37=pmn02
            and rvu03 between to_date($bdate1,'yy/mm/dd') and to_date($edate1,'yy/mm/dd')
            and pmm09!='V110000'
            order by pmm09, pmn01,pmn02";
      if ($vd=='vd110') {
        $erp_sql2 = oci_parse($erp_conn1,$s2 );
      } else {
        $erp_sql2 = oci_parse($erp_conn2,$s2 );
      }
      oci_execute($erp_sql2);
      $bgkleur = "ffffff";
      $i=0;
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
          $i++;
          ?>
            <tr bgcolor="#<?=$bgkleur;?>">
                <td><?=$i;?></td>
                <td><?=$row2["PMM09"];?></td>
                <td><?=$row2["PMC03"];?></td>
                <td><?=$row2["PMN01"];?></td>
                <td><?=$row2["PMN02"];?></td>
                <td><?=$row2["PMN04"];?></td>
                <td><?=$row2["PMN041"];?></td>
                <td><div align=right><?=number_format($row2["PMN20"], 2, ".", ",");?></div></td>
                <td><?=$row2["PMN07"];?></td>
                <td><?=$row2["PMM04"];?></td>
                <td><?=$row2["PMN33"];?></td>
                <td><?=$row2["RVU03"];?></td>
                <td><div align=right><?=number_format($row2["RVVUD07"], 2, ".", ",");?></div></td>
                <td><?=$row2["RVVUD02"];?></td>
            </tr>
          <?
      }
      ?>
</table>
