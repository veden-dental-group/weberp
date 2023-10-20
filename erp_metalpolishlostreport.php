<?php
  session_start();
  $pagetitle = "資材部 &raquo; 金屬鑄造及研磨后重量一覽表";
  include("_data.php");
  include("_erp.php");
  //auth("erp_metalpolishlostreport.php");

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

    if ($_GET["submit"]=="匯出") {
        error_reporting(E_ALL);
        require_once 'classes/PHPExcel.php';
        require_once 'classes/PHPExcel/IOFactory.php';
        $objPHPExcel = new PHPExcel();
        // Set properties
        $objPHPExcel ->getProperties()->setCreator( 'Frank')
                     ->setLastModifiedBy('Frank')
                     ->setTitle('金屬鑄造及研磨后重量一覽表')
                     ->setSubject('金屬鑄造及研磨后重量一覽表')
                     ->setDescription('金屬鑄造及研磨后重量一覽表')
                     ->setKeywords('金屬鑄造及研磨后重量一覽表')
                     ->setCategory('金屬鑄造及研磨后重量一覽表');

        // Add some data

        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A1', '金屬鑄造及研磨后重量一覽表');
        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A3', '序號')
                    ->setCellValue('B3', '物料代號')
                    ->setCellValue('C3', '物料名稱')
                    ->setCellValue('D3', '鑄造后重量')
                    ->setCellValue('E3', '研磨后重量')
                    ->setCellValue('F3', '返工退料重量');

        $y=4;
        $i=0;

        $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);
        $edate1=substr($edate,2,2).substr($edate,5,2).substr($edate,8,2);
      $s2= "select sfe07, ima02, sum(decode(sfe06,'1',sfe16,0)) sfe16, sum(decode(sfe06,'4',sfe16,0)) reject, sum(ta_sfe002) ta_sfe002 from sfe_file, ima_file
            where sfe04>=to_date('$bdate1','yy/mm/dd') and sfe04<=to_date('$edate1','yy/mm/dd')
            and sfe07 like '1K%' and sfe07=ima01 group by sfe07, ima02 order by sfe07 ";
        $erp_sql2 = oci_parse($erp_conn,$s2 );
        oci_execute($erp_sql2);
        while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
            $i++;
            $objPHPExcel->setActiveSheetIndex(0)
                        ->setCellValue('A'. $y, $i)
                        ->setCellValue('B'. $y, $row2["SFE07"])
                        ->setCellValue('C'. $y, $row2["IMA02"])
                        ->setCellValue('D'. $y, $row2["SFE16"])
                        ->setCellValue('E'. $y, $row2["TA_SFE002"])
                        ->setCellValue('F'. $y, $row2["REJECT"]);
            $y++;
        }

        // Rename sheet
        $objPHPExcel->getActiveSheet()->setTitle('金屬鑄造及研磨后重量一覽表');

        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $objPHPExcel->setActiveSheetIndex(0);

        // Redirect output to a client’s web browser (Excel5)
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="MetalPolishLostReport.xls"');
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
<p>金屬鑄造及研磨后重量一覽表</p>
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
         起訖日期:
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> ~~
        <input name="edate" type="text" id="edate" size="12" maxlength="12" value=<?=$edate;?> onfocus="new WdatePicker()">
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
        <th>物料代碼</th>
        <th>物料名稱</th>
        <th>鑄造后重量</th>
        <th>研磨后重量</th>
        <th>返工退牙重量</th>
    </tr>
    <?
      $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);
      $edate1=substr($edate,2,2).substr($edate,5,2).substr($edate,8,2);
      //只取發料量
      $s2= "select sfe07, ima02, sum(decode(sfe06,'1',sfe16,0)) sfe16, sum(decode(sfe06,'4',sfe16,0)) reject, sum(ta_sfe002) ta_sfe002 from sfe_file, ima_file
            where sfe04>=to_date('$bdate1','yy/mm/dd') and sfe04<=to_date('$edate1','yy/mm/dd')
            and sfe07 like '1K%' and sfe07=ima01 group by sfe07, ima02 order by sfe07 ";
      $erp_sql2 = oci_parse($erp_conn,$s2 );
      oci_execute($erp_sql2);
      $bgkleur = "ffffff";
      $i=0;
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
          $i++;
          ?>
            <tr bgcolor="#<?=$bgkleur;?>">
                <td><?=$i;?></td>
                <td><?=$row2["SFE07"];?></td>
                <td><?=$row2["IMA02"];?></td>
                <td><div align=right><?=number_format($row2["SFE16"], 2, ".", ",");?></div></td>
                <td><div align=right><?=number_format($row2["TA_SFE002"], 2, ".", ",");?></div></td>
                <td><div align=right><?=number_format($row2["REJECT"], 2, ".", ",");?></div></td>
            </tr>
          <?
      }
      ?>
</table>
