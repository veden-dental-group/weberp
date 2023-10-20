<?php
  session_start();
  $pagetitle = "資材部 &raquo; 列印盤點清冊";
  include("_data.php");
  include("_erp.php");
//  auth("erp_inventorycheck.php");  
                                               
  
  if (is_null($_GET['fca01'])) {
    $fca01 =  '';
  } else {
    $fca01=$_GET['fca01'];
  }  
  
  //$fca01='201512002';
  if (($_GET["submit"]=="Excel") or  ($_GET["submit"]=="PDF")) {  
        error_reporting(E_ALL);  
        
        require_once 'classes/PHPExcel.php'; 
        require_once 'classes/PHPExcel/IOFactory.php';  
        //$objReader = PHPExcel_IOFactory::createReader('Excel5'); 
        //$objPHPExcel = $objReader->load("$filename");  
        $objPHPExcel = new PHPExcel();  
        // Set properties
                                                   
        $objPHPExcel ->getProperties()->setCreator( 'Inventory Check')
                     ->setLastModifiedBy('Inventory Check')
                     ->setTitle('Inventory Check')
                     ->setSubject('Inventory Check')
                     ->setDescription('Inventory Check')
                     ->setKeywords('Inventory Check')
                     ->setCategory('Inventory Check');     
        
        $s2="select fca05,gem02, fca07,faf02, fca06,gen02, fajud02, fca02,fca03,fca031,fca04, faj04,fab02, faj06,fca09,fca08 " .  
            "from (select fca02,fca03,fca031,fca04,fca05,fca06,fca07,fca08,fca09,faj04,faj06,fajud02 from fca_file,faj_file where fca01='$fca01' and fca04=faj01 ) " .
            "left join gem_file on fca05=gem01 " .
            "left join faf_file on fca07=faf01 " .
            "left join gen_file on fca06=gen01 " .
            "left join fab_file on faj04=fab01 " .  
            "order by fca05,fca07 " ;                  
        $y=1;
        $oldfca05='ZZZZZZZZZZ';
        $erp_sql2 = oci_parse($erp_conn1,$s2 );
        oci_execute($erp_sql2);   
        while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
          if ($row2['FCA05']!=$oldfca05) {
              if ($oldfca05!='ZZZZZZZZZZ') { 
                  
                  $osheet->getStyle('A3')->getBorders()->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet->getStyle('A3')->getBorders()->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
                  $osheet->getStyle('A3')->getBorders()->getLeft()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
                  $osheet->getStyle('A3')->getBorders()->getRight()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
                  
                  $objStyleA3 = $osheet->getStyle('A3');
                  $osheet ->duplicateStyle($objStyleA3, 'A3:M'.($y-1));
              }
              
              if (is_null($row2['GEM02'])) {
                  $gem02='noname';
              } else {
                  $gem02=$row2['GEM02'];
              }
              if($gem02=='3DCAD/CAM') $gem02='3DCAD_CAM';
              $osheet= $objPHPExcel->createSheet();
              $osheet->setTitle($gem02);
              //$objPHPExcel->setActiveSheetIndex(1);             
              //$osheet = $objPHPExcel->getActiveSheet(); 
              $osheet->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
              $osheet->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);             
              $osheet->getDefaultStyle()->getFont()->setName('Sylfaen');
              $osheet->getDefaultStyle()->getFont()->setSize(8);
              $osheet->getDefaultRowDimension()->setRowHeight(15); 
              //$osheet->getPageSetup()->setHorizontalCentered(true);
              //$osheet->getPageSetup()->setVerticalCentered(true);
              $osheet->getPageMargins()->setTop(0);
              $osheet->getPageMargins()->setRight(0);
              $osheet->getPageMargins()->setLeft(0);
              $osheet->getPageMargins()->setBottom(0);
              $osheet->getPageMargins()->setHeader(0);    
              $osheet->getPageMargins()->setFooter(0);    
              $osheet->mergeCells('A1:M1');
              $osheet ->setCellValue('A1', '保管部門: ' . $row2["FCA05"].' '.$row2['GEM02']); 
              $osheet ->setCellValue('A3', '財產編號')    
                      ->setCellValue('B3', '資產名稱') 
                      ->setCellValue('C3', '使用人') 
                      ->setCellValue('D3', '保管人')
                      ->setCellValue('E3', '存放位置')  
                      ->setCellValue('F3', '資產類別')                       
                      ->setCellValue('G3', '數量')                          
                      ->setCellValue('H3', '盤號')   
                      ->setCellValue('I3', '盤點數量')  
                      ->setCellValue('J3', '保管部門')  
                      ->setCellValue('K3', '保管人')  
                      ->setCellValue('L3', '存放位置')  
                      ->setCellValue('M3', '使用人');
              $oldfca05=$row2['FCA05'];   
              $y=4;    
          }
        
             
          $osheet ->setCellValue('A'. $y, $row2["FCA03"]." ".$row2['FCA031'])   
                  ->setCellValue('B'. $y, $row2["FAJ06"])
                  ->setCellValue('C'. $y, $row2["FAJUD02"])                     
                  ->setCellValue('D'. $y, $row2["FCA06"].' '.$row2['GEN02'])
                  ->setCellValue('E'. $y, $row2["FCA07"].' '.$row2['FAF02'])  
                  ->setCellValue('F'. $y, $row2["FAJ04"].' '.$row2['FAB02'])   
                  ->setCellValue('G'. $y, $row2["FCA09"].' '.$row2['FCA08'])
                  ->setCellValue('H'. $y, $row2["FCA02"]);     
          $y++;    
        }
         
        $objStyleA3 = $osheet->getStyle('A3');
        $osheet ->duplicateStyle($objStyleA3, 'A3:M'.($y-1));   
        $objPHPExcel->setActiveSheetIndex(1);  

        header('Content-Type: application/vnd.ms-excel');  
        header('Content-Disposition: attachment;filename="'. 'InventoryCheck_' . $fca01 . '.xls"');    
        header('Cache-Control: max-age=0'); 
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');  

        $objWriter->save('php://output'); 

/* 存成檔案
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        // $filename='email/' . $tdate . '_DelayCasesReportv3.xls';
        // $objWriter->save($filename);
        $filename='inventory.xls';
        $objWriter->save($filename);
*/
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
<p>列印資產盤點清冊 </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">                                                                                                               
        盤點編號: 
        <select name="fca01" id="fca01">  
            <?
              $s1= "select distinct fca01 from fca_file order by fca01 desc ";
              $erp_sql1 = oci_parse($erp_conn1,$s1 );
              oci_execute($erp_sql1);  
              while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                  echo "<option value=" . $row1["FCA01"];  
                  if ($fca01 == $row1["FCA01"]) echo " selected";                  
                  echo ">" . $row1['FCA01'] . "</option>"; 
              }   
            ?>
        </select>                                                                                                                               
        &nbsp;&nbsp;                                                                                                                     
        &nbsp;&nbsp;  
        <input type="submit" name="submit" id="submit" value="Submit">  &nbsp;&nbsp;   &nbsp;&nbsp;      
        <input type="submit" name="submit" id="submit" value="Excel">  &nbsp;&nbsp;   &nbsp;&nbsp;      
      </td></tr>
    </table>
  </div>
</form>
<? if (is_null($_GET['submit'])) die ; ?>   
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
        <th width="16">&nbsp;</th> 
        <th>保管部門</th>         
        <th>存放位置</th>  
        <th>保管人</th>    
        <th>使用人</th>   
        <th>盤點序號</th>   
        <th>財產編號</th>       
        <th>資產類別</th> 
        <th>資產名稱</th>   
        <th>數量</th>   
        <th>單位</th>     
    </tr>
    <?
      //fca:盤點檔
      //faj:資產檔      
      //fab:資產類別
      
      //faf:存放位置
      //gen:人員檔
      //gem:部門檔
      $s1="select fca05,gem02, fca07,faf02, fca06,gen02, fajud02, fca02,fca03,fca031,fca04, faj04,fab02, faj06,fca09,fca08 " .  
          "from (select fca02,fca03,fca031,fca04,fca05,fca06,fca07,fca08,fca09,faj04,faj06,fajud02 from fca_file,faj_file where fca01='$fca01' and fca04=faj01 ) " .
          "left join gem_file on fca05=gem01 " .
          "left join faf_file on fca07=faf01 " .
          "left join gen_file on fca06=gen01 " .
          "left join fab_file on faj04=fab01 " .  
          "order by fca05,fca07 " ; 
      $erp_sql1 = oci_parse($erp_conn1,$s1 );
      oci_execute($erp_sql1);  
      $bgkleur = "ffffff";     
      $oldfac05='ZZZZZZZZ';
      while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {  
    ?>
          <tr bgcolor="#<?=$bgkleur;?>">    
              <td><img src="i/arrow.gif" width="16" height="16"></td>  
              <td><?=$row1["FCA05"]."--".$row1['GEM02'];?></td>     
              <td><?=$row1["FCA07"]."--".$row1['FAF02'];?></td> 
              <td><?=$row1["FCA06"]."--".$row1['GEN02'];?></td> 
              <td><?=$row1["FAJUD02"];?></td>    
              <td><?=$row1["FCA02"];?></td> 
              <td><?=$row1["FCA03"]." ".$row1['FCA031'];?></td> 
              <td><?=$row1["FAJ04"]."--".$row1['FAB02'];?></td> 
              <td><?=$row1["FAJ06"];?></td> 
              <td><?=$row1["FCA09"];?></td> 
              <td><?=$row1["FCA08"];?></td>                              
          </tr>
    <?    
      }
    ?>          
</table>   