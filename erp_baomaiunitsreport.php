<?php
  session_start();
  $pagetitle = "資材部 &raquo; 期間內包埋/沖膠床數統計表";
  include("_data.php");
  include("_erp.php");
 // auth("erp_baomaiunitsreport.php");  
  
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
  $key=$bdate.'--'.$edate; 
  
    if ($_GET["submit"]=="匯出") {   
        error_reporting(E_ALL);  
        require_once 'classes/PHPExcel.php'; 
        require_once 'classes/PHPExcel/IOFactory.php';  
        $objPHPExcel = new PHPExcel();
        //$objReader = PHPExcel_IOFactory::createReader('Excel5');  
        //$objPHPExcel = $objReader->load("templates/metaltakereport.xls"); 
        // Set properties
        $objPHPExcel ->getProperties()->setCreator( 'Frank')
                     ->setLastModifiedBy('Frank')
                     ->setTitle('期間內包埋/沖膠床數統計表')
                     ->setSubject('期間內包埋/沖膠床數統計表')
                     ->setDescription('期間內包埋/沖膠床數統計表')
                     ->setKeywords('期間內包埋/沖膠床數統計表')
                     ->setCategory('期間內包埋/沖膠床數統計表');
        
        // Add some data      
        $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
        $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);             
        $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
        $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(12);
        $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);                
        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A1', "$bdate ~~ $edate 包埋/沖膠床數統計表");   
        $objPHPExcel->getActiveSheet()->mergeCells('A1:C1');
      
        $y=2;      
        $i=1;                             
        $s2= "select to_char(tc_srg007,'mm-dd-yyyy') tc_srg007, sum(sfb08*imaud07) sfb08  " .
             "from tc_srg_file, sfb_file, ima_file " .
             "where tc_srg007 between to_date('$bdate','yy/mm/dd') and to_date('$edate','yy/mm/dd') and tc_srg030='3300' " .
             "and tc_srg001=sfb01 and sfb05=ima01 " .
             "group by tc_srg007 order by tc_srg007";
        $erp_sql2 = oci_parse($erp_conn1,$s2 );
        oci_execute($erp_sql2);  
        while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {      
            $objPHPExcel->setActiveSheetIndex(0)
                        ->setCellValue('A'. $y, $i)  
                        ->setCellValue('B'. $y, $row2["TC_SRG007"])    
                        ->setCellValue('C'. $y, $row2["SFB08"]); 
            if (($y%2)==0){
                $objPHPExcel->getActiveSheet()->getStyle('A'.$y.':C'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');    
            } 
            $y++; 
            $i++;   
        }                       
        $objPHPExcel->getActiveSheet() ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
        $objPHPExcel->getActiveSheet() ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
        $objPHPExcel->getActiveSheet() ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);   
        $objPHPExcel->setActiveSheetIndex(0)
                        ->setCellValue('B'. $y, '合計')
                        ->setCellValue('C'. $y, '=sum(C2:C' . ($y-1) . ')'); 
        // Rename sheet
        $objPHPExcel->getActiveSheet()->setTitle('包埋沖膠床數統計表');
                                                                
        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $objPHPExcel->setActiveSheetIndex(0);

        // Redirect output to a client’s web browser (Excel5)
        header('Content-Type: application/vnd.ms-excel');  
        header('Content-Disposition: attachment;filename="'. $bdate . " -- " .$edate. ' BaomaiUnitsReport.xls"');    
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
<p>期間內包埋/沖膠床數統計表 </p>     
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
        <th>日期</th>         
        <th>數量</th> 
    </tr>
    <?   
      //先刪除所有相同期間的資料            
      $s2= "select to_char(tc_srg007,'mm-dd-yyyy') tc_srg007, sum(sfb08*imaud07) sfb08  " .
           "from tc_srg_file, sfb_file, ima_file " .
           "where tc_srg007 between to_date('$bdate','yy/mm/dd') and to_date('$edate','yy/mm/dd') and tc_srg030='3300' " .
           "and tc_srg001=sfb01 and sfb05=ima01 " .
           "group by tc_srg007 order by tc_srg007";
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);  
      $bgkleur = "ffffff";      
      $total=0;                           
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {   
          $i++;
          $total+=$row2['SFB08']
          ?>
            <tr bgcolor="#<?=$bgkleur;?>">
                <td><?=$i;?></td>      
                <td><?=$row2["TC_SRG007"];?></td>  
                <td><div align=right><?=number_format($row2["SFB08"], 1, ".", ",");?></div></td>
            </tr>
      <?   
      }
      ?>    
      <tr bgcolor="#<?=$bgkleur;?>">
          <td></td>      
          <td>合計</td>  
          <td><div align=right><?=number_format($total, 1, ".", ",");?></div></td>
      </tr>
</table>   