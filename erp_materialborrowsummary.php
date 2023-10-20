<?php
  session_start();
  $pagetitle = "資材部 &raquo; 物料借入未還統計表";
  include("_data.php");
  //auth("erp_materialborrowsummary.php");  
  
  if (is_null($_GET['edate'])) {
    $edate = date("Y-m-d") ;     
  } else {      
    $edate = $_GET['edate'] ;   
  }                                                                            

  if ($_GET["submit"]=="匯出") {   

    $filename='templates/materialborrowsummary.xls';
        
    error_reporting(E_NONE);  
    require_once 'classes/PHPExcel.php'; 
    require_once 'classes/PHPExcel/IOFactory.php';  
    $objReader = PHPExcel_IOFactory::createReader('Excel5');
    $objPHPExcel = $objReader->load($filename);  
    // Set properties                      
    $objPHPExcel ->getProperties()->setCreator('Frank' )
                 ->setLastModifiedBy('Frank')
                 ->setTitle('Frank')
                 ->setSubject('Frank')
                 ->setDescription('Frank')
                 ->setKeywords('Frank')
                 ->setCategory('Frank');  
    $objPHPExcel->setActiveSheetIndex(0);             
    $osheet = $objPHPExcel->getActiveSheet();              
                 
    $osheet->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
    $osheet->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);             
    $osheet->getDefaultStyle()->getFont()->setName('Calibri');
    $osheet->getDefaultStyle()->getFont()->setSize(12);
    $osheet->getDefaultRowDimension()->setRowHeight(15);              
                 
    $osheet     ->setCellValue('A2', "截止日期: $edate") ;               
    
    $s2 = "select pmc01, pmc03, ima01, ima02, ima021, ima25, imo02, imp04, imq07 from  " .
            "(select  imo03, imo02, imp03, imp04, imq07  from " .
            " (select imo03, to_char(imo02,'yyyy-mm-dd') imo02, imp01, imp02, imp03, imp04 from imp_file,imo_file where imp01=imo01 and imopost='Y' and imo02<=to_date('$edate','yy/mm/dd') ) " .
            " left join " .
            " (select imq03, imq04, imq05, sum(imq07) imq07 from imq_file,imr_file where imq01=imr01 and imrpost='Y' and imr09<=to_date('$edate','yy/mm/dd') group by imq03, imq04, imq05 ) on imp01=imq03 and imp02=imq04 ), " .
           "ima_file, pmc_file where imp03=ima01 and (imp04!=imq07 or imq07 is null )and imo03=pmc01 order by pmc01,ima01,imo02 "; 

    $erp_sql2 = oci_parse($erp_conn,$s2 );
    oci_execute($erp_sql2);   
    $y=4;         
    $i=1;  
    while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {               
          $osheet     ->setCellValue('A'. $y, $i)
                      ->setCellValue('B'. $y, $row2["PMC01"])
                      ->setCellValue('C'. $y, $row2["PMC03"])
                      ->setCellValue('D'. $y, $row2["IMA01"])
                      ->setCellValue('E'. $y, $row2["IMA02"])  
                      ->setCellValue('F'. $y, $row2["IMA021"])  
                      ->setCellValue('G'. $y, $row2["IMA25"])
                      ->setCellValue('H'. $y, $row2["IMO02"])  
                      ->setCellValue('I'. $y, $row2["IMP04"])  
                      ->setCellValue('J'. $y, $row2["IMQ07"])
                      ->setCellValue('K'. $y, $row2['IMP04']-$row2["IMQ07"]);
                        
          if (($y%2)==0){
              $osheet->getStyle('A'.$y.':K'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');    
          }  
          $y++;
          $i++;
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
    $osheet ->setTitle('物料借入未還統計表');
                                                            
    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
    $objPHPExcel->setActiveSheetIndex(0);

    // Redirect output to a client’s web browser (Excel5)
    header('Content-Type: application/vnd.ms-excel');  
    header('Content-Disposition: attachment;filename="'. $edate . '_materialborrowsummary.xls"');   
    header('Cache-Control: max-age=0'); 
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');  
    $objWriter->save('php://output'); 
    exit;    
  }   
  
  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>以下為各物料借入未還統計表!! </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
        截止日期:   
        <input name="edate" type="text" id="edate" size="12" maxlength="12" value=<?=$edate;?> onfocus="new WdatePicker()"> &nbsp;&nbsp;                                                                                                     
        <input type="submit" name="submit" id="submit" value="查詢">  &nbsp;&nbsp;     
        <input type="submit" name="submit" id="submit" value="匯出">         
      </td></tr>
    </table>
  </div>
</form>  
          
<? if (is_null($_GET['submit'])) die ; ?>
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
        <th width="16">&nbsp;</th>
        <th>廠商編號</th>  
        <th>廠商名稱</th>  
        <th>物料編號</th>
        <th>物料名稱</th>            
        <th>規格</th>   
        <th>單位</th>
        <th>借入日期</th>    
        <th style="text-align:right">借入數量</th>  
        <th style="text-align:right">還出數量</th>     
        <th style="text-align:right">未還數量</th>      
    </tr>
    <?
      $s2 = "select pmc01, pmc03, ima01, ima02, ima021, ima25, imo02, imp04, imq07 from  " .
            "(select  imo03, imo02, imp03, imp04, imq07  from " .
            " (select imo03, to_char(imo02,'yyyy-mm-dd') imo02, imp01, imp02, imp03, imp04 from imp_file,imo_file where imp01=imo01 and imopost='Y' and imo02<=to_date('$edate','yy/mm/dd') ) " .
            " left join " .
            " (select imq03, imq04, imq05, sum(imq07) imq07 from imq_file,imr_file where imq01=imr01 and imrpost='Y' and imr09<=to_date('$edate','yy/mm/dd') group by imq03, imq04, imq05 ) on imp01=imq03 and imp02=imq04 ), " .
           "ima_file, pmc_file where imp03=ima01 and (imp04!=imq07 or imq07 is null )and imo03=pmc01 order by pmc01,ima01,imo02 "; 
      
      $erp_sql2 = oci_parse($erp_conn,$s2 );
      oci_execute($erp_sql2);  
      $bgkleur = "ffffff";  
      $i=0;
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) { 
          $i++;
    ?>
	        <tr bgcolor="#<?=$bgkleur;?>">
		          <td><?=$i;?></td>
              <td><?=$row2["PMC01"];?></td>  
              <td><?=$row2["PMC03"];?></td>  
              <td><?=$row2["IMA01"];?></td>  
              <td><?=$row2["IMA02"];?></td>                                 
              <td><?=$row2["IMA021"];?></td>
              <td><?=$row2["IMA25"];?></td> 
              <td><?=$row2["IMO02"];?></td> 
		          <td style="text-align:right"><?=number_format($row2["IMP04"], 2);?></td>
              <td style="text-align:right"><?=number_format($row2["IMQ07"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["IMP04"]-$row2['IMQ07'], 2);?></td>   
          </tr>
		  <?
			}
      ?>     
</table>    