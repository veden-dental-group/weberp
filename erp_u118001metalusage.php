<?php
  session_start();
  $pagetitle = "業務部 &raquo; EZ自備金使用報表";
  include("_data.php");
 //auth("erp_u118001metalusage.php");  
  
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
  
  //if (is_null($_GET['occ01'])) {
  //  $occ01 =  'U118001';
  //} else {
  //  $occ01=$_GET['occ01'];
  //}  
  
  if ($_GET["submit"]=="匯出") {                       
    $filename='templates/u118001metalusage.xls';
        
    error_reporting(E_NONE);  
    require_once 'classes/PHPExcel.php'; 
    require_once 'classes/PHPExcel/IOFactory.php';  
    $objReader = PHPExcel_IOFactory::createReader('Excel5');
    $objPHPExcel = $objReader->load($filename);  
    // Set properties
    $objPHPExcel ->getProperties()->setCreator('Frank' )
                 ->setLastModifiedBy('Frank')
                 ->setTitle('Metal Balance')
                 ->setSubject('Metal Balance')
                 ->setDescription('Metal Balance')
                 ->setKeywords('Metal Balance')
                 ->setCategory('Metal Balance');  
                 
    $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
    $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);             
    $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
    $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(12);
    $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);   
    //$objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣              
                
    $objPHPExcel->setActiveSheetIndex(0);             
    $osheet = $objPHPExcel->getActiveSheet();                   
    
    
    $s1="select img10 from img_file where img01='1K01000002'";
    $erp_sql1 = oci_parse($erp_conn1,$s1 );
    oci_execute($erp_sql1);     
    $row1 = oci_fetch_array($erp_sql1, OCI_ASSOC);   
    $osheet->setCellValue('B3',  $edate)
           ->setCellValue('E3',  $row1['IMG10']) ;    
    $y=5;         
    $i=1; 

    $s2= "select to_char(sfp03,'mm-dd-yyyy') sfp03, sfe02, sfe28, sfbud02, sfb08 ,sfpud08, tlf024 " .
           "from sfp_file, sfe_file, sfb_file, tlf_file " .
           "where sfp03  between to_date('$bdate','yy/mm/dd') and to_date('$edate','yy/mm/dd') and sfp06='1' " .                                                         
           "and sfp01=sfe02 and sfe07='1K01000002' " .             
           "and sfe01=sfb01 " .
           "and tlf01='1K01000002' and tlf026=sfe02 and tlf027=sfe28 ".
           "order by sfp03, sfe02, sfe28 desc "; //按發料單號, 工單順序
    $erp_sql2 = oci_parse($erp_conn1,$s2 );
    oci_execute($erp_sql2);   
    $oldsfe02='';
    while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {     
          if($oldsfe02!=$row2['SFE02']) {
            $sfpud08=floatval($row2["SFPUD08"]);
            $tlf024 =floatval($row2["TLF024"]);
            $sfp03  =$row2['SFP03'];
            $oldsfe02=$row2['SFE02'];
          } else {
            $sfpud08 = '';
            $tlf024  = '';
            $sfp03   = '';
          } 
          
          $osheet     ->setCellValue('A'. $y, $i)
                      ->setCellValue('B'. $y, $sfp03) 
                      ->setCellValue('C'. $y, $row2["SFBUD02"])  
                      ->setCellValue('D'. $y, $row2["SFB08"])
                      ->setCellValue('E'. $y, $sfpud08)  ;
                      //->setCellValue('F'. $y, $tlf024)
                      //->setCellValue('G'. $y, $row2["SFE02"] . ' ' . $row2['SFE28']) ;
          if (($y%2)==0){
              $osheet->getStyle('A'.$y.':E'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');    
          }     
          $osheet ->getStyle('E'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT); 
          //$osheet ->getStyle('F'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
          $osheet ->getStyle('E'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
          //$osheet ->getStyle('F'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 ); 
          $y++;
          $i++;
    }                      
    $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
    $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
    $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
    $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
    $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
    //$osheet ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
    //$osheet ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    if ($i>1) {  //i>1表示有資料 才做合計
      $objPHPExcel->setActiveSheetIndex(0)
                  ->setCellValue('A'.($y), 'Total')    
                  ->setCellValue('D'.($y), '=sum(D5:D' . ($y-1) . ')')                 
                  ->setCellValue('E'.($y), '=sum(E5:E' . ($y-1) . ')');  
    }
    $osheet ->getStyle('D'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);                     
    $osheet ->getStyle('D'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
                       
    // Rename sheet
    $objPHPExcel->getActiveSheet()->setTitle('Metal Balance');
                                                            
    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
    $objPHPExcel->setActiveSheetIndex(0);

    // Redirect output to a client’s web browser (Excel5)
    header('Content-Type: application/vnd.ms-excel');  
    header('Content-Disposition: attachment;filename="metalbalance_' . $edate . '.xls"');   
    header('Cache-Control: max-age=0'); 
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');  
    $objWriter->save('php://output'); 
    exit;    
  }           
  
  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>以下為 EZ自備金(1K01000002)金屬發料資料!! </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
        日期:   
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> ~~
        <input name="edate" type="text" id="edate" size="12" maxlength="12" value=<?=$edate;?> onfocus="new WdatePicker()">           
        <input type="submit" name="submit" id="submit" value="查詢">  &nbsp;&nbsp;   &nbsp;&nbsp;      
        <input type="submit" name="submit" id="submit" value="匯出">         
      </td></tr>
    </table>
  </div>
</form>  
          
<? if (is_null($_GET['submit'])) die ; ?>
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
        <th width="16">&nbsp;</th>
        <th>日期</th>  
        <th>發料單號</th>  
        <th>RX #</th> 
        <th>顆數</th>   
        <th>發料量</th>  
    </tr>
    <?
        
      $s2= "select to_char(sfp03,'yyyy-mm-dd') sfp03, sfe02, sfe28, sfbud02, sfb08 ,sfpud08, tlf024 " .
           "from sfp_file, sfe_file, sfb_file, tlf_file " .
           "where sfp03  between to_date('$bdate','yy/mm/dd') and to_date('$edate','yy/mm/dd') and sfp06='1' " .                                                         
           "and sfp01=sfe02 and sfe07='1K01000002' " .             
           "and sfe01=sfb01 " .
           "and tlf01='1K01000002' and tlf026=sfe02 and tlf027=sfe28 ".
           "order by sfp03, sfe02, sfe28 desc "; //按發料單號, 工單順序
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);   
      $amounttotal=0;
      $bgkleur = "ffffff";  
      $i=0;
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {      
          $i++;
    ?>
	        <tr bgcolor="#<?=$bgkleur;?>">
		          <td><?=$i;?></td>
              <td><?=$row2["SFP03"];?></td>  
              <td><?=$row2["SFE02"] . " " . $row2['SFE28'];?></td> 
              <td><?=$row2["SFBUD02"];?></td>   
              <td><?=$row2["SFB08"];?></td>    
              <td style="text-align:right" > <?=number_format($row2["SFPUD08"],3,'.',',');?> </td>  
            </tr>
		  <?
			}
      ?>       
</table>    