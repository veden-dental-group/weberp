<?php
  session_start();
  $pagetitle = "業務部 &raquo; 客戶金屬超重";
  include("_data.php");
 //auth("erp_metaloverweight.php");  
  
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
  
  if (is_null($_GET['occ01'])) {
    $occ01 =  'U121000';
  } else {
    $occ01=$_GET['occ01'];
  }  
  
  if (is_null($_GET['clienttype'])) {
    $clienttype='2';
  } else {
    $clienttype=$_GET['clienttype'];
  }
 
  
  if ($clienttype=='1') {  
    $occfilter = " and oga04='$occ01' ";
  } else if ($clienttype=='2') {
    $occfilter = " and oga18='$occ01' ";  
  } else if ($clienttype=='3') {      
    $occfilter = " and oga04 like 'T183%' "; 
  } else {  
    $occfilter = " and oga04 like 'H101%' "; 
  }  
  
  if ($_GET["submit"]=="匯出") {   
    $socc="select occ02 from occ_file where occ01='$occ01'";
    $erp_sqlocc = oci_parse($erp_conn2,$socc ); 
    oci_execute($erp_sqlocc);  
    $rowocc = oci_fetch_array($erp_sqlocc, OCI_ASSOC); 
    $occ02=$rowocc['OCC02'];
    $filename='templates/metaloverweight.xls';
        
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
                 
    $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
    $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);             
    $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
    $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
    $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);   
    //$objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣              
                
    $objPHPExcel->setActiveSheetIndex(0);             
    $osheet = $objPHPExcel->getActiveSheet();                   
                 
                 
    $s3="select occ18, occ241, occ242, occ261, occ271 from occ_file where occ01='$occ01'"; 
    $erp_sql3 = oci_parse($erp_conn2,$s3 );
    oci_execute($erp_sql3);  
    $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);      
    $tel='';  
    if ($row3['occ261']!='') $tel="Tel: " . $row3['occ261'] . "    ";
    if ($row3['occ271']!='') $tel.="Fax:" . $row3['occ271']; 
                        
    $y=6;
    $osheet -> setCellValue('A'.$y, $row3["OCC18"]);
    $y=7;
    $osheet -> setCellValue('A'.$y, $row3["OCC241"]);
    $y=8;
    $osheet -> setCellValue('A'.$y, $row3["OCC242"]);
    $y=9;
    $osheet -> setCellValue('A'.$y, $tel);  
    
    $y=12;         
    $i=1;   
    

    $s2= "select to_char(oga02,'mm/dd/yyyy') oga02, tc_dex011, ogb04, tc_dex001, tc_dex002, tc_dex003, tc_dex004, tc_dex013, ima1002,occ44, xmf07 from tc_dex_file, oga_file, ogb_file, ima_file,occ_file,xmf_file " .
           "where tc_dex013 < tc_dex004 " .
           "and tc_dex001=oga01 and oga01=ogb01 and tc_dex002=ogb03 $occfilter and oga02 between to_date('$bdate','yy/mm/dd') and to_date('$edate','yy/mm/dd')" .
           "and ogb04=ima01 " .
           "and oga04=occ01 " .
           "and xmf01=occ44 and tc_dex003=xmf03 " .
           "order by oga02, tc_dex011 ";
    $erp_sql2 = oci_parse($erp_conn1,$s2 );
    oci_execute($erp_sql2);   
    while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {     
          $tcdex004     =floatval($row2["TC_DEX004"]);
          $tcdex013     =floatval($row2["TC_DEX013"]);   
          $xmf07        =floatval($row2["XMF07"]);   
          $overweight   =$tcdex004-$tcdex013;  
          $osheet     ->setCellValue('A'. $y, $i)
                      ->setCellValue('B'. $y, $row2["OGA02"])
                      ->setCellValue('C'. $y, $row2["TC_DEX011"])  
                      ->setCellValue('D'. $y, $row2["IMA1002"]) 
                      ->setCellValue('E'. $y, $tcdex004)   
                      ->setCellValue('F'. $y, $overweight)   
                      ->setCellValue('G'. $y, $xmf07)
                      ->setCellValue('H'. $y, $overweight*$xmf07);   
          if (($y%2)==0){
              $osheet->getStyle('A'.$y.':H'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');    
          }     
          $osheet ->getStyle('E'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT); 
          $osheet ->getStyle('F'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT); 
          $osheet ->getStyle('G'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT); 
          $osheet ->getStyle('H'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);                     
          $osheet ->getStyle('E'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
          $osheet ->getStyle('F'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
          $osheet ->getStyle('G'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
          $osheet ->getStyle('H'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 ); 
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
    if ($i>1) {  //i>1表示有資料 才做合計
      $objPHPExcel->setActiveSheetIndex(0)
                  ->setCellValue('A'.($y), 'Total')                     
                  ->setCellValue('H'.($y), '=sum(H12:H' . ($y-1) . ')');  
    }
    $osheet ->getStyle('H'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);                     
    $osheet ->getStyle('H'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
                       
    // Rename sheet
    $objPHPExcel->getActiveSheet()->setTitle('Metal Overweight');
                                                            
    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
    $objPHPExcel->setActiveSheetIndex(0);

    // Redirect output to a client’s web browser (Excel5)
    header('Content-Type: application/vnd.ms-excel');  
    header('Content-Disposition: attachment;filename="metaloverweight_' . $occ01 .'_' . $edate . '.xls"');   
    header('Cache-Control: max-age=0'); 
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');  
    $objWriter->save('php://output'); 
    exit;    
  }           
  
  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>以下為客戶金屬超重資料!! </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
        日期:   
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> ~~
        <input name="edate" type="text" id="edate" size="12" maxlength="12" value=<?=$edate;?> onfocus="new WdatePicker()">
        客戶代碼: 
        <select name="occ01" id="occ01">  
            <?
              $s1= "select occ01,occ02 from occ_file order by occ01 ";
              $erp_sql1 = oci_parse($erp_conn2,$s1 );
              oci_execute($erp_sql1);  
              while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                  echo "<option value=" . $row1["OCC01"];  
                  if ($occ01 == $row1["OCC01"]) echo " selected";                  
                  echo ">" . $row1['OCC01'] ."--" .$row1["OCC02"] . "</option>"; 
              }   
            ?>
        </select>  &nbsp;&nbsp;   &nbsp;&nbsp;     
        客戶種類: 
        <input name="clienttype" type="radio" value="1" id="clienttype1" <?if($clienttype=='1') echo " checked";?>><label for="clienttype1">一般客戶 </label>&nbsp; 
        <input name="clienttype" type="radio" value="2" id="clienttype2" <?if($clienttype=='2') echo " checked";?>><label for="clienttype2">集團碼 </label>&nbsp; 
        <input name="clienttype" type="radio" value="3" id="clienttype3" <?if($clienttype=='3') echo " checked";?>><label for="clienttype3">澳門客戶群 </label>   
        <input name="clienttype" type="radio" value="4" id="clienttype3" <?if($clienttype=='4') echo " checked";?>><label for="clienttype4">HK客戶群 </label>
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
        <th>CASE NO.</th>  
        <th>Product Code</th> 
        <th>Product Description</th>   
        <th>Total Weight</th>  
        <th>Standard Weight</th>   
        <th>OverWeight</th>  
        <th>Price/Gram</th>  
        <th>OverWeight Charge</th>  
    </tr>
    <?
      //$bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);
      //$edate1=substr($edate,2,2).substr($edate,5,2).substr($edate,8,2);  
      $s2= "select to_char(oga02,'mm/dd/yyyy') oga02, tc_dex011, ogb04, tc_dex001, tc_dex002, tc_dex003, tc_dex004, tc_dex013, ima1002,occ44, xmf07 from tc_dex_file, oga_file, ogb_file, ima_file,occ_file,xmf_file " .
           "where tc_dex013 < tc_dex004 " .
           "and tc_dex001=oga01 and oga01=ogb01 and tc_dex002=ogb03 $occfilter and oga02 between to_date('$bdate','yy/mm/dd') and to_date('$edate','yy/mm/dd')" .
           "and ogb04=ima01 " .
           "and oga04=occ01 " .
           "and xmf01=occ44 and tc_dex003=xmf03 " .
           "order by oga02, tc_dex011 ";
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);   
      $amounttotal=0;
      $bgkleur = "ffffff";  
      $i=0;
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) { 
          $tcdex004     =floatval($row2["TC_DEX004"]);
          $tcdex013     =floatval($row2["TC_DEX013"]);   
          $xmf07        =floatval($row2["XMF07"]);   
          $overweight   =$tcdex004-$tcdex013;  
          $amounttotal += ($overweight*$xmf07); 
          $i++;
    ?>
	        <tr bgcolor="#<?=$bgkleur;?>">
		          <td><?=$i;?></td>
              <td><?=$row2["OGA02"];?></td>  
              <td><?=$row2["TC_DEX011"];?></td> 
              <td><?=$row2["OGB04"];?></td>   
              <td><?=$row2["IMA1002"];?></td>      
              <td style="text-align: right"><?=$tcdex004;?></td>  
              <td style="text-align: right"><?=$tcdex013;?></td> 
              <td style="text-align: right"><?=$overweight;?></td>    
              <td style="text-align: right"><?=$xmf07;?></td>                     
              <td style="text-align: right"><?=$overweight*$xmf07;?></td>  
            </tr>
		  <?
			}
      ?>
      <tr bgcolor="#<?=$bgkleur;?>">
        <td><img src="i/arrow.gif" width="16" height="16"></td>
        <td colspan="8">Total</td> 
        <td style="text-align: right"><?=$amounttotal;?></td>    
      </tr>  
</table>    