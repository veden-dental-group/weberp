<?php
  session_start();
  $pagetitle = "財務部 &raquo; 每月各區域產品到貨統計表";
  include("_data.php");
  //auth("erp_casesinmonthlysummary.php");  
  
  if (is_null($_GET['thisyear'])) {
    $thisyear = date("Y") ;     
  } else {      
    $thisyear = $_GET['thisyear'] ;   
  }  
  
  if ($_GET["submit"]=="匯出") {   

    $filename='templates/casesinmonthlysummary_byarea.xls';
        
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
                 
    $objPHPExcel->setActiveSheetIndex(0)
                ->setCellValue('A1', $thisyear.' 年各区域到貨分析表');       
                      
      
    $s2="select gea01, gea02, sum(m1) m1, sum(m2) m2, sum(m3) m3,  sum(m4) m4, sum(m5) m5, sum(m6) m6, sum(m7) m7, sum(m8) m8, sum(m9) m9, sum(m10) m10, sum(m11) m11, sum(m12) m12 from " .
        "  (select gea01, gea02, decode(mm,'01', total,0) m1, decode(mm,'02', total, 0) m2,  decode(mm,'03', total,0) m3,  decode(mm,'04', total, 0) m4, " .
        "                        decode(mm,'05', total,0) m5, decode(mm,'06', total, 0) m6,  decode(mm,'07', total,0) m7,  decode(mm,'08', total, 0) m8, " .
        "                        decode(mm,'09', total,0) m9, decode(mm,'10', total, 0) m10, decode(mm,'11', total,0) m11, decode(mm,'12', total, 0) m12 from " .
        "    (select gea01, gea02, to_char(oea02, 'mm') mm, (oeb12 * imaud07) total from oeb_file, oea_file, ima_file, occ_file, gea_file " .
              "where oeb01=oea01 and to_char(oea02,'yyyy')='$thisyear' and oeb04=ima01 and oea04=occ01 and occ20=gea01 )) " .
        "group by gea01, gea02 order by gea01,gea02 ";   
    $erp_sql2 = oci_parse($erp_conn1,$s2 );
    oci_execute($erp_sql2);   
    $i=1;
    while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {         
          $x=chr($i+65); 
          $objPHPExcel->setActiveSheetIndex(0)
                      ->setCellValue($x.'2',  $row2['GEA02'])  
                      ->setCellValue($x.'3',  $row2['M1'])
                      ->setCellValue($x.'4',  $row2['M2'])   
                      ->setCellValue($x.'5',  $row2['M3'])   
                      ->setCellValue($x.'6',  $row2['M4'])   
                      ->setCellValue($x.'7',  $row2['M5'])   
                      ->setCellValue($x.'8',  $row2['M6'])   
                      ->setCellValue($x.'9',  $row2['M7'])   
                      ->setCellValue($x.'10', $row2['M8'])   
                      ->setCellValue($x.'11', $row2['M9'])   
                      ->setCellValue($x.'12', $row2['M10'])   
                      ->setCellValue($x.'13', $row2['M11'])   
                      ->setCellValue($x.'14', $row2['M12']);    
          $i++;                                                                            
    }                      
                
    $objPHPExcel->getActiveSheet()->setTitle('每月各區域產品到貨統計表');
                                                            
    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
    $objPHPExcel->setActiveSheetIndex(0);

    // Redirect output to a client’s web browser (Excel5)
    header('Content-Type: application/vnd.ms-excel');  
    header('Content-Disposition: attachment;filename="'. $thisyear . '_monthlysummary_byarea.xls"');   
    header('Cache-Control: max-age=0'); 
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');  
    $objWriter->save('php://output'); 
    exit;    
  }   
  
  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>以下為每月各區域產品到貨統計表 </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
        年份:   
        <input name="thisyear" type="text" id="thisyear" onfocus="WdatePicker({dateFmt:'yyyy'})" value="<?=$thisyear;?>"> &nbsp;&nbsp;  
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
        <th>地區別</th>    
        <th style="text-align:right">一月</th>  
        <th style="text-align:right">二月</th> 
        <th style="text-align:right">三月</th>
        <th style="text-align:right">四月</th>
        <th style="text-align:right">五月</th>  
        <th style="text-align:right">六月</th>    
        <th style="text-align:right">七月</th>        
        <th style="text-align:right">八月</th>    
        <th style="text-align:right">九月</th> 
        <th style="text-align:right">十月</th>    
        <th style="text-align:right">十一月</th>  
        <th style="text-align:right">十二月</th> 
        <th style="text-align:right">合計</th>  
    </tr>
    <?
      $s2="select gea01, gea02, sum(m1) m1, sum(m2) m2, sum(m3) m3,  sum(m4) m4, sum(m5) m5, sum(m6) m6, sum(m7) m7, sum(m8) m8, sum(m9) m9, sum(m10) m10, sum(m11) m11, sum(m12) m12 from " .
          "  (select gea01, gea02, decode(mm,'01', total,0) m1, decode(mm,'02', total, 0) m2,  decode(mm,'03', total,0) m3,  decode(mm,'04', total, 0) m4, " .
          "                        decode(mm,'05', total,0) m5, decode(mm,'06', total, 0) m6,  decode(mm,'07', total,0) m7,  decode(mm,'08', total, 0) m8, " .
          "                        decode(mm,'09', total,0) m9, decode(mm,'10', total, 0) m10, decode(mm,'11', total,0) m11, decode(mm,'12', total, 0) m12 from " .
          "    (select gea01, gea02, to_char(oea02, 'mm') mm, (oeb12 * imaud07) total from oeb_file, oea_file, ima_file, occ_file, gea_file " .
                "where oeb01=oea01 and to_char(oea02,'yyyy')='$thisyear' and oeb04=ima01 and oea04=occ01 and occ20=gea01 )) " .
          "group by gea01, gea02 order by gea01,gea02 ";                                                                                                                                   
      
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);  
      $bgkleur = "ffffff";  
      $i=0;
      $totala1=0;
      $totala2=0;
      $totala3=0;
      $totala4=0;
      $totalb1=0;
      $totalb2=0; 
      $totalb3=0; 
      $totalb4=0; 
      $totalb5=0; 
      
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) { 
          $i++;
          $totalm1  += $row2['M1'];
          $totalm2  += $row2['M2']; 
          $totalm3  += $row2['M3']; 
          $totalm4  += $row2['M4']; 
          $totalm5  += $row2['M5']; 
          $totalm6  += $row2['M6']; 
          $totalm7  += $row2['M7']; 
          $totalm8  += $row2['M8']; 
          $totalm9  += $row2['M9']; 
          $totalm10 += $row2['M10'];  
          $totalm11 += $row2['M11'];  
          $totalm12 += $row2['M12'];  
          
    ?>
	        <tr bgcolor="#<?=$bgkleur;?>">
		          <td><img src="i/arrow.gif" width="16" height="16">  
              <td><?=$row2["GEA02"];?></td>   
		          <td style="text-align:right"><?=number_format($row2["M1"], 2);?></td>
              <td style="text-align:right"><?=number_format($row2["M2"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["M3"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["M4"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["M5"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["M6"], 2);?></td>                                                                      
              <td style="text-align:right"><?=number_format($row2["M7"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["M8"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["M9"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["M10"], 2);?></td>   
              <td style="text-align:right"><?=number_format($row2["M11"], 2);?></td>     
              <td style="text-align:right"><?=number_format($row2["M12"], 2);?></td>                         
              <td style="text-align:right"><?=number_format($row2["M1"]+$row2["M2"]+$row2["M3"]+$row2["M4"]+$row2["M5"]+$row2["M6"]+$row2["M7"]+$row2["M8"]+$row2["M9"]+$row2["M10"]+$row2["M11"]+$row2["M12"], 2);?></td> 
          </tr>
		  <?
			}
      ?>  
      <tr bgcolor="#<?=$bgkleur;?>">
              <td><img src="i/arrow.gif" width="16" height="16">  
              <td>合計</td>   
              <td style="text-align:right"><?=number_format($totalm1, 2);?></td>
              <td style="text-align:right"><?=number_format($totalm2, 2);?></td> 
              <td style="text-align:right"><?=number_format($totalm3, 2);?></td> 
              <td style="text-align:right"><?=number_format($totalm4, 2);?></td> 
              <td style="text-align:right"><?=number_format($totalm5, 2);?></td> 
              <td style="text-align:right"><?=number_format($total6, 2);?></td>                                                                      
              <td style="text-align:right"><?=number_format($totalm7, 2);?></td> 
              <td style="text-align:right"><?=number_format($totalm8, 2);?></td> 
              <td style="text-align:right"><?=number_format($totalm9, 2);?></td>  
              <td style="text-align:right"><?=number_format($totalm10, 2);?></td>            
              <td style="text-align:right"><?=number_format($totalm11, 2);?></td>       
              <td style="text-align:right"><?=number_format($totalm12, 2);?></td>                                                         
              <td style="text-align:right"><?=number_format($totalm1+$totalm2+$totalm3+$totalm4+$totalm5+$totalm6+$totalm7+$totalm8+$totalm9+$totalm10+$totalm11+$totalm12, 2);?></td> 
      </tr>   
</table>    