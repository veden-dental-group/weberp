<?php
  session_start();
  $pagetitle = "財務部 &raquo; 每月各種產品到貨統計表";
  include("_data.php");
  auth("erp_casesinmonthlysummary.php");  
  
  if (is_null($_GET['thisyear'])) {
    $thisyear = date("Y") ;     
  } else {      
    $thisyear = $_GET['thisyear'] ;   
  }  
  
  if ($_GET["submit"]=="匯出") {   

    $filename='templates/casesinmonthlysummary.xls';
        
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
                ->setCellValue('A1', $thisyear.' 年各种产品到货统计');       
                      
    $y=4;      
    $s2="select oea02, sum(a1) a1, sum(a2) a2, sum(a3) a3,  sum(a4) a4, sum(b1) b1, sum(b2) b2, sum(b3) b3, sum(b4) b4, sum(b5) b5 from " .
          "  (select oea02, decode(ima06,'9111', total,'9119',total,0) a1, decode(ima06,'9112', total, 0) a2, " .
          "                 decode(ima06,'9113', total,'9117',total,'9118',total,'91181',total,0) a3, decode(ima06,'9114', total,'9115',total, '9116',total, '911Z', total, 0) a4, " .
          "                 decode(ima06,'9121', total,'9122',total,'9126',total, 0) b1, decode(ima06,'9123', total,'9125',total, 0) b2, decode(ima06,'91241', total,'91242',total,'91243',total, '9125',total, 0) b3, " .
          "                 decode(ima06,'9127', total, 0) b4, decode(ima06,'912Y', total, '912Z', total, 0) b5 from " .
          "    (select to_char(oea02, 'yyyy-mm') oea02, ima06, (oeb12 * imaud07) total from oeb_file, oea_file, ima_file where oeb01=oea01 and to_char(oea02,'yyyy')='$thisyear' and oeb04=ima01)) " .
          "group by oea02 order by oea02 "; 
    $erp_sql2 = oci_parse($erp_conn,$s2 );
    oci_execute($erp_sql2);   
    while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {     
          $objPHPExcel->setActiveSheetIndex(0)
                      ->setCellValue('A'. $y, $row2['OEA02'])
                      ->setCellValue('B'. $y, $row2["A1"])
                      ->setCellValue('C'. $y, $row2["A2"])
                      ->setCellValue('D'. $y, $row2["A3"])   
                      ->setCellValue('E'. $y, $row2["A4"])      
                      ->setCellValue('F'. $y, "=sum(B$y:E$y)")   
                      ->setCellValue('G'. $y, $row2["B1"])   
                      ->setCellValue('H'. $y, $row2["B2"])  
                      ->setCellValue('I'. $y, $row2["B3"])  
                      ->setCellValue('J'. $y, $row2["B4"]) 
                      ->setCellValue('K'. $y, $row2["B5"]) 
                      ->setCellValue('L'. $y, "=sum(G$y:K$y)")  
                      ->setCellValue('M'. $y, "=F$y+L$y");       
          $y++;                                                   
    }                      
                
    $objPHPExcel->getActiveSheet()->setTitle('各种产品到货统计');
                                                            
    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
    $objPHPExcel->setActiveSheetIndex(0);

    // Redirect output to a client’s web browser (Excel5)
    header('Content-Type: application/vnd.ms-excel');  
    header('Content-Disposition: attachment;filename="'. $thismonth . '_monthlyusage.xls"');   
    header('Cache-Control: max-age=0'); 
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');  
    $objWriter->save('php://output'); 
    exit;    
  }   
  
  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>以下為年度各種產品到貨月統計表 </p>     
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
        <th>月份</th>   
        <th style="text-align:right">一般烤瓷牙</th>  
        <th style="text-align:right">鋼牙</th> 
        <th style="text-align:right">全瓷系列</th>
        <th style="text-align:right">其它</th>
        <th style="text-align:right">固定小計</th>  
        <th style="text-align:right">金屬床</th>    
        <th style="text-align:right">彈性床</th>        
        <th style="text-align:right">樹脂床</th>    
        <th style="text-align:right">SOS</th> 
        <th style="text-align:right">其他</th>    
        <th style="text-align:right">活動小計</th>  
        <th style="text-align:right">全部合計</th>  
    </tr>
    <?
      $s2="select oea02, sum(a1) a1, sum(a2) a2, sum(a3) a3,  sum(a4) a4, sum(b1) b1, sum(b2) b2, sum(b3) b3, sum(b4) b4, sum(b5) b5 from " .
          "  (select oea02, decode(ima06,'9111', total,'9119',total,0) a1, decode(ima06,'9112', total, 0) a2, " .
          "                 decode(ima06,'9113', total,'9117',total,'9118',total,'91181',total,0) a3, decode(ima06,'9114', total,'9115',total, '9116',total, '911Z', total, 0) a4, " .
          "                 decode(ima06,'9121', total,'9122',total,'9126',total, 0) b1, decode(ima06,'9123', total,'9125',total, 0) b2, decode(ima06,'91241', total,'91242',total,'91243',total, '9125',total, 0) b3, " .
          "                 decode(ima06,'9127', total, 0) b4, decode(ima06,'912Y', total, '912Z', total, 0) b5 from " .
          "    (select to_char(oea02, 'yyyy-mm') oea02, ima06, (oeb12 * imaud07) total from oeb_file, oea_file, ima_file where oeb01=oea01 and to_char(oea02,'yyyy')='$thisyear' and oeb04=ima01)) " .
          "group by oea02 order by oea02 ";                                                                                                                                   
      
      $erp_sql2 = oci_parse($erp_conn,$s2 );
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
          $totala1 += $row2['A1'];
          $totala2 += $row2['A2']; 
          $totala3 += $row2['A3']; 
          $totala4 += $row2['A4']; 
          $totalb1 += $row2['B1']; 
          $totalb2 += $row2['B2']; 
          $totalb3 += $row2['B3']; 
          $totalb4 += $row2['B4']; 
          $totalb5 += $row2['B5']; 
          
    ?>
	        <tr bgcolor="#<?=$bgkleur;?>">
		          <td><img src="i/arrow.gif" width="16" height="16">  
              <td><?=$row2["OEA02"];?></td>   
		          <td style="text-align:right"><?=number_format($row2["A1"], 2);?></td>
              <td style="text-align:right"><?=number_format($row2["A2"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["A3"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["A4"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["A1"]+$row2["A2"]+$row2["A3"]+$row2["A4"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["B1"], 2);?></td>                                                                      
              <td style="text-align:right"><?=number_format($row2["B2"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["B3"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["B4"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["B5"], 2);?></td>   
              <td style="text-align:right"><?=number_format($row2["B1"]+$row2["B2"]+$row2["B3"]+$row2["B4"]+$row2["B5"], 2);?></td>                            
              <td style="text-align:right"><?=number_format($row2["A1"]+$row2["A2"]+$row2["A3"]+$row2["A4"]+$row2["B1"]+$row2["B2"]+$row2["B3"]+$row2["B4"]+$row2["B5"], 2);?></td> 
          </tr>
		  <?
			}
      ?>  
      <tr bgcolor="#<?=$bgkleur;?>">
              <td><img src="i/arrow.gif" width="16" height="16">  
              <td>合計</td>   
              <td style="text-align:right"><?=number_format($totala1, 2);?></td>
              <td style="text-align:right"><?=number_format($totala2, 2);?></td> 
              <td style="text-align:right"><?=number_format($totala3, 2);?></td> 
              <td style="text-align:right"><?=number_format($totala4, 2);?></td> 
              <td style="text-align:right"><?=number_format($totala1+$totala2+$totala3+$totala4, 2);?></td> 
              <td style="text-align:right"><?=number_format($totalb1, 2);?></td>                                                                      
              <td style="text-align:right"><?=number_format($totalb2, 2);?></td> 
              <td style="text-align:right"><?=number_format($totalb3, 2);?></td> 
              <td style="text-align:right"><?=number_format($totalb4, 2);?></td>  
              <td style="text-align:right"><?=number_format($totalb5, 2);?></td>   
              <td style="text-align:right"><?=number_format($totalb1+$totalb2+$totalb3+$totalb4+$totalb5, 2);?></td>                            
              <td style="text-align:right"><?=number_format($totala1+$totala2+$totala3+$totala4+$totalb1+$totalb2+$totalb3+$totalb4+$totalb5, 2);?></td> 
      </tr>   
</table>    