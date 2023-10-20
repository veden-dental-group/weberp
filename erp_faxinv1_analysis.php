<?php
  session_start();
  $pagetitle = "業務部 &raquo; 期間內傳真原因統計";
  include("_data.php");
  include("_erp.php");
  //auth("erp_faxinv1_analysis.php");
  
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
      error_reporting(E_NONE);  
      require_once 'classes/PHPExcel.php'; 
      require_once 'classes/PHPExcel/IOFactory.php';  
      $objReader = PHPExcel_IOFactory::createReader('Excel5');
      $objPHPExcel = $objReader->load("templates/erp_faxv1_analysis.xls");

      //第二個sheet放到貨組數
      $objPHPExcel->setActiveSheetIndex(0)
                  ->setCellValue('B1', $bdate . " -- " . $edate);
                  

      $s2="select occ02, sum(a01) a01, sum(a02) a02, sum(a03) a03, sum(a04) a04, sum(a05) a05, sum(a06) a06,
                         sum(a07) a07, sum(a08) a08, sum(a09) a09, sum(a10) a10, sum(a11) a11, sum(a12) a12,
                         sum(a13) a13, sum(a14) a14, sum(a15) a15
           from
              (select occ02,
                  decode(a.tc_ohf007,'空間不夠',1, 0) as a01,               
                  decode(a.tc_ohf007,'Margin不清',1, 0) as a02,               
                  decode(a.tc_ohf007,'確認咬合',1, 0) as a03,               
                  decode(a.tc_ohf007,'模子有损',1, 0) as a04,               
                  decode(a.tc_ohf007,'支臺齒不平行；有倒凹',1, 0) as a05,               
                  decode(a.tc_ohf007,'确认產品',1, 0) as a06,               
                  decode(a.tc_ohf007,'確認比色',1, 0) as a07,               
                  decode(a.tc_ohf007,'確認翻譯',1, 0) as a08,               
                  decode(a.tc_ohf007,'確認齒位',1, 0) as a09,               
                  decode(a.tc_ohf007,'印模不好',1, 0) as a10,               
                  decode(a.tc_ohf007,'索配件',1, 0) as a11,               
                  decode(a.tc_ohf007,'確認設計',1, 0) as a12,               
                  decode(a.tc_ohf007,'确认蜡型',1, 0) as a13,               
                  decode(a.tc_ohf007,'確認Post 分體/連體',1, 0) as a14,               
                  decode(a.tc_ohf007,'others',1, 0) as a15               
              from
                  (select distinct concat(sfb22, tc_ohf007), sfb22, occ02, tc_ohf007 from tc_ohf_file, sfb_file, oea_file, occ_file 
                    where tc_ohf004 between to_date('$bdate','yy/mm/dd') and to_date('$edate','yy/mm/dd')
                      and tc_ohf001 = sfb01 and sfb22=oea01 and oea04=occ01
                  ) a 
              ) 
           group by occ02 order by occ02";
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);         
      $y=3;
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {    
          $objPHPExcel->setActiveSheetIndex(0)
                      ->setCellValue('A'.($y), $row2['OCC02'])
                      ->setCellValue('B'.($y), $row2['A01'])
                      ->setCellValue('C'.($y), $row2['A02'])
                      ->setCellValue('D'.($y), $row2['A03'])
                      ->setCellValue('E'.($y), $row2['A04'])
                      ->setCellValue('F'.($y), $row2['A05'])
                      ->setCellValue('G'.($y), $row2['A06'])
                      ->setCellValue('H'.($y), $row2['A07'])
                      ->setCellValue('I'.($y), $row2['A08'])
                      ->setCellValue('J'.($y), $row2['A09'])
                      ->setCellValue('K'.($y), $row2['A10'])
                      ->setCellValue('L'.($y), $row2['A11'])
                      ->setCellValue('M'.($y), $row2['A12'])
                      ->setCellValue('N'.($y), $row2['A13'])
                      ->setCellValue('O'.($y), $row2['A14'])
                      ->setCellValue('P'.($y), $row2['A15'])
                      ->setCellValue('Q'.($y), '=sum(B' . ($y) . ':P' . ($y) . ')');

       
          $y++;
      }            
      $objPHPExcel->setActiveSheetIndex(0)
                  ->setCellValue('A'.$y, '合計')
                  ->setCellValue('B'.$y, '=sum(B3:B' . ($y-1) . ')')
                  ->setCellValue('C'.$y, '=sum(C3:C' . ($y-1) . ')')
                  ->setCellValue('D'.$y, '=sum(D3:D' . ($y-1) . ')')
                  ->setCellValue('E'.$y, '=sum(E3:E' . ($y-1) . ')')
                  ->setCellValue('F'.$y, '=sum(F3:F' . ($y-1) . ')')
                  ->setCellValue('G'.$y, '=sum(G3:G' . ($y-1) . ')')
                  ->setCellValue('H'.$y, '=sum(H3:H' . ($y-1) . ')')
                  ->setCellValue('I'.$y, '=sum(I3:I' . ($y-1) . ')')
                  ->setCellValue('J'.$y, '=sum(J3:J' . ($y-1) . ')')
                  ->setCellValue('K'.$y, '=sum(K3:K' . ($y-1) . ')')
                  ->setCellValue('L'.$y, '=sum(L3:L' . ($y-1) . ')')
                  ->setCellValue('M'.$y, '=sum(M3:M' . ($y-1) . ')')
                  ->setCellValue('N'.$y, '=sum(N3:N' . ($y-1) . ')')
                  ->setCellValue('O'.$y, '=sum(O3:O' . ($y-1) . ')')
                  ->setCellValue('P'.$y, '=sum(P3:P' . ($y-1) . ')')
                  ->setCellValue('Q'.$y, '=sum(Q3:Q' . ($y-1) . ')');

                  
      $objPHPExcel->getActiveSheet()->setTitle('Fax-in ');
                                                                                                                                         
      // Set active sheet index to the first sheet, so Excel opens this as the first sheet
      $objPHPExcel->setActiveSheetIndex(0);

      // Redirect output to a client’s web browser (Excel5)
      header('Content-Type: application/vnd.ms-excel');  
      header('Content-Disposition: attachment;filename="faxin_analysis.xls"');
      header('Cache-Control: max-age=0'); 
      $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');  
      $objWriter->save('php://output'); 
      exit;    
}  

  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>期間內傳真原因統計!! </p>
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee" width="60">日期區間:</td>
          <td> 
            <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> ~~
            <input name="edate" type="text" id="edate" size="12" maxlength="12" value=<?=$edate;?> onfocus="new WdatePicker()">   
          </td>
      </tr>       
       <tr>
          <td bgcolor="#eeeeee"></td>
          <td>                   
            <input type="submit" name="submit" id="submit" value="匯出">         
          </td>
       </tr>
    </table>
  </div>
</form> 