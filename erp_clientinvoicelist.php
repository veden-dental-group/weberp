<?php
  // 目前 invoice ilst 只有兩個客戶會用: S4F 及 澳門  
  // 但 template 地址固定為S4F 因此 澳門客戶T183的地址要清為空白
  session_start();
  $pagetitle = "帳單組 &raquo; Monthly Invoice List";
  include("_data.php");
  include("_erp.php");
  auth("erp_clientinvoicelist.php");  
  
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
    $occ01 =  '';
  } else {
    $occ01=$_GET['occ01'];
  }  
  
    if (($_GET["submit"]=="Excel") or  ($_GET["submit"]=="PDF")) {   
        //取出客戶的currency
        $s2= "select occ02, occ42 from occ_file where occ01='$occ01'";
        $erp_sql2 = oci_parse($erp_conn2,$s2 );
        oci_execute($erp_sql2);  
        $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);
        $occ02=$row2['OCC02'];
        $currency=$row2['OCC42'];             
        $filename="templates/clientinvoicelist.xls";
        $currency=$row2['OCC42'];
        error_reporting(E_ALL);  
        
        require_once 'classes/PHPExcel.php'; 
        require_once 'classes/PHPExcel/IOFactory.php';  
        $objReader = PHPExcel_IOFactory::createReader('Excel5'); 
        $objPHPExcel = $objReader->load("$filename");  
        //$objPHPExcel = new PHPExcel();  
        // Set properties
                                                   
        $objPHPExcel ->getProperties()->setCreator( 'Invoicelist')
                     ->setLastModifiedBy('Invoicelist')
                     ->setTitle('Invoicelist')
                     ->setSubject('Invoicelist')
                     ->setDescription('Invoicelist')
                     ->setKeywords('Invoicelist')
                     ->setCategory('Invoicelist'); 
                     
        $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);   
        $edate1=substr($edate,2,2).substr($edate,5,2).substr($edate,8,2); 
        //檢查工單號有無出貨  , 配件不用   
        $s2="select tc_ofa01, to_char(tc_ofa02, 'yyyy-mm-dd') tc_ofa02, tc_ofb11, tc_ofb06, tc_ofb12, tc_ofb13, tc_ofb14, tc_ofbud01 " .
            "from tc_ofa_file, tc_ofb_file where tc_ofa01=tc_ofb01 and tc_ofa02 between to_date('$bdate1','yy/mm/dd') and to_date('$edate1','yy/mm/dd') and tc_ofa04='$occ01' " .
            "order by tc_ofa04,tc_ofa01, tc_ofb11 ";           //按送貨客戶, invoice No, Case No排序
       

        $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
        $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);             
        $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Ariel');
        $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(9);
        $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);   
        //$objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 9); //每頁的1-5row都重複一樣                
        
        $objPHPExcel->setActiveSheetIndex(0);             
        $osheet = $objPHPExcel->getActiveSheet();      
        if (substr($occ01,0,4)=='T183') {
            $osheet ->setCellValue('A5', '')
                    ->setCellValue('A6', '')  
                    ->setCellValue('A7', '')  
                    ->setCellValue('A8', '');  
        }
        $oldtcofa01='';  
        $subtotal=0;
        $total=0;
        $y=10;               
        $i=1;
        $erp_sql2 = oci_parse($erp_conn2,$s2 );
        oci_execute($erp_sql2);         
        while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {           
          if ($oldtcofa01!=$row2['TC_OFA01']) {
              if ($oldtcofa01!=''){        //oldtc_ofa01!='' 表示有資料 要印小計
                  $osheet ->setCellValue('E'. $y, "$currency")   
                          ->setCellValue('F'.$y, number_format($subtotal,2,'.',','));  
                  $osheet ->getStyle('F'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 ); 
                  $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);           
                  $osheet ->getStyle('F'.$y) ->getFont()->setBold(true);  
                  $osheet ->getStyle('F'.$y) ->getFont()->setSize(12); 
                  $osheet ->getStyle('F'.$y) ->getFont()->setSize(12); 
                  $osheet ->getStyle('E'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                  $osheet ->getStyle('F'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);  
                  $osheet ->getStyle('A'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
                  $osheet ->getStyle('B'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('C'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('D'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('E'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);                 
                  $osheet ->getStyle('F'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $y+=3;
              }
              
              $osheet -> setCellValue('A'.$y, "INVOICE NO.:" . $row2["TC_OFA01"])
                      -> setCellValue('F'.$y, "INVOICE DATE:". $row2["TC_OFA02"]);
              $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
              $osheet ->getStyle('F'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
              $osheet ->getStyle('A'.$y) ->getFont()->setSize(11); 
              $osheet ->getStyle('F'.$y) ->getFont()->setSize(11);         
              
              $y+=2;
              $osheet ->setCellValue('A'. $y, 'No.')
                      ->setCellValue('B'. $y, 'Case No.')   
                      ->setCellValue('C'. $y, 'Product Description')                                 
                      ->setCellValue('D'. $y, 'QTY(Unit)')   
                      ->setCellValue('E'. $y, "Unit Price($currency)")   
                      ->setCellValue('F'. $y, "Total($currency)")
                      ->setCellValue('G'. $y, "Remark");
              $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
              $osheet ->getStyle('B'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
              $osheet ->getStyle('C'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
              $osheet ->getStyle('D'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);  
              $osheet ->getStyle('E'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);     
              $osheet ->getStyle('F'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
              $osheet ->getStyle('G'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);  
              $osheet ->getStyle('A'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
              $osheet ->getStyle('B'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
              $osheet ->getStyle('C'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
              $osheet ->getStyle('D'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
              $osheet ->getStyle('E'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);                 
              $osheet ->getStyle('F'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
              $osheet ->getStyle('G'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);   
              $osheet ->getStyle('A'.$y) ->getFont()->setBold(true);  
              $osheet ->getStyle('B'.$y) ->getFont()->setBold(true);  
              $osheet ->getStyle('C'.$y) ->getFont()->setBold(true);  
              $osheet ->getStyle('D'.$y) ->getFont()->setBold(true);  
              $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);  
              $osheet ->getStyle('F'.$y) ->getFont()->setBold(true);  
              $osheet ->getStyle('G'.$y) ->getFont()->setBold(true);  
              $oldtcofa01=$row2['TC_OFA01'];
              $total+=$subtotal;
              $subtotal=0;
              $i=1;
              $y++;
          }
          //印出資料 
          $osheet ->setCellValue('A'. $y, $i)
                  ->setCellValue('B'. $y, $row2['TC_OFB11'])   
                  ->setCellValue('C'. $y, $row2["TC_OFB06"])
                  ->setCellValue('D'. $y, number_format($row2["TC_OFB12"],2,'.',','))  
                  ->setCellValue('E'. $y, number_format($row2["TC_OFB13"],2,'.',','))  
                  ->setCellValue('F'. $y, number_format($row2["TC_OFB14"],2,'.',','))
                  ->setCellValue('G'. $y, $row2["TC_OFBUD01"]);
          $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
          $osheet ->getStyle('B'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
          $osheet ->getStyle('C'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
          $osheet ->getStyle('D'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);  
          $osheet ->getStyle('E'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);     
          $osheet ->getStyle('F'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);  
          $osheet ->getStyle('G'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT); 
          
          //if (($y%2)==0){
          //    $osheet->getStyle('A'.$y.':F'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');    
          //}
          
          $osheet ->getStyle('A'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('A'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('A'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('A'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('B'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('B'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('B'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('B'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('C'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('C'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('C'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('C'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('D'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('D'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('D'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('D'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('E'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('E'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('E'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('E'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('F'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('F'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('F'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('F'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('G'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('G'.($y)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('G'.($y)) ->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
          $osheet ->getStyle('G'.($y)) ->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
          
          //$osheet ->getStyle('A'.($y)) ->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_BLUE);   
          //$osheet ->getStyle('B'.($y)) ->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_BLUE);  
          //$osheet ->getStyle('C'.($y)) ->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_BLUE);  
          //$osheet ->getStyle('D'.($y)) ->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_BLUE);  
          //$osheet ->getStyle('E'.($y)) ->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_BLUE);  
          //$osheet ->getStyle('F'.($y)) ->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_BLUE);                
          $y++;
          $i++; 
          $subtotal+=$row2['TC_OFB14'];
        }   
        
        //還有一個subtotal
        $osheet ->setCellValue('E'. $y, "$currency")   
                ->setCellValue('F'.$y, number_format($subtotal,2,'.',','));  
        $osheet ->getStyle('F'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 ); 
        $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);           
        $osheet ->getStyle('F'.$y) ->getFont()->setBold(true);  
        $osheet ->getStyle('F'.$y) ->getFont()->setSize(12); 
        $osheet ->getStyle('F'.$y) ->getFont()->setSize(12); 
        $osheet ->getStyle('E'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $osheet ->getStyle('F'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);                       
        $osheet ->getStyle('A'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
        $osheet ->getStyle('B'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
        $osheet ->getStyle('C'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
        $osheet ->getStyle('D'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
        $osheet ->getStyle('E'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);                 
        $osheet ->getStyle('F'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
        $osheet ->getStyle('G'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
        
        //Total
        $y+=3;
        $osheet ->setCellValue('E'. $y, "TOTAL:")   
                ->setCellValue('F'.$y, number_format($total,2,'.',','));  
        $osheet ->getStyle('F'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 ); 
        $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);           
        $osheet ->getStyle('F'.$y) ->getFont()->setBold(true);  
        $osheet ->getStyle('F'.$y) ->getFont()->setSize(12); 
        $osheet ->getStyle('F'.$y) ->getFont()->setSize(12); 
        $osheet ->getStyle('E'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        $osheet ->getStyle('F'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);  
        
        
        // Rename sheet
        $osheet ->setTitle('DailyInvoice');                
        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $objPHPExcel->setActiveSheetIndex(0);
                                                            
        header('Content-Type: application/vnd.ms-excel');  
        header('Content-Disposition: attachment;filename="'. 'DailyInvoice_' . $occ01 . '_' .$occ02 . '_' . $bdate . '.xls"');
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
<p>客戶 Invoice 列印 </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
         出貨日期:   
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> ~~ 
        <input name="edate" type="text" id="edate" size="12" maxlength="12" value=<?=$edate;?> onfocus="new WdatePicker()"> ~~  &nbsp;&nbsp; &nbsp;&nbsp; 
        帳款客戶: 
        <select name="occ01" id="occ01">  
            <?
              $s1= "select occ01,occ02 from occ_file order by occ01 ";
              $erp_sql1 = oci_parse($erp_conn1,$s1 );
              oci_execute($erp_sql1);  
              while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                  echo "<option value=" . $row1["OCC01"];  
                  if ($_GET["occ01"] == $row1["OCC01"]) echo " selected";                  
                  echo ">" . $row1['OCC01'] ."--" .$row1["OCC02"] . "</option>"; 
              }   
            ?>
        </select>  
        &nbsp;&nbsp;   &nbsp;&nbsp;  
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
        <th>Invoice No.</th>    
        <th>Invoice Date.</th>      
        <th>Case No.</th>     
        <th>Product Description</th>   
        <th>Qty.</th> 
        <th>Unit Price</th>
        <th>Total</th>    
        <th>Remark</th>   
        <th>Patient</th>
    </tr>
    <?
      $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);   
      $edate1=substr($edate,2,2).substr($edate,5,2).substr($edate,8,2); 
      //檢查工單號有無出貨  , 配件不用   
      $s2="select tc_ofa01, to_char(tc_ofa02, 'yyyy-mm-dd') tc_ofa02, tc_ofb11, tc_ofb06, tc_ofb12, tc_ofb13, tc_ofb14, tc_ofbud01, tc_ofbud05 " .
          "from tc_ofa_file, tc_ofb_file where tc_ofa01=tc_ofb01 and tc_ofa02 between to_date('$bdate1','yy/mm/dd') and to_date('$edate1','yy/mm/dd') and tc_ofa04='$occ01' " .
          "order by tc_ofa04,tc_ofa01, tc_ofb11 ";           //按送貨客戶, invoice No, Case No排序
      $erp_sql2 = oci_parse($erp_conn2,$s2 );
      oci_execute($erp_sql2);  
      $i=0;
      $total=0;
      $bgkleur = "ffffff";  
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {        
          $i++;
          $total+=$row2['TC_OFB14'];
    ?>
          <tr bgcolor="#<?=$bgkleur;?>">
              <td><?=$i;?></td>      
              <td><?=$row2["TC_OFA01"];?></td>  
              <td><?=$row2["TC_OFA02"];?></td>     
              <td><?=$row2["TC_OFB11"];?></td>  
              <td><?=$row2["TC_OFB06"];?></td>      
              <td style="text-align:right" ><?=number_format($row2["TC_OFB12"],2,'.',',');?></td>  
              <td style="text-align:right" ><?=number_format($row2["TC_OFB13"],2,'.',',');?></td>    
              <td style="text-align:right" ><?=number_format($row2["TC_OFB14"],2,'.',',');?></td>   
              <td><?=$row2["TC_OFBUD01"];?></td>
              <td><?=$row2["TC_OFBUD05"];?></td>
          </tr>
    <?    
      }
    ?>
    <tr bgcolor="#<?=$bgkleur;?>">   
        <td colspan="7"></td>    
        <td  style="text-align:right" ><?=number_format($total,2,'.',',');?></td>
        <td colspan="2"></td>
</table>   