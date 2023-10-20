<?php
  session_start();
  $pagetitle = "帳單組 &raquo; 每日invoice中的產品合計";
  include("_data.php");
  include("_erp.php");
  auth("erp_clientdailyinvoiceproducts.php");  
  
  if (is_null($_GET['bdate'])) {
    $bdate =  date('Y-m-d');
  } else {
    $bdate=$_GET['bdate'];
  }                                                   
  
  if (is_null($_GET['occ01'])) {
    $occ01 =  '';
  } else {
    $occ01=$_GET['occ01'];
  }    
  
    if (($_GET["submit"]=="Excel") or  ($_GET["submit"]=="PDF")) {   

        
        $filename="templates/productsummary.xls";
        $currency=$row2['OCC42'];
        error_reporting(E_ALL);  
        
        require_once 'classes/PHPExcel.php'; 
        require_once 'classes/PHPExcel/IOFactory.php';  
        $objReader = PHPExcel_IOFactory::createReader('Excel5'); 
        $objPHPExcel = $objReader->load("$filename");  
        //$objPHPExcel = new PHPExcel();  
        // Set properties
                                                   
        $objPHPExcel ->getProperties()->setCreator( 'Products List')
                     ->setLastModifiedBy('Products List')
                     ->setTitle('Products List')
                     ->setSubject('Products List')
                     ->setDescription('Products List')
                     ->setKeywords('Products List')
                     ->setCategory('Products List'); 
                     
        //$bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);  
        $s2="select imaud04, sum(tc_ofb12) tc_ofb12 from " .
              "(select decode(imaud04, NULL, ima01, imaud04) imaud04, tc_ofb12 from " .
                "(select imaud04, ima01, tc_ofb12 from " .
                  "(select tc_ofb04, sum(tc_ofb12) tc_ofb12 from tc_ofa_file, tc_ofb_file where tc_ofa02=to_date('$bdate','yy/mm/dd') and tc_ofa04='$occ01' and tc_ofa01=tc_ofb01 group by tc_ofb04), ima_file " .
                  "where tc_ofb04=ima01)) " . 
            " group by imaud04 order by imaud04 "; 
       
       //目前只有PLS使用本功能
       // if ($template=='3NN') { //3NN 指的是 PLS的invoice                   
                $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
                $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);             
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
                $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
                $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);   
                $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣              
                
                $objPHPExcel->setActiveSheetIndex(0);             
                $osheet = $objPHPExcel->getActiveSheet();  
                                                                                         
                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);  
                $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);                
                        
                $y=5;
                $osheet -> setCellValue('A'.$y, "DATE:".$bdate);
                        
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
                //$currency=$row2["TC_OFA23"];   
                $i=0;   
                $oldrxno=''; 
                $total=0;
               
                $erp_sql2 = oci_parse($erp_conn2,$s2 );
                oci_execute($erp_sql2);               
                while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {  
                  $i++;
                  //因取出的品代有為VD的 有的為PLS的 所以要判斷長度來取出品名
                  if (strlen($row2["IMAUD04"])==4) {
                      $s4="select ima1002, ima25 from ima_file where imaud04='" . $row2["IMAUD04"] . "'"; 
                      $erp_sql4 = oci_parse($erp_conn2,$s4 );
                      oci_execute($erp_sql4);  
                      $row4 = oci_fetch_array($erp_sql4, OCI_ASSOC);                     
                  } else {
                      $s4="select ima1002, ima25 from ima_file where ima01='" . $row2["IMAUD04"] . "'"; 
                      $erp_sql4 = oci_parse($erp_conn2,$s4 );
                      oci_execute($erp_sql4);  
                      $row4 = oci_fetch_array($erp_sql4, OCI_ASSOC);   
                  }
                  
                  if ($row4['IMA25']=="G") {
                      $unit="G";                     
                  } else  {
                      $unit="Unit(s)";
                      $unittotal+=$row2["TC_OFB12"];   
                  }      

                  // Case No. :           $rxno
                  // Patient Name:        TC_OFBUD05
                  // Product Description: TC_OFB06)                                
                  // Qty                  TC_OFB12
                  // Unit Price           TC_OFB13 
                  // Total                TC_OFB14
                  // Tooth Number         TC_OFBUD02
                  // memo                 TC_OFBUD01      
                              
                  $osheet ->setCellValue('A'. $y, $i)
                          ->setCellValue('B'. $y, $row2["IMAUD04"])  
                          ->setCellValue('C'. $y, $row4["IMA1002"])     
                          ->setCellValue('D'. $y, $row2["TC_OFB12"])   
                          ->setCellValue('E'. $y, $unit);    
                  $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);     
                  $osheet ->getStyle('D'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT); 
                  if (($y%2)==0){
                      $osheet->getStyle('A'.$y.':E'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');    
                  }        
                  
                  if (strlen($row2["IMAUD04"])!=4) {     
                      $osheet->getStyle('B'.$y )->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFF00');   
                  }  
                  
                  $y++;    
                }                                                                                                          
                $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
                $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);    
                //total
                $osheet ->setCellValue('A'.$y, 'Total') 
                        ->setCellValue('D'.$y, $unittotal)  
                        ->setCellValue('E'.$y, 'Units');    
                $osheet ->getStyle('D'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);                  
                $osheet ->getStyle('D'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 ); 
                
                $osheet ->getStyle('A'.$y) ->getFont()->setBold(true);  
                $osheet ->getStyle('B'.$y) ->getFont()->setBold(true);  
                $osheet ->getStyle('C'.$y) ->getFont()->setBold(true);  
                $osheet ->getStyle('D'.$y) ->getFont()->setBold(true);  
                $osheet ->getStyle('E'.$y) ->getFont()->setBold(true);    
                
                // Rename sheet
                $osheet ->setTitle('Daily Products Summary');                
                // Set active sheet index to the first sheet, so Excel opens this as the first sheet
                $objPHPExcel->setActiveSheetIndex(0);
       // } 
            header('Content-Type: application/vnd.ms-excel');  
            header('Content-Disposition: attachment;filename="'. 'DailyProduct_' . $occ01 . '_' . $bdate . '.xls"');    
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
<p>客戶每日產品種類合計清單</p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
         出貨日期:   
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> &nbsp;&nbsp;
        送貨客戶: 
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
        <th>Invoice No.</th>         
        <th>Case No.</th>  
        <th>Patient</th> 
        <th>Product Code</th>
        <th>Product Description</th>   
        <th>Qty.</th>    
        <th>Unit</th>
        <th>Price</th>
        <th>Total</th>
        <th>Currency</th>
        <th>Teeth No.</th>  
        <th>Memo</th>  
    </tr>
    <?
      $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);   
      //檢查工單號有無出貨  , 配件不用   
      $s2="select tc_ofa01, tc_ofa23, tc_ofb11, tc_ofbud05, tc_ofb08, tc_ofb04, tc_ofb06, tc_ofb12, tc_ofb05, tc_ofb13, tc_ofb14, tc_ofbud02, tc_ofbud01 " .  
          "from tc_ofa_file, tc_ofb_file where tc_ofa02=to_date('$bdate1','yy/mm/dd')  and tc_ofa04='$occ01' and tc_ofa01=tc_ofb01 order by tc_ofa01,tc_ofb11,tc_ofb08,tc_ofb04 "; 
      $erp_sql2 = oci_parse($erp_conn2,$s2 );
      oci_execute($erp_sql2);  
      $bgkleur = "ffffff";  
      $i=0;
      $oldrxno='';  
      $total=0;
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
          $total+=$row2["TC_OFB14"];   
          if ( $row2["TC_OFB11"]!=$oldrxno) {
            $i++;  
            $ii=$i;
            $rxno =$row2['TC_OFB11'];
            $oldrxno=$row2['TC_OFB11'];   
          } else {     
            $rxno='';
            $ii='';    
          }
              
          if ($row2['TC_OFB05']=="G") {
              $unit="G";                     
          } else  {
              $unit="Unit";
          }               
    ?>
          <tr bgcolor="#<?=$bgkleur;?>">
              <td><?=$ii;?></td>      
              <td><?=$row2["TC_OFA01"];?></td>  
              <td><?=$rxno;?></td>  
              <td><?=$row2["TC_OFBUD05"];?></td>  
              <td><?=$row2["TC_OFB04"];?></td>   
              <td><?=$row2["TC_OFB06"];?></td>    
              <td style="text-align:right" ><?=number_format($row2["TC_OFB12"],2,'.',',');?></td>    
              <td><?=$unit;?></td>    
              <td style="text-align:right" ><?=number_format($row2["TC_OFB13"],2,'.',',');?></td>    
              <td style="text-align:right" ><?=number_format($row2["TC_OFB14"],2,'.',',');?></td>
              <td><?=$row2["TC_OFA23"];?></td>        
              <td><?=$row2["TC_OFBUD02"];?></td>    
              <td><?=$row2["TC_OFBUD01"];?></td>  
          </tr>
    <?    
      }
    ?>
    <tr bgcolor="#<?=$bgkleur;?>">   
        <td colspan="9"></td>    
        <td  style="text-align:right" ><?=number_format($total,2,'.',',');?></td>  
        <td colspan="3">&nbsp;</td>  
    </tr>   
</table>   