<?php
  // 目前 invoice ilst 只有兩個客戶會用: S4F 及 澳門  
  // 但 template 地址固定為S4F 因此 澳門客戶T183的地址要清為空白
  session_start();
  $pagetitle = "帳單組 &raquo; CDI Invoice";
  include("_data.php");
  include("_erp.php");
  //auth("erp_clientinvoice_cdi.php");  
  
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
    $occ01 =  'E113001';
  } else {
    $occ01=$_GET['occ01'];
  }  
  
    if (($_GET["submit"]=="Excel") or  ($_GET["submit"]=="PDF")) {   
        //取出客戶的currency
        $s2= "select occ42 from occ_file where occ01='$occ01'";
        $erp_sql2 = oci_parse($erp_conn2,$s2 );
        oci_execute($erp_sql2);  
        $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);   
        $currency=$row2['OCC42'];                     
        $currency=$row2['OCC42'];
        error_reporting(E_ALL);  
        
        require_once 'classes/PHPExcel.php'; 
        require_once 'classes/PHPExcel/IOFactory.php';  
        $objReader = PHPExcel_IOFactory::createReader('Excel5'); 
        //$objPHPExcel = $objReader->load("$filename");  
        $objPHPExcel = new PHPExcel();  
        // Set properties
                                                   
        $objPHPExcel ->getProperties()->setCreator( 'Invoicelist')
                     ->setLastModifiedBy('Invoicelist')
                     ->setTitle('Invoicelist')
                     ->setSubject('Invoicelist')
                     ->setDescription('Invoicelist')
                     ->setKeywords('Invoicelist')
                     ->setCategory('Invoicelist'); 
                                                                    
        $s2="select to_char(tc_ofa02, 'yyyy-mm-dd') tc_ofa02, tc_ofb11, tc_ofb01, tc_ofb03, tc_ofb06, tc_ofb12, tc_ofb13, tc_ofb14, tc_ofbud02, tc_ofbud01, ta_oea003, ta_oea046, ta_oea047, ta_oea048, ta_oea064, occud09 " .  
          "from tc_ofa_file, tc_ofb_file, oga_file, oea_file, occ_file where tc_ofa01=tc_ofb01 and tc_ofa02 between to_date('$bdate','yy/mm/dd') and to_date('$edate','yy/mm/dd') and tc_ofa04='$occ01' and oea04=occ01 " .
          "and tc_ofb31=oga01 and oga16=oea01 " .
          "order by tc_ofa02,tc_ofb11 ";
                                              

        //$objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 9); //每頁的1-5row都重複一樣       
             
        $s=0;            //切換sheet用
        $old_tcofb11='';      
        $erp_sql2 = oci_parse($erp_conn2,$s2 );
        oci_execute($erp_sql2);         
        while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {           
          if ($oldtcofb11!=$row2['TC_OFB11']) {
              if ($oldtcofb11 != ''){        //oldtc_ofa01!='' 表示有資料 要印小計    
                  $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
                  $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                  $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);                 
                  $osheet ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
                  $osheet ->setCellValue('D'. $y, "Total")   
                          ->setCellValue('F'. $y, '=sum(F11:F'.($y-1).")",'.',',');
                  $osheet ->getStyle('F'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 ); 
                  $osheet ->setTitle('O'.$oldtcofb11);  
                  $osheet ->getColumnDimension('A')->setAutoSize(true); 
                  $osheet ->getColumnDimension('B')->setAutoSize(true);
                  $osheet ->getColumnDimension('C')->setAutoSize(true); 
                  $osheet ->getColumnDimension('D')->setAutoSize(true); 
                  $osheet ->getColumnDimension('E')->setAutoSize(true); 
                  $osheet ->getColumnDimension('F')->setAutoSize(true); 
                  $s++;
              }
              
              $oldtcofb11=$row2['TC_OFB11'];
              if ($s>0) $objPHPExcel->createSheet($s);
              $objPHPExcel->setActiveSheetIndex($s);             
              $osheet = $objPHPExcel->getActiveSheet();
              $osheet->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
              $osheet->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);             
              $osheet->getDefaultStyle()->getFont()->setName('Ariel');
              $osheet->getDefaultStyle()->getFont()->setSize(10);
              $osheet->getDefaultRowDimension()->setRowHeight(15); 
              $osheet->getPageMargins()->setTop(0.9);  
              $osheet->getPageMargins()->setRight(0.4);
              $osheet->getPageMargins()->setLeft(0.4);
              $osheet->getPageMargins()->setBottom(0.9);
              
              $osheet -> setCellValue('A1', 'To:')
                      -> setCellValue('B1', 'CDI Dental AB.')  
                      -> setCellValue('E1', 'Date:')  
                      -> setCellValue('F1', $row2['TC_OFA02'])  
                      -> setCellValue('A2', 'Address:')  
                      -> setCellValue('B2', 'Center of Dental Innovation')  
                      -> setCellValue('B3', 'Solgatan 9-11')  
                      -> setCellValue('B4', '212 20 Malmo, Sweden.')  
                      -> setCellValue('A6', 'Invoice No.:')  
                      -> setCellValue('B6', $row2['TC_OFB01'] . ' ' . $row2['TC_OFB03'])  
                      -> setCellValue('A7', 'Case No.:')  
                      -> setCellValue('B7', 'O'.$row2['TC_OFB11'])      
                      -> setCellValue('A9', 'Patient:')  
                      -> setCellValue('B9', $row2['TA_OEA003'])
                      -> setCellValue('F10', '(USD)') 
                      -> setCellValue('A11', 'Description') 
                      -> setCellValue('C11', 'Teeth No.') 
                      -> setCellValue('D11', 'Qty') 
                      -> setCellValue('E11', 'Price') 
                      -> setCellValue('F11', 'Amount');
              
              $osheet ->getStyle('A11')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);   
              $osheet ->getStyle('C11')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
              $osheet ->getStyle('D11')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);  
              $osheet ->getStyle('E11')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);     
              $osheet ->getStyle('F11')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
              $osheet ->getStyle('A11') ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
              $osheet ->getStyle('B11') ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
              $osheet ->getStyle('C11') ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
              $osheet ->getStyle('D11') ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
              $osheet ->getStyle('E11') ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);                 
              $osheet ->getStyle('F11') ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                      
              //B8印條碼
              $osheet->getStyle('B8')->getFont()->setName('C39P36DlTt');
              $osheet->setCellValue('B8', '!O'.$row2['TC_OFB11'].'!');
              $osheet->getStyle('B8')->getFont()->setSize(14);    
              $osheet ->getStyle('B8')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);             
              $osheet->getDefaultStyle()->getFont()->setSize(10); 
              $osheet->getDefaultStyle()->getFont()->setName('Ariel');   
              $y=12;
          }
          
          //印出資料 
          $osheet ->setCellValue('A'. $y, $row2['TC_OFB06'])   
                  ->setCellValue('C'. $y, $row2["TC_OFBUD02"])
                  ->setCellValue('D'. $y, number_format($row2["TC_OFB12"],2,'.',','))  
                  ->setCellValue('E'. $y, number_format($row2["TC_OFB13"],2,'.',','))  
                  ->setCellValue('F'. $y, number_format($row2["TC_OFB14"],2,'.',','));                           
          $osheet ->getStyle('F'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 ); 
          $osheet ->mergeCells('A'.$y.':B'.$y);     
          $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
          $osheet ->getStyle('C'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
          $osheet ->getStyle('D'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);  
          $osheet ->getStyle('E'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);     
          $osheet ->getStyle('F'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);             
          //取出使用磁粉的名字及lotno
          if (is_null($row2[TA_OEA064])) {
              $model=$row2['OCCUD09']  ;
          } else {
              $model=$row2['TA_OEA064'];
          }  
          switch ($model){
              case '1':
                $modelcode='01';
                break;
              case '2':
                $modelcode='02';
                break;
              case '3':
                $modelcode='03';
                break;
              case '4':
                $modelcode='04';
                break;
              case '5':
                $modelcode='05';
                break;
              case '6':   
                $modelcode='52';
                break;
              default:        
                $modelcode='';
                break;   
          }
          
          $y++; 
          if (!is_null($row2['TA_OEA046'])) {
              $shade=$row2['TA_OEA046'];
              $q1="select code from erp_tc_imb1 where shade='$shade' limit 1 ";
              $r1=mysql_query($q1) or die ('191 erp_tc_imb1 error!!' . mysql_error());
              if (mysql_num_rows($r1)==1) {
                  $rr1=mysql_fetch_array($r1);
                  $shadecode=$rr1['code'];
              } else {
                  $shadecode='';
              }                
              $key= '1A' . $modelcode . 'B' . $shadecode ; 
              $q2="select name, lotno from erp_tc_imb2 where code='$key' limit 1";
              $r2=mysql_query($q2) or die ('200 erp_tc_imb2 error!!' . mysql_error());
              if (mysql_num_rows($r2)==1) {
                  $rr2=mysql_fetch_array($r2); 
                  $osheet ->setCellValue('A'. $y, '    '.$rr2['name'])   
                          ->setCellValue('C'. $y, $rr2["lotno"]);
                  $y++;                  
              }
          }    
          
          if (!is_null($row2['TA_OEA047'])) {
              $shade=$row2['TA_OEA047'];
              $q1="select code from erp_tc_imb1 where shade='$shade' limit 1 ";
              $r1=mysql_query($q1) or die ('191 erp_tc_imb1 error!!' . mysql_error());
              if (mysql_num_rows($r1)==1) {
                  $rr1=mysql_fetch_array($r1);
                  $shadecode=$rr1['code'];
              } else {
                  $shadecode='';
              }                
              $key= '1A' . $modelcode . 'B' . $shadecode ; 
              $q2="select name, lotno from erp_tc_imb2 where code='$key' limit 1";
              $r2=mysql_query($q2) or die ('200 erp_tc_imb2 error!!' . mysql_error());
              if (mysql_num_rows($r2)==1) {
                  $rr2=mysql_fetch_array($r2); 
                  $osheet ->setCellValue('A'. $y, '    '.$rr2['name'])   
                          ->setCellValue('C'. $y, $rr2["lotno"]);
                  $y++;                  
              }
          } 
          
          if (!is_null($row2['TA_OEA048'])) {
              $shade=$row2['TA_OEA048'];
              $q1="select code from erp_tc_imb1 where shade='$shade' limit 1 ";
              $r1=mysql_query($q1) or die ('191 erp_tc_imb1 error!!' . mysql_error());
              if (mysql_num_rows($r1)==1) {
                  $rr1=mysql_fetch_array($r1);
                  $shadecode=$rr1['code'];
              } else {
                  $shadecode='';
              }                
              $key= '1A' . $modelcode . 'B' . $shadecode ; 
              $q2="select name, lotno from erp_tc_imb2 where code='$key' order by ldate desc limit 1";
              $r2=mysql_query($q2) or die ('200 erp_tc_imb2 error!!' . mysql_error());
              if (mysql_num_rows($r2)==1) {
                  $rr2=mysql_fetch_array($r2); 
                  $osheet ->setCellValue('A'. $y, '    '.$rr2['name'])   
                          ->setCellValue('C'. $y, $rr2["lotno"]);
                  $y++;                  
              }
          }   
        }   
        
        $osheet ->getColumnDimension('A')->setWidth(14); 
        $osheet ->getColumnDimension('B')->setWidth(30);
        $osheet ->getColumnDimension('C')->setWidth(21);
        $osheet ->getColumnDimension('D')->setWidth(7); 
        $osheet ->getColumnDimension('E')->setWidth(7);
        $osheet ->getColumnDimension('F')->setWidth(12);
        //還有一個subtotal
        $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
        $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
        $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
        $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
        $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);                 
        $osheet ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
        $osheet ->setCellValue('D'. $y, "Total")   
                ->setCellValue('F'. $y, '=sum(F11:F'.($y-1).")",'.',',');
        $osheet ->getStyle('F'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 ); 
        $osheet ->setTitle('O'.$oldtcofb11);              
                
        // Rename sheet
                      
        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $objPHPExcel->setActiveSheetIndex(0);
                                                            
        header('Content-Type: application/vnd.ms-excel');  
        header('Content-Disposition: attachment;filename="'. 'DailyInvoice_' . $occ01 . '_' . $bdate . '.xls"');    
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
        送貨客戶: 
        <select name="occ01" id="occ01">  
            <?
              $s1= "select occ01,occ02 from occ_file order by occ01 ";
              $erp_sql1 = oci_parse($erp_conn1,$s1 );
              oci_execute($erp_sql1);  
              while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                  echo "<option value=" . $row1["OCC01"];  
                  if ($occ01 == $row1["OCC01"]) echo " selected";                  
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
        <th>Invoice Date.</th>  
        <th>Invoice No.</th>                
        <th>Case No.</th>     
        <th>Product Description</th>   
        <th>Teeth No.</th>       
        <th>Qty.</th> 
        <th>Unit Price</th>
        <th>Total</th>    
        <th>Patient</th>   
    </tr>
    <?                                                                
 
      $s2="select to_char(tc_ofa02, 'yyyy-mm-dd') tc_ofa02, tc_ofb11, tc_ofa01, tc_ofb03, tc_ofb06, tc_ofb12, tc_ofb13, tc_ofb14, tc_ofbud02, tc_ofbud01, ta_oea003 " .  
          "from tc_ofa_file, tc_ofb_file, oga_file, oea_file where tc_ofa01=tc_ofb01 and tc_ofa02 between to_date('$bdate','yy/mm/dd') and to_date('$edate','yy/mm/dd') and tc_ofa04='$occ01' " .
          "and tc_ofb31=oga01 and oga16=oea01 " . 
          "order by tc_ofa02,tc_ofb11 ";
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
              
              <td><?=$row2["TC_OFA02"];?></td>     
              <td><?=$row2["TC_OFB11"];?></td>  
              <td><?=$row2["TC_OFA01"]. ' ' . $row2['TCOFB03'];?></td>  
              <td><?=$row2["TC_OFB06"];?></td>      
              <td><?=$row2["TC_OFBUD02"];?></td> 
              <td style="text-align:right" ><?=number_format($row2["TC_OFB12"],2,'.',',');?></td>  
              <td style="text-align:right" ><?=number_format($row2["TC_OFB13"],2,'.',',');?></td>    
              <td style="text-align:right" ><?=number_format($row2["TC_OFB14"],2,'.',',');?></td>   
              <td><?=$row2["TA_OEA003"];?></td> 
          </tr>
    <?    
      }
    ?>
    <tr bgcolor="#<?=$bgkleur;?>">   
        <td colspan="7"></td>    
        <td  style="text-align:right" ><?=number_format($total,2,'.',',');?></td>   
    </tr>   
</table>   