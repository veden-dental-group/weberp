<?php
  session_start();
  $pagetitle = "帳單組 &raquo; Monthly Statement V2";
  include("_data.php");
  include("_erp.php");
  auth("erp_clientmonthlystatement.php");  
  
  if (is_null($_GET['bmonth'])) {
    $bmonth =  date('Y-m');
  } else {
    $bmonth=$_GET['bmonth'];
  }  
  
  if (is_null($_GET['sdate'])) {
    $sdate =  date('Y-m-d');
  } else {
    $sdate=$_GET['sdate'];
  }                                                    
  
  if (is_null($_GET['occ01'])) {
    $occ01 =  '';
  } else {
    $occ01=$_GET['occ01'];
  }  
  
  if (($_GET["submit"]=="Excel") or  ($_GET["submit"]=="PDF")) {   
        //取出客戶基本資料:姓名 地址1, 地址2, 電話1,  幣別
        $socc= "select occ02, occ18, occ231, occ232, occ261, occ271, occ42, occud04 from occ_file where occ01='$occ01'";
        $erp_sqlocc = oci_parse($erp_conn2,$socc );
        oci_execute($erp_sqlocc);  
        $rowocc = oci_fetch_array($erp_sqlocc, OCI_ASSOC);  
        $occ02=$rowocc['OCC02'];
        if (is_null($rowocc['OCCUD04'])) {
            $occud04='N';                            
        } else {
            $occud04=$rowocc['OCCUD04'];
        }
        

        if($occ01=='E185001') {
            $filename = "templates/monthlystatementE185001.xls";
        } else {
            $filename="templates/monthlystatement".$occud04.".xls";
        }

        if ($rowocc['OCC42']=='USD') {
            $currency ='US$';
        } else if ($rowocc['OCC42']=='EUR'){
            $currency='€';
        } else if ($rowocc['OCC42']=='CAD'){
            $currency='CAD';
        } else {
            $currency=$rowocc['OCC42'].'$';
        }

        error_reporting(E_ALL);  
        
        require_once 'classes/PHPExcel.php'; 
        require_once 'classes/PHPExcel/IOFactory.php';  
        $objReader = PHPExcel_IOFactory::createReader('Excel5'); 
        $objPHPExcel = $objReader->load("$filename");  
        //$objPHPExcel = new PHPExcel();  
        // Set properties
                                                   
        $objPHPExcel ->getProperties()->setCreator( 'Monthly Statement')
                     ->setLastModifiedBy('Monthly Statement')
                     ->setTitle('Monthly Statement')
                     ->setSubject('Monthly Statement')
                     ->setDescription('Monthly Statement')
                     ->setKeywords('Monthly Statement')
                     ->setCategory('Monthly Statement');     
        
        $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
        $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);             
        $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Sylfaen');
        $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(11);
        $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);   
       // $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣                
        
        $objPHPExcel->setActiveSheetIndex(0);             
        $osheet = $objPHPExcel->getActiveSheet();  
        
        $sdate=date('M,d,Y', strtotime($sdate));

        $osheet -> setCellValue('A6', 'Bill to: (' . $rowocc['OCC02'] .')');

        if ($occud04=='1') {           //$occud04=1表示是DM 它的statement的表頭和別人不太一樣 ,
            $osheet -> setCellValue('C6', $sdate);                    
        } else {
            $osheet -> setCellValue('C5', $sdate);  
            $osheet -> setCellValue('A7', $rowocc["OCC18"]);
            $osheet -> setCellValue('A8', $rowocc["OCC231"]);   
            $osheet -> setCellValue('A9', $rowocc["OCC232"]);   
            $osheet -> setCellValue('A10', 'Tel:'. $rowocc["OCC261"] . '   Fax:'. $rowocc["OCC271"]);  
        }
        
        $osheet ->getStyle('A12') ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);
        $osheet ->getStyle('B12') ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);
        $osheet ->getStyle('C12') ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);
                                
        $byear1 =substr($bmonth,0,4);
        $bmonth1=substr($bmonth,5,2);  
        //$s2="select tc_ofd03, tc_ofdud03, to_char(tc_ofdud13,'Month,dd,yyyy') tc_ofdud13, tc_ofdud04, tc_ofd14, tc_ofc23 " .  
        $s2="select tc_ofd03, tc_ofdud03, to_char(tc_ofdud13, 'yyyy-mm-dd') tc_ofdud13, tc_ofdud04, tc_ofd14, tc_ofc23 " .  
            "from tc_ofd_file, tc_ofc_file where tc_ofc01=tc_ofd01 and tc_ofc021=$byear1 and tc_ofc022=$bmonth1 and tc_ofc04='$occ01' order by tc_ofd03 "; 
                                                                                
        $erp_sql2 = oci_parse($erp_conn2,$s2 );
        oci_execute($erp_sql2);  
        $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);     
                                                                                          
        $y=13;         
        $total=0;    
        $erp_sql2 = oci_parse($erp_conn2,$s2 );
        oci_execute($erp_sql2);               
        while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {  
          $total+=$row2["TC_OFD14"];            
          $sdate=date('F,d,y', strtotime($row2["TC_OFDUD13"])); 
          $osheet ->setCellValue('A'. $y, $row2["TC_OFDUD03"])                                 
                  //->setCellValue('B'. $y, $row2["TC_OFDUD13"])   
                  ->setCellValue('B'. $y, $sdate)   
                  ->setCellValue('C'. $y, number_format($row2["TC_OFD14"],2,'.',','))  ;
          $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
          $osheet ->getStyle('B'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); 
          $osheet ->getStyle('C'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT); 
//          $osheet ->getStyle('C'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
          //if (($y%2)==0){
          //    $osheet->getStyle('A'.$y.':C'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');    
          //}       
          //負數設成紅色字體
          if ($row2['TC_OFD14']<0) {
              $osheet ->getStyle('C'.$y) ->getFont()->setColor( new PHPExcel_Style_Color( PHPExcel_Style_Color::COLOR_RED));
          }                                    
          $y++;    
        }                                                                                                          
        $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);
        $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);
        $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);     
        //total
        $y++;

        $osheet ->setCellValue('A'.$y, 'Total:')
                ->setCellValue('B'.$y, $currency)
                ->setCellValue('C'.$y,  number_format($total,2,'.',','));

                
        $osheet ->getStyle('A'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT); 
        $osheet ->getStyle('B'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);  
        $osheet ->getStyle('C'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);   
        $osheet ->getStyle('C'.$y)->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);                    
       // $osheet ->getStyle('C'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 ); 
        
        $osheet ->getStyle('A'.$y) ->getFont()->setBold(true);  
        $osheet ->getStyle('B'.$y) ->getFont()->setBold(true);  
        $osheet ->getStyle('C'.$y) ->getFont()->setBold(true);      
        $osheet ->getStyle('C'.$y) ->getFont()->setUnderLine(true);     

        $y++;
        $y++;
        if ($occ01=='E185001') {
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'PLEASE  REMIT  THE  PAYMENT TO  OUR REPRESENTATIVE IN UK:');

            //$osheet ->getStyle('A'.$y) ->getFont()->setBold(true);
            $y++;
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'Payment Term : 30 days after monthly invoice');
            $y++;
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'Interest rate of  2% per month will be charged on all past due invoices after 30 days past due.');
            $y++;
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'Account Number : 068760039023');
            $y++;
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'Bank Name :  TAISHIN INTERNATIONAL BANK,OBU Branch');
            $y++;
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'Account Name : SOAR GAIN INVESTMENTS LIMITED');
            $y++;
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'Swift Code : TSIBTWTP019');
            $y++;
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'Bank Address : No.17, Sec.2, JIANGUO N. RD ZHONGSHAN DIST, Taipei City 104, Taiwan (R.O.C)');
            $y++;
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'Bank Telephone : +886-2-2568-3988');
            $y++;
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'Bank Fax : +886-2-2522-1490');
            $y++;
            $y++;
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'Sincerely:');
            $y++;
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'CK Change');
            $y++;
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'General Manager');
            $y++;
            $y++;
            $osheet->mergeCells('B' . $y . ':C' . $y);
            $osheet->setCellValue('B' . $y, 'Please send back to us after you sign it');
            $y++;
            $y++;
            $y++;
            $osheet->mergeCells('B' . $y . ':C' . $y);
            $osheet->setCellValue('B' . $y, 'CONFIRMED＆SIGNATURE:');
            $y++;
            $osheet->mergeCells('B' . $y . ':C' . $y);
            $osheet->setCellValue('B' . $y, $rowocc["OCC18"]);
        } else {
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'PLEASE  REMIT  THE  PAYMENT TO:');
            $y++;
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'Payment Term : 30 days after monthly invoice');
            $y++;
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'Interest rate of  2% per month will be charged on all past due invoices after 30 days past due.');
            $y++;
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'Account Number : 068730029160');
            $y++;
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'Bank Name :  TAISHIN INTERNATIONAL BANK,OBU Branch');
            $y++;
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'Account Name : V-Best Dental Technology Limited');
            $y++;
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'Swift Code : TSIBTWTP019');
            $y++;
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'Bank Address : No.17, Sec.2, JIANGUO N. RD ZHONGSHAN DIST, Taipei City 104, Taiwan (R.O.C)');
            $y++;
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'Bank Telephone : +886-2-2568-3988');
            $y++;
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'Bank Fax : +886-2-2522-1490');
            $y++;
            $y++;
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'Sincerely:');
            $y++;
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'CK Chang');
            $y++;
            $osheet->mergeCells('A' . $y . ':C' . $y);
            $osheet->setCellValue('A' . $y, 'General Manager');
            $y++;
            $y++;
            $osheet->mergeCells('B' . $y . ':C' . $y);
            $osheet->setCellValue('B' . $y, 'Please send back to us after you sign it');
            $y++;
            $y++;
            $y++;
            $osheet->mergeCells('B' . $y . ':C' . $y);
            $osheet->setCellValue('B' . $y, 'CONFIRMED＆SIGNATURE:');
            $y++;
            $osheet->mergeCells('B' . $y . ':C' . $y);
            $osheet->setCellValue('B' . $y, $rowocc["OCC18"]);
        }
             

        $osheet ->setTitle('Monthly Statement');                

        $objPHPExcel->setActiveSheetIndex(0);     
      
 
        header('Content-Type: application/vnd.ms-excel');  
        header('Content-Disposition: attachment;filename="'. 'MonthlyStatement_' . $occ01 . '_' .$occ02 .'_' . $bmonth . '.xls"');
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
         月份:   
        <input name="bmonth" type="text" id="bmonth" onfocus="WdatePicker({dateFmt:'yyyy-MM'})" value="<?=$bmonth;?>" size="8">  &nbsp;&nbsp;    
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
        Statement日期:   
        <input name="sdate" type="text" id="sdate" size="12" maxlength="12" value=<?=$sdate;?> onfocus="new WdatePicker()"> &nbsp;&nbsp;
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
        <th>Invoice Date</th>  
        <th>Amount</th>    
    </tr>
    <?
      $byear1 =substr($bmonth,0,4);
      $bmonth1=substr($bmonth,5,2);     
      //$s2="select tc_ofd03, tc_ofdud03, to_char(tc_ofdud13,'Month,dd,yyyy') tc_ofdud13, tc_ofdud04, tc_ofd14, tc_ofc23 " .  
      $s2="select tc_ofd03, tc_ofdud03, to_char(tc_ofdud13, 'yyyy-mm-dd') tc_ofdud13, tc_ofdud04, tc_ofd14, tc_ofc23 " .  
          "from tc_ofd_file, tc_ofc_file where tc_ofc01=tc_ofd01 and tc_ofc021=$byear1 and tc_ofc022=$bmonth1 and tc_ofc04='$occ01' order by tc_ofd03 "; 
      $erp_sql2 = oci_parse($erp_conn2,$s2 );
      oci_execute($erp_sql2);  
      $bgkleur = "ffffff";   
      $total=0;
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {

          if ($row2['TC_OFC23']=='USD') {
              $currency ='US$';
          } else if ($row2['TC_OFC23']=='EUR'){
              $currency='€';
          } else if ($row2['TC_OFC23']=='CAD'){
              $currency='CAD';
          } else {
              $currency=$row2['TC_OFC23'].'$';
          }

          $total+=$row2["TC_OFD14"];    
          $sdate=date('F,d,y', strtotime($row2["TC_OFDUD13"])); 
    ?>
          <tr bgcolor="#<?=$bgkleur;?>">
              <td><?=$row2["TC_OFD03"];?></td> 
              <td><?=$row2["TC_OFDUD03"];?></td>  
              <td><?=$sdate;?></td>    
              <td style="text-align:right"><?=number_format($row2["TC_OFD14"],2,'.',',');?></td>   
          </tr>
    <?    
      }
    ?>
    <tr bgcolor="#<?=$bgkleur;?>">   
        <td colspan="2">&nbsp;</td>    
        <td ><?=$currency;?></td>   
        <td style="text-align:right" ><?=number_format($total,2,'.',',');?></td>  
    </tr>   
</table>   