<?php
  session_start();
  $pagetitle = "帳單組 &raquo; Cases Prices";
  include("_data.php");
  include("_erp.php");
  auth("erp_clientcaseprice.php");  
  
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
  
  $occfilter='';
  $imafilter='';   
  

  $bocc01=$_GET['bocc01'];
  $eocc01=$_GET['eocc01'];  
  $bima01=$_GET['bima01'];    
  $eima01=$_GET['eima01'];    
    
  $occfilter=" and tc_ofa04>='$bocc01' and tc_ofa04<='$eocc01' ";
  $imafilter=" and tc_ofb04>='$bima01' and tc_ofb04<='$eima01' ";   
  
    if (($_GET["submit"]=="Excel") or  ($_GET["submit"]=="PDF")) {        
        $filename="templates/clientcaseprice.xls";   
        error_reporting(E_ALL);  
        
        require_once 'classes/PHPExcel.php'; 
        require_once 'classes/PHPExcel/IOFactory.php';  
        $objReader = PHPExcel_IOFactory::createReader('Excel5'); 
        $objPHPExcel = $objReader->load("$filename");  
                                                   
        $objPHPExcel ->getProperties()->setCreator( 'Frank')
                     ->setLastModifiedBy('Frank')
                     ->setTitle('Frank')
                     ->setSubject('Frank')
                     ->setDescription('Frank')
                     ->setKeywords('Frank')
                     ->setCategory('InvFrankoice'); 
        $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
        $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);             
        $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
        $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
        $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);  
                     
      $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);   
      $edate1=substr($edate,2,2).substr($edate,5,2).substr($edate,8,2);   
       
      if ($_GET['isdetail']=='Y') {   //顯示每筆資料時 不用group
          $s2="select tc_ofa04, occ01, occ02, to_char(tc_ofa02,'yyyy-mm-dd') tc_ofa02, tc_ofa01, tc_ofb11, tc_ofb04, ima01, ima1002, tc_ofb12, tc_ofb13, tc_ofb14, tc_ofbud01, tc_ofbud02 " .
              "from tc_ofa_file, tc_ofb_file, occ_file, ima_file " .
              "where tc_ofa01=tc_ofb01 and tc_ofa04=occ01 and tc_ofb04=ima01 and tc_ofa02 between to_date('$bdate1','yy/mm/dd') and to_date('$edate1','yy/mm/dd') " .
              $occfilter . $imafilter . " order by occ02, ima1002, tc_ofa04, tc_ofa02, tc_ofb04, tc_ofb13  ";          
      } else {
          $s2="select * from ( " .    //把相同價錢合成一筆時 不秀日期 invoice # 但要合計顆數     
                  "select tc_ofa04, occ01, occ02, tc_ofb04, ima01, ima1002, ' ' tc_ofa02, ' ' tc_ofa01, ' ' tc_ofb11, tc_ofb13, '' tc_ofbud01, '' tc_ofbud02, sum(tc_ofb12) tc_ofb12, sum(tc_ofb14) tc_ofb14  from ( " .
                      "select tc_ofa04, occ01, occ02, tc_ofb04, ima01, ima1002, ' ' tc_ofa02, ' ' tc_ofa01, ' ' tc_ofb11, tc_ofb12 , tc_ofb13, tc_ofb14, ''  tc_ofbud01 " .
                      "from tc_ofa_file, tc_ofb_file, occ_file, ima_file " .
                      "where tc_ofa01=tc_ofb01 and tc_ofa04=occ01 and tc_ofb04=ima01 and tc_ofa02 between to_date('$bdate1','yy/mm/dd') and to_date('$edate1','yy/mm/dd') " .
                  $occfilter . $imafilter .
                  " )  group by tc_ofa04, occ01, occ02, tc_ofb04, ima01, ima1002, tc_ofa02, tc_ofa01, tc_ofb11,  tc_ofb13, tc_ofbud01 " .
              " ) order by tc_ofa04, tc_ofb04, tc_ofb13  "; 
      }   
       
        
      $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
      $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);             
      $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
      $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
      $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);   
      //$objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣                
      
      $objPHPExcel->setActiveSheetIndex(0);             
      $osheet = $objPHPExcel->getActiveSheet();  
                                                      
      $i=0;
      $y=2;
      $erp_sql2 = oci_parse($erp_conn2,$s2 );
      oci_execute($erp_sql2);               
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {          
        $i++;                                      
        $osheet ->setCellValue('A'. $y, $row2["OCC01"])
                ->setCellValue('B'. $y, $row2["OCC02"])
                ->setCellValue('C'. $y, $row2["TC_OFA02"])
                ->setCellValue('D'. $y, $row2["IMA01"])
                ->setCellValue('E'. $y, $row2["IMA1002"])
                ->setCellValue('F'. $y, $row2["TC_OFA01"])
                ->setCellValue('G'. $y, $row2["TC_OFB11"])
                ->setCellValue('H'. $y, number_format($row2["TC_OFB12"],2,'.',','))
                ->setCellValue('I'. $y, number_format($row2["TC_OFB13"],2,'.',','))
                ->setCellValue('J'. $y, number_format($row2["TC_OFB14"],2,'.',','))
                ->setCellValue('K'. $y, $row2["TC_OFBUD01"])
                ->setCellValue('L'. $y, $row2["TC_OFBUD02"]);
        $osheet ->getStyle('H'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
        $osheet ->getStyle('I'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
        $osheet ->getStyle('J'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
        if (($y%2)==0){
            $osheet->getStyle('A'.$y.':L'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
        } 
        $y++;
      }        
      
      $osheet ->getStyle('A'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
      $osheet ->getStyle('B'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
      $osheet ->getStyle('C'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
      $osheet ->getStyle('D'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
      $osheet ->getStyle('E'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);                 
      $osheet ->getStyle('F'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
      $osheet ->getStyle('G'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
      $osheet ->getStyle('H'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
      $osheet ->getStyle('I'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
      $osheet ->getStyle('J'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
      $osheet ->getStyle('K'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
      $osheet ->getStyle('L'.($y)) ->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
      //F秀合計
      $osheet ->setCellValue('G'. $y, "合計:")
              ->setCellValue('H'. $y, "=sum(H2:H" . ($y-1) . ")")
              ->setCellValue('J'. $y, "=sum(J2:J" . ($y-1) . ")");
      $osheet ->getStyle('H'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
      $osheet ->getStyle('J'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );
      // Rename sheet
      $osheet ->setTitle('Client Cases Prices');                
      // Set active sheet index to the first sheet, so Excel opens this as the first sheet
      $objPHPExcel->setActiveSheetIndex(0);
   
        header('Content-Type: application/vnd.ms-excel');  
        header('Content-Disposition: attachment;filename="clientcaseprice.xls"');    
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
<p>客戶期間內各產品價格列印 </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
         出貨日期:   
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()">  ~~ 
        <input name="edate" type="text" id="edate" size="12" maxlength="12" value=<?=$edate;?> onfocus="new WdatePicker()">  &nbsp;&nbsp;   
        客戶: 
        <select name="bocc01" id="bocc01">  
            <?
              $s1= "select occ01,occ02 from occ_file order by occ01 ";
              $erp_sql1 = oci_parse($erp_conn2,$s1 );
              oci_execute($erp_sql1);  
              while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {                     
                  echo "<option value=" . $row1["OCC01"];  
                  if ($_GET["bocc01"] == $row1["OCC01"]) echo " selected";                  
                  echo ">" . $row1['OCC01'] ."--" .$row1["OCC02"] . "</option>"; 
              }   
            ?>
        </select>  ~~ 
        <select name="eocc01" id="eocc01">  
            <?
              $s1= "select occ01,occ02 from occ_file order by occ01 ";
              $erp_sql1 = oci_parse($erp_conn2,$s1 );
              oci_execute($erp_sql1);  
              while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {                     
                  echo "<option value=" . $row1["OCC01"];  
                  if ($_GET["eocc01"] == $row1["OCC01"]) echo " selected";                  
                  echo ">" . $row1['OCC01'] ."--" .$row1["OCC02"] . "</option>"; 
              }   
            ?>
        </select> 
        &nbsp;&nbsp;   
        產品: 
        <select name="bima01" id="bima01">      
            <?
              $s1= "select ima01,ima02 from ima_file where substr(ima06,1,1)='9' order by ima01 ";
              $erp_sql1 = oci_parse($erp_conn2,$s1 );
              oci_execute($erp_sql1);  
              while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {                     
                  echo "<option value=" . $row1["IMA01"];  
                  if ($_GET["ima01"] == $row1["IMA01"]) echo " selected";                  
                  echo ">" . $row1['IMA01'] ."--" .$row1["IMA02"] . "</option>"; 
              }   
            ?>
        </select> ~~  
        <select name="eima01" id="eima01">      
            <?
              $s1= "select ima01,ima02 from ima_file where substr(ima06,1,1)='9' order by ima01 ";
              $erp_sql1 = oci_parse($erp_conn2,$s1 );
              oci_execute($erp_sql1);  
              while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {                     
                  echo "<option value=" . $row1["IMA01"];  
                  if ($_GET["eima01"] == $row1["IMA01"]) echo " selected";                  
                  echo ">" . $row1['IMA01'] ."--" .$row1["IMA02"] . "</option>"; 
              }   
            ?>
        </select> 
        &nbsp;&nbsp;  
        <input type="checkbox" name="isdetail" id="isdetail" value='Y' <? if ($_GET['isdetail']=='Y') echo " checked";?>>列印細項
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
        <th>客戶</th>         
        <th>日期</th>  
        <th>品名</th> 
        <th>Invoice </th>   
        <th>Case No.</th>           
        <th>顆數</th>    
        <th>單價</th>
        <th>總價</th>     
        <th>齒位</th>
        <th>備註</th>
    </tr>
    <?
      $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);   
      $edate1=substr($edate,2,2).substr($edate,5,2).substr($edate,8,2);   
       
      if ($_GET['isdetail']=='Y') {   //顯示每筆資料時 不用group
          $s2="select tc_ofa04, occ02, to_char(tc_ofa02,'yyyy-mm-dd') tc_ofa02, tc_ofa01, tc_ofb11, tc_ofb04, ima1002, tc_ofb12, tc_ofb13, tc_ofb14, tc_ofbud01, tc_ofbud02 " .
              "from tc_ofa_file, tc_ofb_file, occ_file, ima_file " .
              "where tc_ofa01=tc_ofb01 and tc_ofa04=occ01 and tc_ofb04=ima01 and tc_ofa02 between to_date('$bdate1','yy/mm/dd') and to_date('$edate1','yy/mm/dd') " .
              $occfilter . $imafilter . " order by tc_ofa04, tc_ofa02, tc_ofb04, tc_ofb13  ";          
      } else {
          $s2="select * from ( " .    //把相同價錢合成一筆時 不秀日期 invoice # 但要合計顆數     
                  "select tc_ofa04, occ02, tc_ofb04, ima1002, ' ' tc_ofa02, ' ' tc_ofa01, ' ' tc_ofb11, tc_ofb13, '' tc_ofbud01, sum(tc_ofb12) tc_ofb12, sum(tc_ofb14) tc_ofb14  from ( " .             
                      "select tc_ofa04, occ02, tc_ofb04, ima1002, ' ' tc_ofa02, ' ' tc_ofa01, ' ' tc_ofb11, tc_ofb12 , tc_ofb13, tc_ofb14, ''  tc_ofbud01, '' tc_ofbud02 " .
                      "from tc_ofa_file, tc_ofb_file, occ_file, ima_file " .
                      "where tc_ofa01=tc_ofb01 and tc_ofa04=occ01 and tc_ofb04=ima01 and tc_ofa02 between to_date('$bdate1','yy/mm/dd') and to_date('$edate1','yy/mm/dd') " .
                  $occfilter . $imafilter .
                  " )  group by tc_ofa04, occ02, tc_ofb04, ima1002, tc_ofa02, tc_ofa01, tc_ofb11,  tc_ofb13, tc_ofbud01 " .
              " ) order by tc_ofa04, tc_ofb04, tc_ofb13  "; 
      }    
      $erp_sql2 = oci_parse($erp_conn2,$s2 );
      oci_execute($erp_sql2);  
      $bgkleur = "ffffff";  
      $i=0;                  
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {  
        $i++;        
    ?>
          <tr bgcolor="#<?=$bgkleur;?>">
              <td><?=$ii;?></td>      
              <td><?=$row2["TC_OFA04"].'--'.$row2['OCC02'];?></td>   
              <td><?=$row2["TC_OFA02"];?></td>
              <td><?=$row2["TC_OFB04"].'--'.$row2['IMA1002'];?></td>      
              <td><?=$row2["TC_OFA01"];?></td>  
              <td><?=$row2["TC_OFB11"];?></td>        
              <td style="text-align:right" ><?=number_format($row2["TC_OFB12"],2,'.',',');?></td>    
              <td style="text-align:right" ><?=number_format($row2["TC_OFB13"],2,'.',',');?></td>    
              <td style="text-align:right" ><?=number_format($row2["TC_OFB14"],2,'.',',');?></td>   
              <td><?=$row2["TC_OFBUD02"];?></td>
              <td><?=$row2["TC_OFBUD01"];?></td>
          </tr>
    <?    
      }
    ?>           
</table>   