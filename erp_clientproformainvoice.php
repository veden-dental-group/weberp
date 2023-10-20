<?php
  session_start();
  $pagetitle = "業務部 &raquo; Proforma Invoice";
  include("_data.php");
  include("_erp.php");
  auth("erp_clientproformainvoice.php");  
    
  if (is_null($_GET['bdate'])) {
    $bdate =  date('Y-m-d');
  } else {
    $bdate=$_GET['bdate'];
  }       
  
  if (is_null($_GET['clienttype'])) {
    $clienttype='1';    
  } else {
    $clienttype= $_GET['clienttype'];
  }
  $month=substr($bdate,0,4).substr($bdate,5,2);                                      
  
  $occfilter='';
  $datefilter='';    

  $bocc01=$_GET['bocc01'];                                          
  if ($_GET['clienttype']=='1'){ 
      $occfilter  = " and oea04='$bocc01' ";    
  } else if ($_GET['clienttype']=='2'){ 
      $occfilter  = " and oea04 like 'U121%' "; 
  } else if ($_GET['clienttype']=='3'){     
      $occfilter  = " and oea04 like 'T183%' "; 
  } else {     
      $occfilter  = " and oea04 like 'H101%' ";     
  }
   
  $datefilter = " and oea02=to_date('$bdate','yy/mm/dd') ";
  
  if (($_GET["submit"]=="Excel") or  ($_GET["submit"]=="PDF")) {  
      //取出客戶的shipping advice格式設定
      //$s2= "select occud11 from occ_file where occ01='$bocc01'";
      //$erp_sql2 = oci_parse($erp_conn1,$s2 );
      //oci_execute($erp_sql2);  
      //$row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);   
      //$type=$row2['OCCUD11'];         
      $filename="templates/erp_clientproformainvoice.xls";   
      error_reporting(E_ALL);  
      
      require_once 'classes/PHPExcel.php'; 
      require_once 'classes/PHPExcel/IOFactory.php';  
      $objReader = PHPExcel_IOFactory::createReader('Excel5'); 
      $objPHPExcel = $objReader->load("$filename");  
                                                 
      $objPHPExcel ->getProperties()->setCreator( 'Proforma Invoice')
                   ->setLastModifiedBy('Proforma Invoice')
                   ->setTitle('Proforma Invoice')
                   ->setSubject('Proforma Invoice')
                   ->setDescription('Proforma Invoice')
                   ->setKeywords('Proforma Invoice')
                   ->setCategory('Proforma Invoice'); 
      $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
      $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);             
      $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
      $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
      $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);  
 
      //$objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 11); //每頁的1-5row都重複一樣                
      
      $objPHPExcel->setActiveSheetIndex(0);             
      $osheet = $objPHPExcel->getActiveSheet();  
      
      $socc="select occ02 from occ_file where occ01='$bocc01' ";
      $erp_sqlocc = oci_parse($erp_conn1,$socc );
      oci_execute($erp_sqlocc);               
      $rowocc = oci_fetch_array($erp_sqlocc, OCI_ASSOC);
      $osheet ->setCellValue('C2', $rowocc['OCC02'])
              ->setCellValue('C5', $bdate);
       
      $s2  ="select ta_oea006, ta_oea003, ta_oea001, ima1002, oeb04, oeb12, xmf07 from " .
            " (select ta_oea006, ta_oea003, ta_oea001, ima1002, oeb04, oeb12, occ44  from " . 
            "  oea_file, oeb_file, ima_file, occ_file   ".
            "  where oea01=oeb01 and oeb04=ima01 and oea04=occ01  $datefilter  $occfilter )" .  
            "left join xmf_file on occ44=xmf01 and oeb04=xmf03 " .              
            "order by ta_oea006  ";   
      $i=0;
      $y=8;
      $totalcase=0;
      $totalunit=0;
      $total=0;
      $oldcaseno='';
      $caseno=='';
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);               
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {   
        $totalunit+=$row2["OEB12"];
        $total+=$row2["OEB12"]*$row2['XMF07'];
        if ($row2['TA_OEA006']!= $oldcaseno) {  
              $i++;  
              $oldcaseno= $rowoga['TA_OEA006'];   
              $caseno   = $rowoga['TA_OEA006']; 
              $patient  = $rowoga['TA_OEA003'];    
              $clinic   = $rowoga['TA_OEA001']; 
              $ii=$i;   
          } else {            //重複的case no 不秀
              $caseno   = '';
              $patient  = '';
              $clinic   = '';
              $ii='';
          } 

  
        //要另外由vd210.ogb去取得售價  ogb14                                    
        $osheet ->setCellValue('A'. $y, $ii)    
                ->setCellValue('B'. $y, $caseno) 
                ->setCellValue('C'. $y, $patient)       
                ->setCellValue('D'. $y, $clinic)                                
                ->setCellValue('E'. $y, $row2["IMA1002"] )       
                ->setCellValue('F'. $y, $row2["OEB12"])
                ->setCellValue('G'. $y, $row2["XMF07"]) 
                ->setCellValue('H'. $y, $row2["OEB12"]*$row2['XMF07']) ;
        $osheet ->getStyle('B'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);              
        $osheet ->getStyle('G'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );  
        $osheet ->getStyle('H'.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 ); 
         
        if (($y%2)==0){
            $osheet->getStyle('A'.$y.':H'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');    
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
      //F秀合計
      $osheet ->setCellValue('B'. $y, "TOTAL:")
              ->setCellValue('C'. $y, $i)
              ->setCellValue('D'. $y, "CASES") 
              ->setCellValue('F'. $y, $totalunit)
              ->setCellValue('G'. $y, "UNITS")
              ->setCellValue('H'. $y, $total);                                                                        
      // Rename sheet
      $osheet ->setTitle('Performa Invoice');                
      // Set active sheet index to the first sheet, so Excel opens this as the first sheet
      $objPHPExcel->setActiveSheetIndex(0);
   
        header('Content-Type: application/vnd.ms-excel');  
        header('Content-Disposition: attachment;filename="Proforma_invoice_'. $bocc01 . '_'.$bdate.'.xls"');    
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
<p>客戶Proforma Invoice </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
         到貨日期:   
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()">  
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
        </select>   
        &nbsp;&nbsp;     
        客戶種類: 
            <input name="clienttype" type="radio" value="1" id="clienttype1" <?if($clienttype=='1') echo " checked";?>><label for="clienttype1">一般客戶 </label>&nbsp; 
            <input name="clienttype" type="radio" value="2" id="clienttype2" <?if($clienttype=='2') echo " checked";?>><label for="clienttype2">台安集團碼 </label>&nbsp; 
            <input name="clienttype" type="radio" value="3" id="clienttype3" <?if($clienttype=='3') echo " checked";?>><label for="clienttype3">澳門客戶群 </label>    
            <input name="clienttype" type="radio" value="4" id="clienttype4" <?if($clienttype=='4') echo " checked";?>><label for="clienttype4">HK客戶群 </label>      
        &nbsp;&nbsp;      
        <input type="submit" name="submit" id="submit" value="Submit">  &nbsp;&nbsp;   &nbsp;&nbsp;      
        <input type="submit" name="submit" id="submit" value="Excel">  &nbsp;&nbsp;   &nbsp; &nbsp;      
      </td></tr>
    </table>
  </div>
</form>
<? if (is_null($_GET['submit'])) die ; ?>   
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
        <th width="16">&nbsp;</th>    
        <th>RX No</th> 
        <th>Patient</th>
        <th>Clinic</th>       
        <th>Product</th>
        <th>Unit(s)</th>    
        <th>Price</th> 
        <th>Account</th>    
    </tr>
    <?    
      $soga="select ta_oea006, ta_oea003, ta_oea001, ima1002, oeb04, oeb12, xmf07 from " .
            " (select ta_oea006, ta_oea003, ta_oea001, ima1002, oeb04, oeb12, occ44  from " . 
            "  oea_file, oeb_file, ima_file, occ_file   ".
            "  where oea01=oeb01 and oeb04=ima01 and oea04=occ01  $datefilter  $occfilter )" .  
            "left join xmf_file on occ44=xmf01 and oeb04=xmf03 " .              
            "order by ta_oea006  ";          
 
      $erp_sqloga = oci_parse($erp_conn1,$soga );
      oci_execute($erp_sqloga);  
      $bgkleur = "ffffff";  
      $i=0;   
      $oldcaseno='';
      $caseno='';      
      $patient='';         
      while ($rowoga = oci_fetch_array($erp_sqloga, OCI_ASSOC)) { 
          if ($rowoga['TA_OEA006']!= $oldcaseno) {  
              $i++;  
              $oldcaseno= $rowoga['TA_OEA006'];   
              $caseno   = $rowoga['TA_OEA006']; 
              $patient  = $rowoga['TA_OEA003'];    
              $clinic   = $rowoga['TA_OEA001']; 
              $ii=$i;   
          } else {            //重複的case no 不秀
              $caseno   = '';
              $patient  = '';
              $clinic   = '';
              $ii='';
          }     
                
    ?>
          <tr bgcolor="#<?=$bgkleur;?>">
              <td><?=$ii;?></td>                 
              <td><?=$caseno;?></td> 
              <td><?=$patient;?></td>  
              <td><?=$clinic;?></td>   
              <td><?=$rowoga['IMA1002'];?></td>   
              <td><?=$rowoga['OEB12'];?></td> 
              <td><?=$rowoga['XMF07'];?></td> 
              <td style="text-align:right" ><?=$rowoga["OEB12"]*$rowoga['XMF07'];?></td>   
              <td>&nbsp;</td>       
          </tr>
    <?    
      }
    ?>           
</table>   