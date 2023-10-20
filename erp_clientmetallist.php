<?php
  session_start();
  $pagetitle = "帳單組 &raquo; Monthly Metal List";
  include("_data.php");
  include("_erp.php");
//  auth("erp_clientinvoicelist.php");
  
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
        $filename="templates/clientmetallist.xls";
        error_reporting(E_ALL);  
        
        require_once 'classes/PHPExcel.php'; 
        require_once 'classes/PHPExcel/IOFactory.php';  
        $objReader = PHPExcel_IOFactory::createReader('Excel5'); 
//        $objPHPExcel = $objReader->load("$filename");
        $objPHPExcel = new PHPExcel();
        // Set properties
                                                   
        $objPHPExcel ->getProperties()->setCreator( 'MetalList')
                     ->setLastModifiedBy('MetalList')
                     ->setTitle('MetalList')
                     ->setSubject('MetalList')
                     ->setDescription('MetalList')
                     ->setKeywords('MetalList')
                     ->setCategory('MetalList');
                     
        $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);   
        $edate1=substr($edate,2,2).substr($edate,5,2).substr($edate,8,2); 
        //檢查工單號有無出貨  , 配件不用   
        $s2="select tc_ofa01, to_char(tc_ofa02, 'yyyy-mm-dd') tc_ofa02, tc_ofb31, tc_ofb32, tc_ofb11, tc_ofb04, tc_ofb06, tc_ofbud02, tc_dex004, ogb31, ogb32, sfb01, sfb08
            from   vd210.tc_ofa_file, vd210.tc_ofb_file, vd210.tc_dex_file, vd210.ogb_file, vd110.sfb_file
            where tc_ofa01=tc_ofb01
            and  tc_ofa04='$occ01'
            and tc_ofa02 between to_date('$bdate1','yy/mm/dd') and to_date('$edate1','yy/mm/dd')
            and tc_ofb08='2'
            and tc_ofb31=tc_dex001 and tc_ofb04=tc_dex003
            and tc_dex001=ogb01 and tc_dex002=ogb03
            and ogb31=sfb22 and ogb32=sfb221
            order by tc_ofa02, tc_ofb11";
       

        $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
        $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);             
        $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Ariel');
        $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(9);
        $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);   
        //$objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 9); //每頁的1-5row都重複一樣                
        
        $objPHPExcel->setActiveSheetIndex(0);             
        $osheet = $objPHPExcel->getActiveSheet();
        $osheet ->setCellValue('A1', 'Date')
                ->setCellValue('B1', 'Ticket No.')
                ->setCellValue('C1', 'RX')
                ->setCellValue('D1', 'Metal')
                ->setCellValue('E1', 'Teeth No.')
                ->setCellValue('F1', 'Unit')
                ->setCellValue('G1', 'Weight');
        $y=2;
        $erp_sql2 = oci_parse($erp_conn2,$s2 );
        oci_execute($erp_sql2);         
        while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {           

          //印出資料
          $osheet ->setCellValue('A'. $y, $row2['TC_OFA02'])
                  ->setCellValue('B'. $y, $row2['SFB01'])
                  ->setCellValue('C'. $y, $row2["TC_OFB11"])
                  ->setCellValue('D'. $y, $row2["TC_OFB06"])
                  ->setCellValue('E'. $y, $row2["TC_OFBUD02"])
                  ->setCellValue('F'. $y, $row2["SFB08"])
                  ->setCellValue('G'. $y, number_format($row2["TC_DEX004"],2,'.',','));
          $osheet ->getStyle('F'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
          $osheet ->getStyle('G'.$y)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
          $y++;
        }
        
        // Rename sheet
        $osheet ->setTitle('Metal List');
        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $objPHPExcel->setActiveSheetIndex(0);
                                                            
        header('Content-Type: application/vnd.ms-excel');  
        header('Content-Disposition: attachment;filename="'. 'MetalList_' . $occ01 . '_' . $bdate . '.xls"');
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
        <th>Date</th>
        <th>Ticket No.</th>
        <th>Case No.</th>
        <th>Name</th>
        <th>Teeth No.</th>
        <th>Unit</th>
        <th>Weight</th>

    </tr>
    <?
      $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);   
      $edate1=substr($edate,2,2).substr($edate,5,2).substr($edate,8,2);

      $s2="select tc_ofa01, to_char(tc_ofa02, 'yyyy-mm-dd') tc_ofa02, tc_ofb31, tc_ofb32, tc_ofb11, tc_ofb04, tc_ofb06, tc_ofbud02, tc_dex004, ogb31, ogb32, sfb01, sfb08
            from   vd210.tc_ofa_file, vd210.tc_ofb_file, vd210.tc_dex_file, vd210.ogb_file, vd110.sfb_file
            where tc_ofa01=tc_ofb01
            and  tc_ofa04='$occ01'
            and tc_ofa02 between to_date('$bdate1','yy/mm/dd') and to_date('$edate1','yy/mm/dd')
            and tc_ofb08='2'
            and tc_ofb31=tc_dex001 and tc_ofb04=tc_dex003
            and tc_dex001=ogb01 and tc_dex002=ogb03
            and ogb31=sfb22 and ogb32=sfb221
            order by tc_ofa02, tc_ofb11";

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
              <td><?=$row2["SFB01"];?></td>
              <td><?=$row2["TC_OFB11"];?></td>  
              <td><?=$row2["TC_OFB06"];?></td>
              <td><?=$row2["TC_OFBUD02"];?></td>
              <td style="text-align:right" ><?=number_format($row2["SFB08"],0,'.',',');?></td>
              <td style="text-align:right" ><?=number_format($row2["TC_DEX004"],2,'.',',');?></td>
          </tr>
    <?    
      }
    ?>
</table>   