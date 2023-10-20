<?php
  session_start();
  $pagetitle = "採購部 &raquo; 物料最后採購日期";
  include("_data.php");
  //auth("erp_materiallatestorderdate.php");  
  
  $tdate=date('Y-m-d');
  
  if ($_GET["submit"]=="匯出") {   

    $filename='templates/erp_materiallatestorderdate.xls';
        
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
    $objPHPExcel->setActiveSheetIndex(0);             
    $osheet = $objPHPExcel->getActiveSheet();              
                 
    $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
    $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);             
    $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
    $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(12);
    $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);              

    $y=2;         
    $i=1;  
    
    $query2 = "select * from erp_materiallatestorderdate order by id";                                                                                        
    $result2 = mysql_query($query2) or die ('37 erp_materiallatestorderdate error!!' . mysql_error());    
    while ($row2= mysql_fetch_array($result2)) {      
          $objPHPExcel->setActiveSheetIndex(0)
                      ->setCellValue('A'. $y, $i)
                      ->setCellValue('B'. $y, $row2["id"] . ' ')
                      ->setCellValue('C'. $y, $row2["name"])
                      ->setCellValue('D'. $y, $row2["orderdate"])
                      ->setCellValue('E'. $y, $row2["price"])
                      ->setCellValue('F'. $y, $row2["qty"])
                      ->setCellValue('G'. $y, $row2["unit"]) 
                      ->setCellValue('H'. $y, $row2["supplierid"] .' ') 
                      ->setCellValue('I'. $y, $row2["suppliername"]) 
                      ->setCellValue('J'. $y, $row2["stockqty"]);
          $y++;
          $i++;
    }                           
           
    $objPHPExcel->setActiveSheetIndex(0)
                      ->setCellValue('A'. $y, '以下無資料');
     
    $objPHPExcel->getActiveSheet()->setTitle('各物料最后採購日期');                   
    $objPHPExcel->setActiveSheetIndex(0);                 
    header('Content-Type: application/vnd.ms-excel');  
    header('Content-Disposition: attachment;filename="'. $tdate . '_materiallatestorderdate.xls"');   
    header('Cache-Control: max-age=0'); 
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');  
    $objWriter->save('php://output'); 
    exit;    
  }   
  
  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>以下為各物料最后採購日期  </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">                                                                                                         
        <input type="submit" name="submit" id="submit" value="計算">  &nbsp;&nbsp;     
        <input type="submit" name="submit" id="submit" value="匯出">         
      </td></tr>
    </table>
  </div>
</form>  
          
<? if (is_null($_GET['submit'])) die ; ?>                                        
    <?
      //先清空暫存檔
      $query="delete from erp_materiallatestorderdate ";
      $result=mysql_query($query) or die ('153 MaterialLatestOrderDate delete error!!' . mysql_error());
    
      $s1 ="select ima01, ima02 from ima_file where ima08='P' order by ima01 ";     
      $erp_sql1 = oci_parse($erp_conn1,$s1 );
      oci_execute($erp_sql1);  
      $bgkleur = "ffffff";  
      $i=0;
      while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {  
          $id=$row1['IMA01'];
          $name=$row1['IMA02'];
          
          // 取出最后交易日及廠商
          $s2 ="select to_char(pmm04,'yyyy-mm-dd') pmm04, pmm09, pmc03, pmn07, pmn20, pmn31 from pmm_file, pmn_file, pmc_file where pmm01=pmn01 and pmm09=pmc01 and pmn04='$id' order by pmm04 desc  ";     
          $erp_sql2 = oci_parse($erp_conn1,$s2 );
          oci_execute($erp_sql2);
          $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);
          if (is_null($row2['PMM04'])) {
              $orderdate=NULL;
              $supplierid='';
              $suppliername=''; 
              $qty=0;
              $price=0;
              $unit='';    
          } else {
              $orderdate=$row2['PMM04'];
              $supplierid=$row2['PMM09'];
              $suppliername=$row2['PMC03'];
              $qty=$row2['PMN20'];
              $price=$row2['PMN31'];
              $unit=$row2['PMN07'];
          }
          
          //取出目前庫存 
          $s3 ="select sum(img10) img10 from img_file where img01='$id'  ";     
          $erp_sql3 = oci_parse($erp_conn1,$s3 );
          oci_execute($erp_sql3);
          $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);   
          if (is_null($row3['IMG10'])) {  
            $stockqty=0;
          } else {
            $stockqty=$row3['IMG10'];
          }
          
          $queryp = "insert into erp_materiallatestorderdate ( guid, tdate, id, name, orderdate, qty, price, unit, supplierid, suppliername, stockqty) values (       
                      '" . uuid()                   . "',
                      '" . $tdate                   .  "', 
                      '" . $id                      . "', 
                      '" . safetext($name)          . "',   
                      '" . $orderdate               . "',   
                      '" . $qty                     . "',   
                      '" . $price                   . "',   
                      '" . $unit                    . "',   
                      '" . $supplerid               . "',   
                      '" . safetext($suppliername)  . "',
                      '" . $stockqty                . "')"; 
            $resultp = mysql_query($queryp) or die ('204 erp_materiallatestorderdate add error!!'.mysql_error());
      }   
      msg('計算完畢, 請點擊匯出把資料匯出到Excel');
      ?>       