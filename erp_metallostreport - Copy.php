<?php
  session_start();
  $pagetitle = "資材部 &raquo; 金屬損耗一覽表";
  include("_data.php");
  include("_erp.php");
  auth("erp_metallostreport.php");  
  
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
        error_reporting(E_ALL);  
        require_once 'classes/PHPExcel.php'; 
        require_once 'classes/PHPExcel/IOFactory.php';  
        $objPHPExcel = new PHPExcel();
        // Set properties
        $objPHPExcel ->getProperties()->setCreator( 'Frank')
                     ->setLastModifiedBy('Frank')
                     ->setTitle('金屬損耗統計表')
                     ->setSubject('金屬損耗統計表')
                     ->setDescription('金屬損耗統計表')
                     ->setKeywords('金屬損耗統計表')
                     ->setCategory('金屬損耗統計表');
        
        // Add some data      
                        
        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A1', '金屬損耗統計表');                    
        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A3', '序號')                   
                    ->setCellValue('B3', '日期')
                    ->setCellValue('C3', '物料代號')   
                    ->setCellValue('D3', '物料名稱')   
                    ->setCellValue('E3', '損耗量')  
                    ->setCellValue('F3', '筆數');  

        $y=4;      
        $i=0;   
        
        $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);  
        $edate1=substr($edate,2,2).substr($edate,5,2).substr($edate,8,2);  
        $s2= "select to_char(sfp02,'mm/dd/yy') sfp02, ima01, ima02, sum(sfpud11) sfpud11, count(*) qty from sfp_file , ( select sfe02, sfe07 from sfe_file group by sfe02, sfe07), ima_file 
            where sfpud11 >0 and sfp02>=to_date('$bdate1','yy/mm/dd') and sfp02<=to_date('$edate1','yy/mm/dd') and sfp01=sfe02 
            and sfe07=ima01 group by sfp02, ima01, ima02 order by sfp02, ima01";
        $erp_sql2 = oci_parse($erp_conn,$s2 );
        oci_execute($erp_sql2);  
        while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {                   
            $i++;
            $objPHPExcel->setActiveSheetIndex(0)
                        ->setCellValue('A'. $y, $i)  
                        ->setCellValue('B'. $y, $row2["SFP02"])    
                        ->setCellValue('C'. $y, $row2["IMA01"])   
                        ->setCellValue('D'. $y, $row2["IMA02"])   
                        ->setCellValue('E'. $y, $row2["SFPUD11"])  
                        ->setCellValue('F'. $y, $row2["QTY"]);
            $y++;    
        }                       
        
        // Rename sheet
        $objPHPExcel->getActiveSheet()->setTitle('金屬損耗統計表');
                                                                
        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $objPHPExcel->setActiveSheetIndex(0);

        // Redirect output to a client’s web browser (Excel5)
        header('Content-Type: application/vnd.ms-excel');  
        header('Content-Disposition: attachment;filename="MetalLostReport.xls"');    
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
<p>客戶Delay Report </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
         起訖日期:   
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> ~~ 
        <input name="edate" type="text" id="edate" size="12" maxlength="12" value=<?=$edate;?> onfocus="new WdatePicker()">    
        <input type="submit" name="submit" id="submit" value="查詢">  &nbsp;&nbsp;   &nbsp;&nbsp;      
        <input type="submit" name="submit" id="submit" value="匯出">  &nbsp;&nbsp;   &nbsp;&nbsp;            
      </td></tr>
    </table>
  </div>
</form>
<? if (is_null($_GET['submit'])) die ; ?>   
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
        <th width="16">&nbsp;</th> 
        <th>日期</th>         
        <th>物料代碼</th>  
        <th>物料名稱</th> 
        <th>損耗量</th>   
        <th>筆數</th>  
    </tr>
    <?
      $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);  
      $edate1=substr($edate,2,2).substr($edate,5,2).substr($edate,8,2);
      //檢查工單號有無出貨  , 配件不用   
      $s2= "select to_char(sfp02,'mm/dd/yy') sfp02, ima01, ima02, sum(sfpud11) sfpud11, count(*) qty from sfp_file , ( select sfe02, sfe07 from sfe_file group by sfe02, sfe07), ima_file 
            where sfpud11 >0 and sfp02>=to_date('$bdate1','yy/mm/dd') and sfp02<=to_date('$edate1','yy/mm/dd') and sfp01=sfe02 
            and sfe07=ima01 group by sfp02, ima01, ima02 order by sfp02, ima01";
      $erp_sql2 = oci_parse($erp_conn,$s2 );
      oci_execute($erp_sql2);  
      $bgkleur = "ffffff";  
      $i=0;            
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {    
          $i++;
          ?>
            <tr bgcolor="#<?=$bgkleur;?>">
                <td><?=$i;?></td>      
                <td><?=$row2["SFP02"];?></td>  
                <td><?=$row2["IMA01"];?></td>  
                <td><?=$row2["IMA02"];?></td>      
                <td><div align=right><?=number_format($row2["SFPUD11"], 2, ".", ",");?></div></td>      
                <td><div align=right><?=$row2["QTY"];?></div></td>
            </tr>
          <?   
      }
      ?>    
</table>   