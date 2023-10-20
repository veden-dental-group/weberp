<?php
  session_start();
  $pagetitle = "資材部 &raquo; 各單位耗材領用統計表";
  include("_data.php");
  include("_erp.php");
  //auth("erp_metaltakereport.php");  
  
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
  $key=$bdate.'--'.$edate; 
  
    if ($_GET["submit"]=="匯出") {   
        error_reporting(E_ALL);  
        require_once 'classes/PHPExcel.php'; 
        require_once 'classes/PHPExcel/IOFactory.php';  
        //$objPHPExcel = new PHPExcel();
        $objReader = PHPExcel_IOFactory::createReader('Excel5');  
        $objPHPExcel = $objReader->load("templates/metaltakereport.xls"); 
        // Set properties
        $objPHPExcel ->getProperties()->setCreator( 'Frank')
                     ->setLastModifiedBy('Frank')
                     ->setTitle('各單位耗材領用統計表')
                     ->setSubject('各單位耗材領用統計表')
                     ->setDescription('各單位耗材領用統計表')
                     ->setKeywords('各單位耗材領用統計表')
                     ->setCategory('各單位耗材領用統計表');
        
        // Add some data      
        $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
        $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);             
        $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Calibri');
        $objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setSize(10);
        $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);                
        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A1', $key . '  各單位耗材領用統計表');                    
        //$objPHPExcel->setActiveSheetIndex(0)
        //            ->setCellValue('A3', '序號') 
        //            ->setCellValue('B3', '製處代號') 
        //            ->setCellValue('C3', '製處名稱') 
        //            ->setCellValue('D3', '物料代號')   
        //            ->setCellValue('E3', '物料名稱')   
        //            ->setCellValue('F3', '領用量')  
        //            ->setCellValue('G3', '單價') 
        //            ->setCellValue('H3', '金額'); 

        $y=4;      
        $i=0;   
        
        $q3= "select tdate, mid, mname, pid, pname, spec, qty, unit, price, (qty*price) total from erp_metaltake where tdate='$key' order by mid, pid";
        $r3=mysql_query($q3) or die ('158 erp_metaltake error!!'.mysql_error());    
        $bgkleur = "ffffff";        
        $oldmid='';
        $oldy=4; //第一個開始的合計位罝
        while ($row3 = mysql_fetch_array($r3)) {                  
            if ($oldmid!=$row3['mid'] && $oldmid!='') {
                $objPHPExcel->getActiveSheet() ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $objPHPExcel->getActiveSheet() ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $objPHPExcel->getActiveSheet() ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
                $objPHPExcel->getActiveSheet() ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $objPHPExcel->getActiveSheet() ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
                $objPHPExcel->getActiveSheet() ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);    
                $objPHPExcel->getActiveSheet() ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
                $objPHPExcel->getActiveSheet() ->getStyle('H'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $objPHPExcel->getActiveSheet() ->getStyle('I'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
                $objPHPExcel->getActiveSheet() ->getStyle('J'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
                $objPHPExcel->setActiveSheetIndex(0)
                        ->setCellValue('B'. $y, '合計')
                        ->setCellValue('J'. $y, '=sum(J' . $oldy . ':J' . ($y-1) . ')');
                $y+=2;                //不同製處就跳兩行                
                $oldy=$y;
                $i=0;
            }  
            $oldmid=$row3['mid'];
            $i++;
            $objPHPExcel->setActiveSheetIndex(0)
                        ->setCellValue('A'. $y, $i)  
                        ->setCellValue('B'. $y, $row3["mid"])    
                        ->setCellValue('C'. $y, $row3["mname"])   
                        ->setCellValue('D'. $y, $row3["pid"])   
                        ->setCellValue('E'. $y, $row3["pname"])  
                        ->setCellValue('F'. $y, $row3["spec"])
                        ->setCellValue('H'. $y, $row3["unit"])
                        ->setCellValue('G'. $y, $row3["qty"])
                        ->setCellValue('I'. $y, $row3["price"])
                        ->setCellValue('J'. $y, $row3["total"]);
            if (($y%2)==0){
                $objPHPExcel->getActiveSheet()->getStyle('A'.$y.':J'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');
            } 
            $y++;    
        }                       
        $objPHPExcel->getActiveSheet() ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
        $objPHPExcel->getActiveSheet() ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
        $objPHPExcel->getActiveSheet() ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
        $objPHPExcel->getActiveSheet() ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
        $objPHPExcel->getActiveSheet() ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
        $objPHPExcel->getActiveSheet() ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);    
        $objPHPExcel->getActiveSheet() ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
        $objPHPExcel->getActiveSheet() ->getStyle('H'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
        $objPHPExcel->getActiveSheet() ->getStyle('I'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
        $objPHPExcel->getActiveSheet() ->getStyle('J'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
        $objPHPExcel->setActiveSheetIndex(0)
                        ->setCellValue('B'. $y, '合計')
                        ->setCellValue('J'. $y, '=sum(J' . $oldy . ':J' . ($y-1) . ')');
        // Rename sheet
        $objPHPExcel->getActiveSheet()->setTitle('各單位耗材領用統計表');
                                                                
        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $objPHPExcel->setActiveSheetIndex(0);

        // Redirect output to a client’s web browser (Excel5)
        header('Content-Type: application/vnd.ms-excel');  
        header('Content-Disposition: attachment;filename="'. $key . ' MetaltakeReport.xls"');    
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
  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>各單位耗材領用統計表 </p>     
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
        <th>期間</th>         
        <th>製處</th>  
        <th>製處</th>   
        <th>物料代碼</th> 
        <th>物料名稱</th> 
        <th>規格</th>
        <th>發料量</th>
        <th>單位</th>      
        <th>單價</th>  
        <th>金額</th>   
    </tr>
    <?
      //$bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);  
      //$edate1=substr($edate,2,2).substr($edate,5,2).substr($edate,8,2);
      //開始計算領料 並寫入table中
      
      //先刪除所有相同期間的資料
      
      $q1="delete from erp_metaltake where tdate='$key' ";
      $r1=mysql_query($q1);
      //
      $s2= "select ina04, gem02, inb04, ima02, ima021, inb08, sum(decode(ina00, '1', inb09, '3', (-1)*inb09, 0)) inb09 " .
           "from inb_file,ima_file, ina_file, gem_file " .
           "where inb04=ima01 and inb01=ina01 and ina04=gem01 " .
           "and ina02 between to_date('$bdate','yy/mm/dd') and to_date('$edate','yy/mm/dd') and inapost='Y' " .
           "group by ina04, gem02, inb04, ima02, ima021, inb08";
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);  
      $bgkleur = "ffffff"; 
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) { 
          //取出最新的一筆單價
          $mid=$row2['INB04'];
          $s3= "select ccc23a from ccc_file where ccc01='$mid' order by to_number(ccc02) desc ,to_number(ccc03) desc ";
          $erp_sql3 = oci_parse($erp_conn1,$s3 );
          oci_execute($erp_sql3);  
          $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);
          if (is_null($row3['CCC23A'])) {
              $ccc23a=0 ;
          } else {
              $ccc23a=$row3['CCC23A'] ;
          } 
          $query1= "insert into erp_metaltake ( tdate, mid, mname, pid, pname, spec, unit, qty, price ) values ( 
                   '" . $key                . "',  
                   '" . $row2['INA04']      . "', 
                   '" . $row2['GEM02']      . "',  
                   '" . $row2['INB04']      . "',  
                   '" . $row2['IMA02']      . "',  
                   '" . $row2['IMA021']      . "',  
                   '" . $row2['INB08']      . "', 
                    " . $row2['INB09']      . ", 
                    " . $ccc23a             . ")";
          $result1= mysql_query($query1) or die ('153 erp_metaltake Added error. == ' . $query1 . ' == ' .mysql_error());               
      
      }
      //
      $q3= "select tdate, mid, mname, pid, pname, spec, unit, qty, price, (qty*price) total from erp_metaltake where tdate='$key' order by mid, pid";
      $r3=mysql_query($q3) or die ('158 erp_metaltake error!!'.mysql_error());    
      $bgkleur = "ffffff";        
      while ($row3 = mysql_fetch_array($r3)) {   
          $i++;
          ?>
            <tr bgcolor="#<?=$bgkleur;?>">
                <td><?=$i;?></td>      
                <td><?=$row3["tdate"];?></td>  
                <td><?=$row3["mid"];?></td>  
                <td><?=$row3["mname"];?></td>   
                <td><?=$row3["pid"];?></td>  
                <td><?=$row3["pname"];?></td>      
                <td><?=$row3["spec"];?></td>
                <td><div align=right><?=number_format($row3["qty"], 2, ".", ",");?></div></td>
                <td><?=$row3["unit"];?></td>  
                <td><div align=right><?=number_format($row3["price"], 6, ".", ",");?></div></td> 
                <td><div align=right><?=number_format($row3["total"], 2, ".", ",");?></div></td>  
            </tr>
          <?   
      }
      ?>    
</table>   