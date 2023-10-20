<?php
  session_start();
  $pagetitle = "業務部 &raquo; 期間內所有客戶delay合計";
  include("_data.php");
 // include("_erp.php");
 // auth("erp_clientdelayreport_all.php");  
  
  if (is_null($_GET['bdate'])) {
    $bdate =  date('Y').'-01-01';
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
  
    if ($_GET["submit"]=="匯出") {   
        error_reporting(E_ALL);  
        require_once 'classes/PHPExcel.php'; 
        require_once 'classes/PHPExcel/IOFactory.php';  
        $objPHPExcel = new PHPExcel();
        // Set properties
        $objPHPExcel ->getProperties()->setCreator( 'DelayReport')
                     ->setLastModifiedBy('DelayReport')
                     ->setTitle('DelayReport')
                     ->setSubject('DelayReport')
                     ->setDescription('DelayReport')
                     ->setKeywords('DelayReport')
                     ->setCategory('DelayReport');
        
        // Add some data      
                        
        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A1', $bdate .' -- ' . $edate . '  Delay Report');                    
        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A3', '客戶名稱.')  
                    ->setCellValue('B3', '客戶代號')                  
                    ->setCellValue('C3', '原始數據')
                    ->setCellValue('D3', '廠務數據'); 

        $y=4; 
      
      $query="  select b.clientid, b.clientname, nn, mm from 
                (select clientid, clientname, count(*) nn from
                (select clientid , clientname, rx  
                from delay 
                where tdate between '$bdate' and '$edate' 
                group by clientname, rx ) a
                group by clientname ) b , 

                (select clientid, clientname, count(*) mm from
                (select clientid, clientname, rx  
                from delay 
                where tdate between '$bdate' and '$edate' and status=''
                group by clientname, rx ) c
                group by clientname ) d 
                where b.clientname = d.clientname 
                order by clientname " ;  
      $result=mysql_query($query);
      while ($row = mysql_fetch_array($result)) {   
            $objPHPExcel->setActiveSheetIndex(0)           
                        ->setCellValue('A'. $y, $row["clientname"])   
                        ->setCellValue('B'. $y, $row["clientid"])   
                        ->setCellValue('C'. $y, $row["nn"])
                        ->setCellValue('D'. $y, $row["mm"]);
            $y++;  
        }
        //total
        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A'.($y+1), 'Total');  
                                                   
        $objPHPExcel->setActiveSheetIndex(0)        
                    ->setCellValue('C'.($y+1), '=sum(C3:C' . ($y-1). ")")                 
                    ->setCellValue('D'.($y+1), '=sum(D3:D' . ($y-1). ")");   
        
        // Rename sheet
        $objPHPExcel->getActiveSheet()->setTitle('DelayReport');
                                                                
        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $objPHPExcel->setActiveSheetIndex(0);

        // Redirect output to a client’s web browser (Excel5)
        header('Content-Type: application/vnd.ms-excel');  
        header('Content-Disposition: attachment;filename="'. $bdate .'--'. $edate . 'DelayReport.xls"');    
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
         計算日期:   
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()">  ~~
        <input name="edate" type="text" id="edate" size="12" maxlength="12" value=<?=$edate;?> onfocus="new WdatePicker()">  ~~     
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
        <th>客戶名稱</th>  
        <th>客戶代號</th>      
        <th>原始delay數</th> 
        <th>廠商delay數</th>       
    </tr>
    <?
      $query="  select b.clientid, b.clientname, nn, mm from 
                (select clientid, clientname, count(*) nn from
                (select clientid , clientname, rx  
                from delay 
                where tdate between '$bdate' and '$edate' 
                group by clientname, rx ) a
                group by clientname ) b , 

                (select clientid, clientname, count(*) mm from
                (select clientid, clientname, rx  
                from delay 
                where tdate between '$bdate' and '$edate' and status=''
                group by clientname, rx ) c
                group by clientname ) d 
                where b.clientname = d.clientname 
                order by clientname " ;      
      $total1=0;
      $total2=0;
      $result=mysql_query($query);
      while ($row = mysql_fetch_array($result)) {  
         $total1 += $row['nn'];  
         $total2 += $row['mm'];                    
          ?>
              <tr>  
                  <td><img src="i/arrow.gif" width="16" height="16"></td>    
                  <td><?=$row["clientname"];?></td> 
                  <td><?=$row["clientid"];?></td> 
                  <td><?=$row["nn"];?></td> 
                  <td><?=$row["mm"];?></td>   
             </tr>
          <?  
      }        
      ?>
      <tr>
        <td></td>      
        <td colspan="2">Total</td>    
        <td><?=$total1;?></td>            
        <td><?=$total2;?></td> 
      </tr>   
          
</table>   