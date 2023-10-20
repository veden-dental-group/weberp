<?php
  session_start();
  $pagetitle = "業務部 &raquo; 每日到貨檢查";
  include("_data.php");
  include("_erp.php");
  auth("erp_clientdailycaserx.php");  
  
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
     
  if (is_null($_GET['clienttype'])) {
    $clienttype='1';
  } else {
    $clienttype=$_GET['clienttype'];
  }
 
  $occfield='oea04';
  
  if ($clienttype=='1') {  
    $occfilter = " and oea04='$occ01' ";
  } else if ($clienttype=='2') {
    $occfilter = " and oea17='$occ01' ";  
  } else if ($clienttype=='3') {      
    $occfilter = " and oea04 like 'T183%' "; 
  } else {  
    $occfilter = " and oea04 like 'H101%' "; 
  }                          
  
    if ($_GET["submit"]=="匯出") {   
        error_reporting(E_ALL);  
        require_once 'classes/PHPExcel.php'; 
        require_once 'classes/PHPExcel/IOFactory.php';  
        $objPHPExcel = new PHPExcel();
        // Set properties
        $objPHPExcel ->getProperties()->setCreator( 'Frank')
                     ->setLastModifiedBy('Frank')
                     ->setTitle('Daily RX')
                     ->setSubject('Daily RX')
                     ->setDescription('Daily RX')
                     ->setKeywords('Daily RX')
                     ->setCategory('Daily RX');
        
        // Add some data      
                        
        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A1', 'Daily RX')  
                    ->setCellValue('B1', $bdate)     
                    ->setCellValue('C1', $occ01);    
                                       
        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A2', 'RX #')  
                    ->setCellValue('B2', 'RX #')   
                    ->setCellValue('C2', 'RX #')   
                    ->setCellValue('D2', 'RX #')   
                    ->setCellValue('E2', 'RX #');  

        $y=2; 
        $total=0;  
        $i=0;        
        $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);  
        //開始查詢當日到貨的CASE #  
        $s2= "select ta_oea006, count(*) howmany from oea_file where oea02=to_date('$bdate1','yy/mm/dd') $occfilter group by ta_oea006 order by ta_oea006";
        $erp_sql2 = oci_parse($erp_conn2,$s2 );
        oci_execute($erp_sql2);  
        while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {  
              if ($row2["HOWMANY"]>1){
              $howmany="(" . $row2["HOWMANY"] . ")";
          } else {
              $howmany=' ';
          }
              $total++; 
              if (($i%5)==0) {  
                  $y++;
                  $z=65;
              } else {          
                  $z++;
              } 
              $b=chr($z);
              
              $objPHPExcel->setActiveSheetIndex(0)   
                          ->setCellValue($b.$y, $row2['TA_OEA006'] . $howmany);    
              $i++;
        }
        //total
        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A'.($y+1), 'Total')  
                    ->setCellValue('B'.($y+1), $total.'組'); 
                                                                       
        // Rename sheet
        $objPHPExcel->getActiveSheet()->setTitle('Daily RX');
                                                                
        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $objPHPExcel->setActiveSheetIndex(0);

        // Redirect output to a client’s web browser (Excel5)
        header('Content-Type: application/vnd.ms-excel');  
        header('Content-Disposition: attachment;filename="Dailyrx.xls"');    
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
<p>客戶到貨RX No.檢查 </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
         到貨日期:   
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value="<?=$bdate;?>" onfocus="new WdatePicker()"> 
        客戶: 
        <select name="occ01" id="occ01">  
            <?
              $s1= "select occ01,occ02 from occ_file group by occ01, occ02 order by occ01 ";
              $erp_sql1 = oci_parse($erp_conn1,$s1 );
              oci_execute($erp_sql1);  
              while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                  echo "<option value=" . $row1["OCC01"];  
                  if ($_GET["occ01"] == $row1["OCC01"]) echo " selected";                  
                  echo ">" . $row1['OCC01'] ."--" .$row1["OCC02"] . "</option>"; 
              }   
            ?>
        </select>   &nbsp;&nbsp;   &nbsp;&nbsp;                                                                                                       
        客戶種類: 
        <input name="clienttype" type="radio" value="1" id="clienttype1" <?if($clienttype=='1') echo " checked";?>><label for="clienttype1">一般客戶 </label>&nbsp; 
        <input name="clienttype" type="radio" value="2" id="clienttype2" <?if($clienttype=='2') echo " checked";?>><label for="clienttype2">集團碼 </label>&nbsp; 
        <input name="clienttype" type="radio" value="3" id="clienttype3" <?if($clienttype=='3') echo " checked";?>><label for="clienttype3">澳門客戶群 </label>   
        <input name="clienttype" type="radio" value="4" id="clienttype3" <?if($clienttype=='4') echo " checked";?>><label for="clienttype4">HK客戶群 </label>  
            
        <input type="submit" name="submit" id="submit" value="查詢">  &nbsp;&nbsp;   &nbsp;&nbsp;      
        <input type="submit" name="submit" id="submit" value="匯出">  &nbsp;&nbsp;   &nbsp;&nbsp;            
      </td></tr>
    </table>
  </div>
</form>
<? if (is_null($_GET['submit'])) die ; ?>   
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr> 
        <th>RX #</th>         
        <th>RX #</th>    
        <th>RX #</th>  
        <th>RX #</th>  
        <th>RX #</th>  
    </tr>
    <?           
      $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);  
      //開始查詢當日到貨的CASE #  
      $s2= "select ta_oea006, count(*) howmany from oea_file where oea02=to_date('$bdate1','yy/mm/dd') $occfilter group by ta_oea006 order by ta_oea006";
      $erp_sql2 = oci_parse($erp_conn2,$s2 );
      oci_execute($erp_sql2);  
      $i=0;  
      $total=0;
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {  
          $total++;
          if ($row2["HOWMANY"]>1){
              $howmany="(" . $row2["HOWMANY"] . ")";
          } else {
              $howmany='';
          }
          if (($i%5)==0) {
              if ($i==0) {
                  echo "<tr>"; 
              } else {
                  echo "</tr>";
                  echo "<tr>";    
              }
          } 
          echo "<td>" . $row2["TA_OEA006"] . $howmany . "</td>";   
          $i++;
      } 
      echo "</tr>";
      ?>
      <tr>
        <td>合計</td>      
        <td><?=$total;?> 組</td>  
        <td colspan="3"></td>      
      </tr>   
          
</table>   
