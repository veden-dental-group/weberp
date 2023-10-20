<?php
  session_start();
  $pagetitle = "業務部 &raquo; 每日到貨檢查";
  include("_data.php");
  include("_erp.php");
  auth("erp_clientdailycaserxv2.php");  
  
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

  if (is_null($_GET['rec'])) {
    $rec = 5;
  } else {
    $rec=intval($_GET['rec']);
  }    
   
  
  if (is_null($_GET['occ20'])) {
    $occ20 =  '';
  } else {
    $occ20=$_GET['occ20'];
  }  

  if ($occ01=='all') {
      $occ01filter = '';
  } else {
      $occ01filter = " and occ01 = '$occ01' ";
  }

  if ($occ20=='all') {
      $occ20filter = '';
  } else {
      $occ20filter = " and substr(occ20,1,1) = '$occ20' ";
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
        
          
        $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);  
        //開始查詢當日到貨的CASE #  
        $s2= "select occ01, occ02, ta_oea006, count(*) howmany from oea_file, occ_file " .
             "where oea04=occ01 $occ01filter $occ20filter and oea02=to_date('$bdate1','yy/mm/dd') " .
             "group by occ01, occ02, ta_oea006 order by occ01, occ02, ta_oea006";
        $erp_sql2 = oci_parse($erp_conn2,$s2 );
        oci_execute($erp_sql2);  
        $occ01_old='';
        $y=1; 
        $total=0;
        $occtotal=0;  
        $i=0;     
        while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {  

          if ($occ01_old!=$row2['OCC01']) {
              if ($occ01_old != '') {
                  $y++;
                  $objPHPExcel->setActiveSheetIndex(0)
                              ->setCellValue('A'.$y, 'Total')  
                              ->setCellValue('B'.$y, $occtotal.'組'); 
                  $y++; 
                  $y++;
                  $i=0;   
                  $occtotal = 0;    
              }        

              $occ01_old = $row2['OCC01'];
              $objPHPExcel->setActiveSheetIndex(0)
                          ->setCellValue('A'.$y, 'Daily RX')  
                          ->setCellValue('B'.$y, $bdate)     
                          ->setCellValue('C'.$y, $row2['OCC02']);  
          }

          if ($row2["HOWMANY"]>1){
              $howmany="(" . $row2["HOWMANY"] . ")";
          } else {
              $howmany=' ';
          }

          $total++; 
          $occtotal++;

          if (($i%$rec)==0) {  
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
        $y++;
        $objPHPExcel->setActiveSheetIndex(0)
                              ->setCellValue('A'.$y, 'Total')  
                              ->setCellValue('B'.$y, $occtotal.'組'); 

        $y++;
        $y++;
        //total
        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A'.$y, 'Total')  
                    ->setCellValue('B'.$y, $total.'組'); 
                                                                       
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
        洲別: 
        <select name="occ20" id="occ20"> 
            <option value="all" <? if ($occ20=='all') echo " selected";?> >全部</option> 
            <option value="1" <? if ($occ20=='1') echo " selected";?> >亞洲</option>
            <option value="2" <? if ($occ20=='2') echo " selected";?> >美洲</option>
            <option value="3" <? if ($occ20=='3') echo " selected";?> >歐洲</option>
            <option value="4" <? if ($occ20=='4') echo " selected";?> >非洲</option>
            <option value="5" <? if ($occ20=='5') echo " selected";?> >澳洲</option>
        </select>   &nbsp;&nbsp;   &nbsp;&nbsp;                                                                                                       
        客戶: 
        <select name="occ01" id="occ01">  
            <option value="all" <? if ($occ01=='all') echo " selected";?> >全部</option> 
            <?
              $s1= "select occ01,occ02 from occ_file group by occ01, occ02 order by occ01 ";
              $erp_sql1 = oci_parse($erp_conn1,$s1 );
              oci_execute($erp_sql1);  
              while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                  echo "<option value=" . $row1["OCC01"];  
                  if ($occ01 == $row1["OCC01"]) echo " selected";                  
                  echo ">" . $row1['OCC01'] ."--" .$row1["OCC02"] . "</option>"; 
              }   
            ?>
        </select>   &nbsp;&nbsp;   &nbsp;&nbsp;  
        一行幾筆: 
        <select name="rec" id="rec"> 
            <option value="5" <? if ($rec=='5') echo " selected";?> >5</option>
            <option value="9" <? if ($rec=='9') echo " selected";?> >9</option>
        </select>   &nbsp;&nbsp;   &nbsp;&nbsp;    


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
      $s2= "select occ01, occ02, ta_oea006, count(*) howmany from oea_file, occ_file " .
           "where oea04=occ01 $occ01filter $occ20filter and oea02=to_date('$bdate1','yy/mm/dd') " .
           "group by occ01, occ02, ta_oea006 order by occ01, occ02, ta_oea006";
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
