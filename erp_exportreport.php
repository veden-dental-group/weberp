<?php
  //自MAI中取出數字匯出
  session_start();
  $pagetitle = "業務部 &raquo; Export Report";
  include("_data.php");
  include("_erp.php");
  //auth("erp_clientdelayreport.php");  

  
    if ($_GET["submit"]=="匯出") {   
        error_reporting(E_ALL);  
        require_once 'classes/PHPExcel.php'; 
        require_once 'classes/PHPExcel/IOFactory.php';  
        $objPHPExcel = new PHPExcel();
        // Set properties
        $objPHPExcel ->getProperties()->setCreator( 'FinanceReport')
                     ->setLastModifiedBy('FinanceReport')
                     ->setTitle('FinanceReport')
                     ->setSubject('FinanceReport')
                     ->setDescription('FinanceReport')
                     ->setKeywords('FinanceReport')
                     ->setCategory('FinanceReport');
        
        // Add some data      
                        
        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A1', 'SN')   
                    ->setCellValue('B1', 'Code')  
                    ->setCellValue('C1', 'Name')                  
                    ->setCellValue('D1', 'Amount') ;

         
      $query = "select * from financereport order by sn ";   
      $result = mysql_query($query) or die ('213 Users error!!');  
      $y=2;
      while ($row= mysql_fetch_array($result)) {
          //total
          $objPHPExcel->setActiveSheetIndex(0)
                      ->setCellValue('A'.$y, $row['sn'])  
                      ->setCellValue('B'.$y, $row['code']) 
                      ->setCellValue('C'.$y, $row['name']) 
                      ->setCellValue('D'.$y, $row['amount']); 
          $y++;
      }              
          
      // Rename sheet
      $objPHPExcel->getActiveSheet()->setTitle('DelayReport');
                                                              
      // Set active sheet index to the first sheet, so Excel opens this as the first sheet
      $objPHPExcel->setActiveSheetIndex(0);

      // Redirect output to a client’s web browser (Excel5)
      header('Content-Type: application/vnd.ms-excel');  
      header('Content-Disposition: attachment;filename="FinanceReport.xls"');    
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
<p>匯出財務 Report </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">                                                                                             
        報表: 
        <select name="mai01" id="mai01">  
            <?
              $s1= "select mai01, mai02 from mai_file order by mai01 ";
              $erp_sql1 = oci_parse($erp_conn7,$s1 );
              oci_execute($erp_sql1);  
              while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                  echo "<option value=" . $row1["MAI01"];  
                  if ($_GET["mai01"] == $row1["MAI01"]) echo " selected";                  
                  echo ">" . $row1['MAI01'] ."--" .$row1["MAI02"] . "</option>"; 
              }   
            ?>
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
        <th width="16">&nbsp;</th> 
        <th>順序</th>         
        <th>代碼</th> 
        <th>名稱</th>  
        <th>金額</th>         
    </tr>
    <?
      #先清空資料 再寫入 再顯示
      $query="delete from financereport ";
      $result = mysql_query($query) or die ('37 Users error!!');   
      $mai01=$_GET['mai01'];
      
      $s1="select * from maj_file where maj01='$mai01' order by maj02";
      $erp_sql7 = oci_parse($erp_conn7,$s1 );
      oci_execute($erp_sql7);  
      $i='1';
      while ($row7 = oci_fetch_array($erp_sql7, OCI_ASSOC)) {
          $maj02=str_pad($row7['MAJ02'],10,'0',STR_PAD_LEFT);
          $maj21=$row7['MAJ21'];
          $maj22=$row7['MAJ22'];    
          $maj20=$row7['MAJ20'];
          $s2= " SELECT SUM(aah04-aah05) amt FROM aah_file,aag_file WHERE aah00='07' AND aag00='07' " .
               " AND aah01 BETWEEN '$maj21' AND '$maj22' " .
               " AND aah02 = '2012' ".                   
               " AND aah03 BETWEEN '1' AND '12' ".
               " AND aah01 = aag01 ".
               " AND aag07 IN ('2','3')  ";
          $erp_sql2 = oci_parse($erp_conn7,$s2 );
          oci_execute($erp_sql2);  
          $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);
               
          $s3= " SELECT SUM(decode(abb06,'1',abb07, -abb07)) amt FROM abb_file,aba_file " .    
               " WHERE abb03 BETWEEN '$maj21' AND '$maj22' " .     
               " AND aba03 = '2012' ".   
               " AND aba04 BETWEEN '1' AND '12' ".
               " AND aba01 = abb01 ".
               " AND aba06 ='CE' ";                    
          $erp_sql3 = oci_parse($erp_conn7,$s3 );
          oci_execute($erp_sql3);  
          $row3 = oci_fetch_array($erp_sql3, OCI_ASSOC);
          
          $amt=$row2['AMT'] - $row3['AMT'];
          $queryus = "insert into financereport ( guid, sn, code, name, amount ) values (
                  '" . uuid()       . "',   
                   " . $i       . ",     
                  '" . $maj02   . "',
                  '" . $maj20   . "',
                  '" . $amt     . "' )";  
          if (is_null($maj20)) {
          } else { 
              $result = mysql_query($queryus) or die ('209 FinanceReport Add error!!'.mysql_error()); 
              $i++; 
          }
      }
      
      $query = "select * from financereport order by sn ";   
      $result = mysql_query($query) or die ('213 Users error!!');   
      $result=mysql_query($query);   
      while ($row= mysql_fetch_array($result)) {
          $bgkleur = "ffffff";
    ?>
          <tr bgcolor="#<?=$bgkleur;?>">
              <td><img src="i/arrow.gif" width="16" height="16"></td>
              <td><?=$row["sn"];?></td>  
              <td><?=$row["code"];?></td>  
              <td><?=$row["name"];?></td>
              <td><?=$row["amount"];?></td>       
          </tr>
      <?
      }
      ?>
          
</table>   