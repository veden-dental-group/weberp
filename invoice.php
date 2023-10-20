<?php
  session_start();
  $pagetitle = "業務部 &raquo; Invoice";
  include("_data.php");
  //auth("invoice.php");  
  
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
  
  if ($_GET["submit"]=="匯出") {   
     
    //$s2= "select occ02 from occ_file where occ01='$occ01'";
    //$erp_sql2 = oci_parse($erp_conn,$s2 );
    //oci_execute($erp_sql2);  
    //$row = oci_fetch_array($erp_sql2, OCI_ASSOC);
    
    error_reporting(E_ALL);  
    require_once 'classes/PHPExcel.php'; 
    require_once 'classes/PHPExcel/IOFactory.php';  
    $objPHPExcel = new PHPExcel();
    // Set properties
    $objPHPExcel ->getProperties()->setCreator( $row["OCC01"] )
                 ->setLastModifiedBy($row["OCC01"])
                 ->setTitle($row["OCC01"])
                 ->setSubject($row["OCC01"])
                 ->setDescription($row["OCC01"])
                 ->setKeywords($row["OCC01"])
                 ->setCategory($row["OCC01"]);
    
    // Add some data      
                    
    $objPHPExcel->setActiveSheetIndex(0)
                ->setCellValue('F5', 'INVOICE'); 
    //$objPHPExcel->setActiveSheetIndex(0)
    //            ->setCellValue('A6', 'NO.:'.$row["OCC02"].substr($bdate,5,2).substr($bdate,8,2));
    $objPHPExcel->setActiveSheetIndex(0)
                ->setCellValue('H12', 'INVOICE DATE:'. date('Y/m/d'));             
    $objPHPExcel->setActiveSheetIndex(0)
                ->setCellValue('A13', 'No.')                   
                ->setCellValue('C13', 'Case No.')
                ->setCellValue('D13', 'Product Code')  
                ->setCellValue('F13', 'Product Description')  
                ->setCellValue('G13', 'Qty.')  
                ->setCellValue('H13', 'U.Price')  
                ->setCellValue('I13', 'Amount')  
                ->setCellValue('J13', 'Comment'); 

    $y=14; 
    $qtytotal=0;
    $amounttotal=0;       
    $i=1;   
    
    $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);
    $edate1=substr($edate,2,2).substr($edate,5,2).substr($edate,8,2);  
    $s2= "select to_char(oga02,'yyyy/mm/dd') oga02, oga04, occ02, ogaud02, ogb04, ogb12, ogb13, ogb12*ogb13 amount, ima1002 " .
         "from oga_file, ogb_file, ima_file, occ_file " .
         "where (oga02 between to_date('$bdate1','yymmdd') and to_date('$edate1','yymmdd') )and oga04='$occ01' and oga01=ogb01 and oga04=occ01 and ogb04=ima01 order by ogaud02";
    $erp_sql2 = oci_parse($erp_conn,$s2 );
    oci_execute($erp_sql2);   
    while ($row = oci_fetch_array($erp_sql2, OCI_ASSOC)) {     
          $objPHPExcel->setActiveSheetIndex(0)
                      ->setCellValue('A'. $y, $i)
                      ->setCellValue('B'. $y, $row["OGA02"])
                      ->setCellValue('C'. $y, $row["OGAUD02"])
                      ->setCellValue('D'. $y, $row["OGB04"])   
                      ->setCellValue('F'. $y, $row["IMA1002"])   
                      ->setCellValue('G'. $y, $row["OGB12"])   
                      ->setCellValue('H'. $y, $row["OGB13"])   
                      ->setCellValue('I'. $y, $row["AMOUNT"]);      
          $y++;
          $i++;
    }                                
 
    
    //total
    $objPHPExcel->setActiveSheetIndex(0)
                ->setCellValue('A'.($y+1), 'Total');  
                                               
    $objPHPExcel->setActiveSheetIndex(0)        
                ->setCellValue('G'.($y+1), '=sum(g14:g' . ($y-1). ")");                  
    $objPHPExcel->setActiveSheetIndex(0)        
                ->setCellValue('J'.($y+1), '=sum(j14:j' . ($y-1). ")");                  
    
    // Rename sheet
    $objPHPExcel->getActiveSheet()->setTitle('f');
                                                            
    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
    $objPHPExcel->setActiveSheetIndex(0);

    // Redirect output to a client’s web browser (Excel5)
    header('Content-Type: application/vnd.ms-excel');  
    header('Content-Disposition: attachment;filename="invoice.xls"');    
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
<p>以下為客戶 invoice 資料!! </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
        日期:   
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> ~~
        <input name="edate" type="text" id="edate" size="12" maxlength="12" value=<?=$edate;?> onfocus="new WdatePicker()">
        客戶: 
        <select name="occ01" id="occ01">  
            <?
              $s1= "select occ01,occ02 from occ_file order by occ01 ";
              $erp_sql1 = oci_parse($erp_conn,$s1 );
              oci_execute($erp_sql1);  
              while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                  echo "<option value=" . $row1["OCC01"];  
                  if ($_GET["occ01"] == $row1["OCC01"]) echo " selected";                  
                  echo ">" . $row1['OCC01'] ."--" .$row1["OCC02"] . "</option>"; 
              }   
            ?>
        </select>      
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
        <th>CASE NO.</th>  
        <th>Product Code</th> 
        <th>Product Description</th>   
        <th>Qty.</th>    
        <th>U.Price</th>  
        <th>Amount</th>   
    </tr>
    <?
      $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);
      $edate1=substr($edate,2,2).substr($edate,5,2).substr($edate,8,2);  
      $s2= "select to_char(oga02,'yyyy/mm/dd') oga02, oga04, occ02, ogaud02, ogb04, ogb12, ogb13, ogb12*ogb13 amount, ima1002 " .
           "from oga_file, ogb_file, ima_file, occ_file " .
           "where (oga02 between to_date('$bdate1','yymmdd') and to_date('$edate1','yymmdd') )and oga04='$occ01' and oga01=ogb01 and oga04=occ01 and ogb04=ima01 order by ogaud02";
      $erp_sql2 = oci_parse($erp_conn,$s2 );
      oci_execute($erp_sql2);  
      $qtytotal=0;
      $amounttotal=0;
      $bgkleur = "ffffff";  
      $i=0;
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) { 
          $qtytotal   +=$row2['OGB12'];
          $amounttotal+=$row2['AMOUNT']; 
          $i++;
    ?>
	        <tr bgcolor="#<?=$bgkleur;?>">
		          <td><?=$i;?></td>
              <td><?=$row2["OGA02"];?></td>  
              <td><?=$row2["OGAUD02"];?></td>  
		          <td><?=$row2["OGB04"];?></td>
              <td><?=$row2["IMA1002"];?></td>      
              <td style="text-align: right"><?=$row2["OGB12"];?></td>  
              <td style="text-align: right"><?=$row2["OGB13"];?></td>   
              <td style="text-align: right"><?=$row2["AMOUNT"];?></td>   
            </tr>
		  <?
			}
      ?>
      <tr bgcolor="#<?=$bgkleur;?>">
        <td><img src="i/arrow.gif" width="16" height="16"></td>
        <td colspan="4">Total</td>  
        <td style="text-align: right"><?=$qtytotal;?></td> 
        <td colspan="1">&nbsp;</td>  
        <td style="text-align: right"><?=$amounttotal;?></td>    
      </tr>  
</table>   
