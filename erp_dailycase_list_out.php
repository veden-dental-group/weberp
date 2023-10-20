<?php
  session_start();
  $pagetitle = "客服部 &raquo; 客戶每日出貨清單";
  include("_data.php");
  //auth("erp_monthlystock.php");  

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
    $occ01=$_GET['occ01'];
  
  
  if ($_GET["submit"]=="匯出") {   

    $filename='templates/dailyout_list.xls';
        
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
                 
    $objPHPExcel->setActiveSheetIndex(0)
                ->setCellValue('B2', $imd01)
                ->setCellValue('J2', $thismonth);   
                
                                        
    $y=2;
    $i=1;    
          
    $sql="select to_char('yyyy/mm/dd', tc_oga002) tc_oga002, tc_ogb011 from tc_oga_file, tc_ogb_file " .
         "where tc_oga001=tc_ogb001 and tc_oga004='$occ01' order by tc_ogb011 ";
    $erp_sql2 = oci_parse($erp_conn,$s2 );
    oci_execute($erp_sql2);   
    while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {     
          $objPHPExcel->setActiveSheetIndex(0)
                      ->setCellValue('A'. $y, $i)
                      ->setCellValue('B'. $y, $row2["TC_OGA002"])
                      ->setCellValue('C'. $y, $row2["TC_OGB011"]);
          $y++;
          $i++;
    }                      
                                             
    // Rename sheet
    $objPHPExcel->getActiveSheet()->setTitle('每日出貨清單');
                                                            
    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
    $objPHPExcel->setActiveSheetIndex(0);

    // Redirect output to a client’s web browser (Excel5)
    header('Content-Type: application/vnd.ms-excel');  
    header('Content-Disposition: attachment;filename="dailyout_list.xls"');
    header('Cache-Control: max-age=0'); 
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');  
    $objWriter->save('php://output'); 
    exit;    
  }           
  
  
  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>每日客戶出貨清單!! </p>
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
    <div align="left">
        <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
            <tr><td bgcolor="#eeeeee" width="60">出貨區間:</td>

                <td>客戶:
                    <select name="occ01" id="occ01">
                      <?
                      $s1= "select occ01, occ02 from occ_file order by occ01 ";
                      $erp_sql1 = oci_parse($erp_conn,$s1 );
                      oci_execute($erp_sql1);
                      while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                        echo "<option value=" . $row1["OCC01"];
                        if ($_GET["occ01"] == $row1["OCC01"]) echo " selected";
                        echo ">" . $row1['OCC01'] ."--" .$row1["OCC02"] . "</option>";
                      }
                      ?>
                    </select>&nbsp;&nbsp;
                    <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> ~~
                    <input name="edate" type="text" id="edate" size="12" maxlength="12" value=<?=$edate;?> onfocus="new WdatePicker()">
                </td>
            </tr>
            <tr>
                <td bgcolor="#eeeeee"></td>
                <td>
                    <input type="submit" name="submit" id="submit" value="匯出">
                </td>
            </tr>
        </table>
    </div>
</form>