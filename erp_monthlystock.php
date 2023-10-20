<?php
  session_start();
  $pagetitle = "資材部 &raquo; 每月月底結帳後庫存一覽表";
  include("_data.php");
  //auth("erp_monthlystock.php");  
  
  if (is_null($_GET['thismonth'])) {
    $thismonth = date("Y-m",strtotime("-1 month")) ;     
  } else {      
    $thismonth = $_GET['thismonth'] ;   
  }    
  $previousmonth=date('Y-m',strtotime("-1 month", strtotime($thismonth."-01")));
  
  if ($_GET['imd01']=='AAAA') {
      $imdfilter='';
  } else {
      $imdfilter= " and imk02='" . $_GET['imd01'] . "' ";
  }
  
  
  
  if ($_GET["submit"]=="匯出") {   

    $filename='templates/monthlystock.xls';
        
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
                
                                        
    $y=4;         
    $i=1;    
          
    $sql="select imk01 code, imk02, imk09 pimk09, 0 in1, 0 out2, 0 out3, 0 out4, 0 out5 , 0 timk09 from imk_file where (trim(to_char(imk05,'9999'))|| '-'||trim(to_char(imk06,'00')))='$previousmonth' " .
       "union all " .  
       "select   imk01 code, imk02, 0 pimk09, imk081 in1, imk082 out2, imk083 out3, imk084 out4, imk085 out5, imk09 timk09 from imk_file where (trim(to_char(imk05,'9999'))|| '-'||trim(to_char(imk06,'00')))='$thismonth' " ;
    $s2 ="select ima01, ima02, ima25, ima021, sum(pimk09) pimk09, sum(in1) in1, sum(out2) out2, sum(out3) out3, sum(out4) out4, sum(out5) out5, sum(timk09) timk09 from " .
       "ima_file,  ($sql) b where substr(ima06,1,1)!='9' $imdfilter and code=ima01 group by ima01,ima02,ima25, ima021 order by ima01,ima02  ";
    $erp_sql2 = oci_parse($erp_conn,$s2 );
    oci_execute($erp_sql2);   
    while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {     
          $objPHPExcel->setActiveSheetIndex(0)
                      ->setCellValue('A'. $y, $i)
                      ->setCellValue('B'. $y, $row2["IMA01"])
                      ->setCellValue('C'. $y, $row2["IMA02"])    
                      ->setCellValue('D'. $y, $row2["IMA25"]) 
                      ->setCellValue('E'. $y, $row2["IMA021"])  
                      ->setCellValue('F'. $y, $row2["PIMK09"])   
                      ->setCellValue('G'. $y, $row2["IN1"])   
                      ->setCellValue('H'. $y, $row2["OUT2"])   
                      ->setCellValue('I'. $y, $row2["OUT3"])   
                      ->setCellValue('J'. $y, $row2["OUT4"])  
                      ->setCellValue('K'. $y, $row2["OUT5"])  
                      ->setCellValue('L'. $y, $row2["TIMK09"]);   
          $y++;
          $i++;
    }                      
                                             
    // Rename sheet
    $objPHPExcel->getActiveSheet()->setTitle('每月庫存一覽表');
                                                            
    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
    $objPHPExcel->setActiveSheetIndex(0);

    // Redirect output to a client’s web browser (Excel5)
    header('Content-Type: application/vnd.ms-excel');  
    header('Content-Disposition: attachment;filename="monthlystock.xls"');   
    header('Cache-Control: max-age=0'); 
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');  
    $objWriter->save('php://output'); 
    exit;    
  }           
  
  
  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>以下為各倉庫的每月月底結帳後庫存一覽表!! </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
        月份:   
        <input name="thismonth" type="text" id="thismonth" onfocus="WdatePicker({dateFmt:'yyyy-MM'})" value="<?=$thismonth;?>"> &nbsp;&nbsp; 
        倉庫: 
        <select name="imd01" id="imd01">  
            <option value="AAAA">AAAA-不分倉庫</option>
            <?
              $s1= "select imd01, imd02 from imd_file order by imd01 ";
              $erp_sql1 = oci_parse($erp_conn,$s1 );
              oci_execute($erp_sql1);  
              while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                  echo "<option value=" . $row1["IMD01"];  
                  if ($_GET["imd01"] == $row1["IMD01"]) echo " selected";                  
                  echo ">" . $row1['IMD01'] ."--" .$row1["IMD02"] . "</option>"; 
              }   
            ?>
        </select> &nbsp;&nbsp;
        <input type="submit" name="submit" id="submit" value="查詢">  &nbsp;&nbsp;     
        <input type="submit" name="submit" id="submit" value="匯出">         
      </td></tr>
    </table>
  </div>
</form>  
          
<? if (is_null($_GET['submit'])) die ; ?>
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
        <th width="16">&nbsp;</th>
        <th>物料編號</th>  
        <th>物料名稱</th>  
        <th style="text-align:right">上月庫存</th>   
        <th style="text-align:right">本月入庫</th>   
        <th style="text-align:right">本月銷貨</th>   
        <th style="text-align:right">本月領用</th>   
        <th style="text-align:right">本月轉撥</th>   
        <th style="text-align:right">本月調整</th>        
        <th style="text-align:right">本月庫存</th>   
    </tr>
    <?
      
      $sql="select imk01 code, imk02 imk02, imk09 pimk09, 0 in1, 0 out2, 0 out3, 0 out4, 0 out5 , 0 timk09 from imk_file where (trim(to_char(imk05,'9999'))|| '-'||trim(to_char(imk06,'00')))='$previousmonth' " .
         "union all " .  
         "select imk01 code, imk02 imk02, 0 pimk09, imk081 in1, imk082 out2, imk083 out3, imk084 out4, imk085 out5, imk09 timk09 from imk_file where (trim(to_char(imk05,'9999'))|| '-'||trim(to_char(imk06,'00')))='$thismonth' " ;
      $s2 ="select ima01, ima02, sum(pimk09) pimk09, sum(in1) in1, sum(out2) out2, sum(out3) out3, sum(out4) out4, sum(out5) out5, sum(timk09) timk09 from " .
         "ima_file,  ($sql) b where substr(ima06,1,1)!='9' $imdfilter and code=ima01 group by ima01,ima02 order by ima01,ima02  ";
      $erp_sql2 = oci_parse($erp_conn,$s2 );
      oci_execute($erp_sql2);  
      $bgkleur = "ffffff";  
      $i=0;
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) { 
          $i++;
    ?>
	        <tr bgcolor="#<?=$bgkleur;?>">
		          <td><?=$i;?></td>
              <td><?=$row2["IMA01"];?></td>  
              <td><?=$row2["IMA02"];?></td>  
		          <td style="text-align:right"><?=number_format($row2["PIMK09"], 2);?></td>    
              <td style="text-align:right"><?=number_format($row2["IN1"], 2);?></td>       
              <td style="text-align:right"><?=number_format($row2["OUT2"], 2);?></td>    
              <td style="text-align:right"><?=number_format($row2["OUT3"], 2);?></td>    
              <td style="text-align:right"><?=number_format($row2["OUT4"], 2);?></td>    
              <td style="text-align:right"><?=number_format($row2["OUT5"], 2);?></td>      
              <td style="text-align:right"><?=number_format($row2["TIMK09"], 2);?></td>   
          </tr>
		  <?
			}
      ?>     
</table>    