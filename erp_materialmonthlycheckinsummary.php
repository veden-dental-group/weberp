<?php
  session_start();
  $pagetitle = "資材部 &raquo; 期間內物料入庫月統計表";
  include("_data.php");
  //auth("erp_materialmonthlycheckinsummary.php");  
  
  if (is_null($_GET['thismonth'])) {
    $thismonth = date("Y-m",strtotime("-1 month")) ;     
  } else {      
    $thismonth = $_GET['thismonth'] ;   
  }                                                                            
  
  $imd01=$_GET['imd01']; 
  
    //往前算出12個月的月份
  $montha=array();
  $montha[1]=$thismonth;       
  $yy=intval(substr($thismonth,0,4));
  $mm=intval(substr($thismonth,5,2));
  for ($x=2; $x<13; $x++) {
      $mm--;
      if ($mm==0) {
          $mm=12;
          $yy--;  
      } 
      $montha[$x]=strval($yy) . '-' . str_pad($mm,2,'0',STR_PAD_LEFT); 
  }
  
  
  if ($_GET["submit"]=="匯出") {   

    $filename='templates/materialyearlycheckinsummary.xls';
        
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
                 
    $objPHPExcel->setActiveSheetIndex(0)
                ->setCellValue('B2', $imd01) ; 
    $objPHPExcel->setActiveSheetIndex(0)
                ->setCellValue('F3', $montha[1])
                ->setCellValue('G3', $montha[2]) 
                ->setCellValue('H3', $montha[3]) 
                ->setCellValue('I3', $montha[4]) 
                ->setCellValue('J3', $montha[5]) 
                ->setCellValue('K3', $montha[6]) 
                ->setCellValue('L3', $montha[7]) 
                ->setCellValue('M3', $montha[8]) 
                ->setCellValue('N3', $montha[9]) 
                ->setCellValue('O3', $montha[10]) 
                ->setCellValue('P3', $montha[11]) 
                ->setCellValue('Q3', $montha[12]); 
    $y=4;         
    $i=1;  
    
    $s2 ="select ima01, ima02, ima021, ima25, month01, month02, month03, month04, month05, month06, month07, month08, month09, month10, month11, month12 from " .
           "ima_file left join (select rvv31, sum(month01) month01, sum(month02) month02, sum(month03) month03, sum(month04) month04, sum(month05) month05, sum(month06) month06, sum(month07) month07, sum(month08) month08, sum(month09) month09,sum(month10) month10,sum(month11) month11, sum(month12) month12 ".
           "from (select rvv31, decode(rvu03,'$montha[1]',rvvud07, 0) month01, decode(rvu03,'$montha[2]',rvvud07, 0)  month02, decode(rvu03,'$montha[3]',rvvud07, 0)  month03, decode(rvu03,'$montha[4]',rvvud07, 0) month04, " .
           "                    decode(rvu03,'$montha[5]',rvvud07, 0) month05, decode(rvu03,'$montha[6]',rvvud07, 0)  month06, decode(rvu03,'$montha[7]',rvvud07, 0)  month07, decode(rvu03,'$montha[8]',rvvud07, 0) month08, " .
           "                    decode(rvu03,'$montha[9]',rvvud07, 0) month09, decode(rvu03,'$montha[10]',rvvud07, 0) month10, decode(rvu03,'$montha[11]',rvvud07, 0) month11, decode(rvu03,'$montha[12]',rvvud07, 0) month12 " . 
           "       from (select to_char(rvu03, 'yyyy-mm') rvu03, rvv31 , rvvud07 from rvv_file, rvu_file where  rvv01=rvu01 and rvu00='1') )  group by rvv31 ) " .
           "on rvv31=ima01 where substr(ima06,1,1)!='9' and ima35='$imd01' order by ima01 " ;            //  rvv32='$imd01' and
           //and (month01+month02+month03+month04+month05+month06+month07+month08+month09+month10+month10+month11+month12)>0  ";  

    $erp_sql2 = oci_parse($erp_conn,$s2 );
    oci_execute($erp_sql2);   
    while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {     
          $objPHPExcel->setActiveSheetIndex(0)
                      ->setCellValue('A'. $y, $i)
                      ->setCellValue('B'. $y, $row2["IMA01"])
                      ->setCellValue('C'. $y, $row2["IMA02"])
                      ->setCellValue('D'. $y, $row2["IMA021"])
                      ->setCellValue('E'. $y, $row2["IMA25"])
                      ->setCellValue('F'. $y, $row2["MONTH01"])
                      ->setCellValue('G'. $y, $row2["MONTH02"]) 
                      ->setCellValue('H'. $y, $row2["MONTH03"]) 
                      ->setCellValue('I'. $y, $row2["MONTH04"]) 
                      ->setCellValue('J'. $y, $row2["MONTH05"]) 
                      ->setCellValue('K'. $y, $row2["MONTH06"]) 
                      ->setCellValue('L'. $y, $row2["MONTH07"]) 
                      ->setCellValue('M'. $y, $row2["MONTH08"]) 
                      ->setCellValue('N'. $y, $row2["MONTH09"]) 
                      ->setCellValue('O'. $y, $row2["MONTH10"]) 
                      ->setCellValue('P'. $y, $row2["MONTH11"]) 
                      ->setCellValue('Q'. $y, $row2["MONTH12"])    
                      ->setCellValue('R'. $y, $row2["MONTH01"]+$row2["MONTH02"]+$row2["MONTH03"]+$row2["MONTH04"]+$row2["MONTH05"]+$row2["MONTH06"]+$row2["MONTH07"]+$row2["MONTH08"]+$row2["MONTH09"]+$row2["MONTH10"]+$row2["MONTH11"]+$row2["MONTH12"]);       
          if (($y%2)==0){
              $osheet->getStyle('A'.$y.':R'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');    
          }  
          $y++;
          $i++;
    }
    $osheet ->getStyle('A'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
    $osheet ->getStyle('B'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
    $osheet ->getStyle('C'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
    $osheet ->getStyle('D'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
    $osheet ->getStyle('E'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
    $osheet ->getStyle('F'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
    $osheet ->getStyle('G'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('H'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
    $osheet ->getStyle('I'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
    $osheet ->getStyle('J'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
    $osheet ->getStyle('K'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
    $osheet ->getStyle('L'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
    $osheet ->getStyle('M'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
    $osheet ->getStyle('N'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
    $osheet ->getStyle('O'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
    $osheet ->getStyle('P'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
    $osheet ->getStyle('Q'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
    $osheet ->getStyle('R'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
    $objPHPExcel->getActiveSheet()->setTitle('期間內物料入庫月統計表');
                                                            
    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
    $objPHPExcel->setActiveSheetIndex(0);

    // Redirect output to a client’s web browser (Excel5)
    header('Content-Type: application/vnd.ms-excel');  
    header('Content-Disposition: attachment;filename="'. $thismonth . '_monthlycheckinsummary.xls"');   
    header('Cache-Control: max-age=0'); 
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');  
    $objWriter->save('php://output'); 
    exit;    
  }   
  
  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>以下為各倉庫 期間內材料入庫月統計表!! </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
        截止月份:   
        <input name="thismonth" type="text" id="thismonth" onfocus="WdatePicker({dateFmt:'yyyy-MM'})" value="<?=$thismonth;?>"> &nbsp;&nbsp; 
        倉庫: 
        <select name="imd01" id="imd01">  
            <?
              $s1= "select imd01, imd02 from imd_file order by imd01 ";
              $erp_sql1 = oci_parse($erp_conn,$s1 );
              oci_execute($erp_sql1);  
              while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                  echo "<option value=" . $row1["IMD01"];  
                  if ($imd01 == $row1["IMD01"]) echo " selected";                  
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
        <th>規格</th>   
        <th>單位</th>   
        <th style="text-align:right"><?=$montha[1];?></th>  
        <th style="text-align:right"><?=$montha[2];?></th> 
        <th style="text-align:right"><?=$montha[3];?></th>
        <th style="text-align:right"><?=$montha[4];?></th>
        <th style="text-align:right"><?=$montha[5];?></th>  
        <th style="text-align:right"><?=$montha[6];?></th>    
        <th style="text-align:right"><?=$montha[7];?></th>        
        <th style="text-align:right"><?=$montha[8];?></th>    
        <th style="text-align:right"><?=$montha[9];?></th> 
        <th style="text-align:right"><?=$montha[10];?></th>    
        <th style="text-align:right"><?=$montha[11];?></th>  
        <th style="text-align:right"><?=$montha[12];?></th>  
        <th style="text-align:right">合計</th>       
    </tr>
    <?
      $s2 ="select ima01, ima02, ima021, ima25, month01, month02, month03, month04, month05, month06, month07, month08, month09, month10, month11, month12 from " .
           "ima_file left join (select rvv31, sum(month01) month01, sum(month02) month02, sum(month03) month03, sum(month04) month04, sum(month05) month05, sum(month06) month06, sum(month07) month07, sum(month08) month08, sum(month09) month09,sum(month10) month10,sum(month11) month11, sum(month12) month12 ".
           "from (select rvv31, decode(rvu03,'$montha[1]',rvvud07, 0) month01, decode(rvu03,'$montha[2]',rvvud07, 0)  month02, decode(rvu03,'$montha[3]',rvvud07, 0)  month03, decode(rvu03,'$montha[4]',rvvud07, 0) month04, " .
           "                    decode(rvu03,'$montha[5]',rvvud07, 0) month05, decode(rvu03,'$montha[6]',rvvud07, 0)  month06, decode(rvu03,'$montha[7]',rvvud07, 0)  month07, decode(rvu03,'$montha[8]',rvvud07, 0) month08, " .
           "                    decode(rvu03,'$montha[9]',rvvud07, 0) month09, decode(rvu03,'$montha[10]',rvvud07, 0) month10, decode(rvu03,'$montha[11]',rvvud07, 0) month11, decode(rvu03,'$montha[12]',rvvud07, 0) month12 " . 
           "       from (select to_char(rvu03, 'yyyy-mm') rvu03, rvv31 , rvvud07 from rvv_file, rvu_file where  rvv01=rvu01 and rvu00='1') )  group by rvv31 ) " .
           "on rvv31=ima01 where substr(ima06,1,1)!='9' and ima35='$imd01' order by ima01 " ;            //  rvv32='$imd01' and
           //and (month01+month02+month03+month04+month05+month06+month07+month08+month09+month10+month10+month11+month12)>0  ";                            
      
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
              <td><?=$row2["IMA021"];?></td>
              <td><?=$row2["IMA25"];?></td> 
		          <td style="text-align:right"><?=number_format($row2["MONTH01"], 2);?></td>
              <td style="text-align:right"><?=number_format($row2["MONTH02"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH03"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH04"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH05"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH06"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH07"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH08"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH09"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH10"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH11"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH12"], 2);?></td>                                      
              <td style="text-align:right"><?=number_format($row2["MONTH01"]+$row2["MONTH02"]+$row2["MONTH03"]+$row2["MONTH04"]+$row2["MONTH05"]+$row2["MONTH06"]+$row2["MONTH07"]+$row2["MONTH08"]+$row2["MONTH09"]+$row2["MONTH10"]+$row2["MONTH11"]+$row2["MONTH12"] , 2);?></td> 
          </tr>
		  <?
			}
      ?>     
</table>    