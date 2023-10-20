<?php
  session_start();
  $pagetitle = "資材部 &raquo; 期間內物料領用月統計表";
  include("_data.php");
  //auth("erp_materialmonthlycheckoutsummary.php");  
  
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

    $filename='templates/materialyearlycheckoutsummary.xls';
        
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
                 
    $osheet->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
    $osheet->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);             
    $osheet->getDefaultStyle()->getFont()->setName('Calibri');
    $osheet->getDefaultStyle()->getFont()->setSize(12);
    $osheet->getDefaultRowDimension()->setRowHeight(15);              
                 
    $osheet     ->setCellValue('B2', $imd01) ; 
    $osheet     ->setCellValue('F3', $montha[1])
                ->setCellValue('I3', $montha[2]) 
                ->setCellValue('L3', $montha[3]) 
                ->setCellValue('O3', $montha[4]) 
                ->setCellValue('R3', $montha[5]) 
                ->setCellValue('U3', $montha[6]) 
                ->setCellValue('X3', $montha[7]) 
                ->setCellValue('AA3', $montha[8]) 
                ->setCellValue('AD3', $montha[9]) 
                ->setCellValue('AG3', $montha[10]) 
                ->setCellValue('AJ3', $montha[11]) 
                ->setCellValue('AM3', $montha[12]) 
                ->setCellValue('AP3', '總計'); 
    
    
      $s2 ="select ima01, ima02, ima021, ima25, month011, month012, month021, month022, month031, month032, month041, month042, month051, month052, month061, month062, " .
           "month071, month072, month081, month082, month091, month092, month101, month102, month111, month112, month121, month122 from " .
           "ima_file left join (select sfe07, sum(month011) month011, sum(month012) month012, sum(month021) month021, sum(month022) month022, sum(month031) month031, sum(month032) month032, " .
           "               sum(month041) month041, sum(month042) month042, sum(month051) month051, sum(month052) month052, sum(month061) month061, sum(month062) month062, " .
           "               sum(month071) month071, sum(month072) month072, sum(month081) month081, sum(month082) month082, sum(month091) month091, sum(month092) month092, " .  
           "               sum(month101) month101, sum(month102) month102, sum(month111) month111, sum(month112) month112, sum(month121) month121, sum(month122) month122 " .  
           "from (select sfe07, decode(sfp03,'$montha[1]',out1, 0) month011,  decode(sfp03,'$montha[1]',out2, 0) month012,  decode(sfp03,'$montha[2]',out1, 0)  month021, decode(sfp03,'$montha[2]',out2, 0) month022, " .
           "                    decode(sfp03,'$montha[3]',out1, 0) month031,  decode(sfp03,'$montha[3]',out2, 0) month032,  decode(sfp03,'$montha[4]',out1, 0)  month041, decode(sfp03,'$montha[4]',out2, 0)  month042, " .    
           "                    decode(sfp03,'$montha[5]',out1, 0) month051,  decode(sfp03,'$montha[5]',out2, 0) month052,  decode(sfp03,'$montha[6]',out1, 0)  month061, decode(sfp03,'$montha[6]',out2, 0)  month062, " .    
           "                    decode(sfp03,'$montha[7]',out1, 0) month071,  decode(sfp03,'$montha[7]',out2, 0) month072,  decode(sfp03,'$montha[8]',out1, 0)  month081, decode(sfp03,'$montha[8]',out2, 0)  month082, " .    
           "                    decode(sfp03,'$montha[9]',out1, 0) month091,  decode(sfp03,'$montha[9]',out2, 0) month092,  decode(sfp03,'$montha[10]',out1, 0) month101, decode(sfp03,'$montha[10]',out2, 0)  month102, " .    
           "                    decode(sfp03,'$montha[11]',out1, 0) month111, decode(sfp03,'$montha[11]',out2, 0) month112, decode(sfp03,'$montha[12]',out1, 0) month121, decode(sfp03,'$montha[12]',out2, 0)  month122 " .    
           "       from ( " . 
                         "select sfe07      , to_char(sfp03,'yyyy-mm') sfp03, sfe16 out1, 0 out2 from sfe_file, sfp_file where  sfe02=sfp01 and sfp04='Y' and sfp06='1'   " .
                         "union all " .                                                                                        // sfe08='$imd01' and
                         "select inb04 sfe07, to_char(ina02,'yyyy-mm') sfp03, 0 out1, inb09 out2 from inb_file, ina_file where  inb01=ina01 and inapost='Y' and ina00='1' " .  
           "            ) ) group by sfe07 ) " .                                                                               // inb05='$imd01' and
           "on sfe07=ima01 where substr(ima06,1,1)!='9' and ima35='$imd01' order by ima01 ";
           
           //and (month011+month012+month021+month022+month031+month032+month041+month042+month051+month052+month061+month062+ " .
           //"month071+month072+month081+month082+month091+month092+month101+month102+month111+month112+month121+month122 )>0 order by ima01 ";       

    $erp_sql2 = oci_parse($erp_conn,$s2 );
    oci_execute($erp_sql2);   
    $y=5;         
    $i=1;  
    while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {     
          $ap="=(F$y+I$y+L$y+O$y+R$y+U$y+X$y+AA$y+AD$y+AG$y+AJ$y+AM$y)";
          $aq="=(G$y+J$y+M$y+P$y+S$y+V$y+Y$y+AB$y+AE$y+AH$y+AK$y+AN$y)";
          
          $osheet     ->setCellValue('A'. $y, $i)
                      ->setCellValue('B'. $y, $row2["IMA01"])
                      ->setCellValue('C'. $y, $row2["IMA02"])
                      ->setCellValue('D'. $y, $row2["IMA021"])
                      ->setCellValue('E'. $y, $row2["IMA25"])
                      ->setCellValue('F'. $y, $row2["MONTH011"])
                      ->setCellValue('G'. $y, $row2["MONTH012"]) 
                      ->setCellValue('H'. $y, "=sum(F$y:G$y)")  
                      ->setCellValue('I'. $y, $row2["MONTH021"])
                      ->setCellValue('J'. $y, $row2["MONTH022"]) 
                      ->setCellValue('K'. $y, "=sum(I$y:J$y)")
                      ->setCellValue('L'. $y, $row2["MONTH031"])
                      ->setCellValue('M'. $y, $row2["MONTH032"]) 
                      ->setCellValue('N'. $y, "=sum(L$y:M$y)")
                      ->setCellValue('O'. $y, $row2["MONTH041"])
                      ->setCellValue('P'. $y, $row2["MONTH042"]) 
                      ->setCellValue('Q'. $y, "=sum(O$y:P$y)")
                      ->setCellValue('R'. $y, $row2["MONTH051"])
                      ->setCellValue('S'. $y, $row2["MONTH052"]) 
                      ->setCellValue('T'. $y, "=sum(R$y:S$y)")
                      ->setCellValue('U'. $y, $row2["MONTH061"])
                      ->setCellValue('V'. $y, $row2["MONTH062"]) 
                      ->setCellValue('W'. $y, "=sum(U$y:V$y)")
                      ->setCellValue('X'. $y, $row2["MONTH071"])
                      ->setCellValue('Y'. $y, $row2["MONTH072"]) 
                      ->setCellValue('Z'. $y, "=sum(X$y:Y$y)")
                      ->setCellValue('AA'. $y, $row2["MONTH081"])
                      ->setCellValue('AB'. $y, $row2["MONTH082"]) 
                      ->setCellValue('AC'. $y, "=sum(AA$y:AB$y)")
                      ->setCellValue('AD'. $y, $row2["MONTH091"])
                      ->setCellValue('AE'. $y, $row2["MONTH092"]) 
                      ->setCellValue('AF'. $y, "=sum(AD$y:AE$y)")
                      ->setCellValue('AG'. $y, $row2["MONTH101"])
                      ->setCellValue('AH'. $y, $row2["MONTH102"]) 
                      ->setCellValue('AI'. $y, "=sum(AG$y:AH$y)")
                      ->setCellValue('AJ'. $y, $row2["MONTH111"])
                      ->setCellValue('AK'. $y, $row2["MONTH112"]) 
                      ->setCellValue('AL'. $y, "=sum(AJ$y:AK$y)")  
                      ->setCellValue('AM'. $y, $row2["MONTH121"])
                      ->setCellValue('AN'. $y, $row2["MONTH122"]) 
                      ->setCellValue('AO'. $y, "=sum(AM$y:AN$y)")  
                      ->setCellValue('AP'. $y, $ap)
                      ->setCellValue('AQ'. $y, $aq)   
                      ->setCellValue('AR'. $y, "=sum(AP$y:AQ$y)");
                        
          if (($y%2)==0){
              $osheet->getStyle('A'.$y.':AR'.$y)->getFill() ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFCCFFFF');    
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
    $osheet ->getStyle('S'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('T'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('U'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('V'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('W'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('X'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('Y'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('Z'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('AA'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('AB'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('AC'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('AD'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('AE'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('AF'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('AG'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('AH'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('AI'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('AJ'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('AK'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('AL'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('AM'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('AN'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('AO'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('AP'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('AQ'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->getStyle('AR'.($y-1)) ->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
    $osheet ->setTitle('期間內物料領用月統計表');
                                                            
    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
    $objPHPExcel->setActiveSheetIndex(0);

    // Redirect output to a client’s web browser (Excel5)
    header('Content-Type: application/vnd.ms-excel');  
    header('Content-Disposition: attachment;filename="'. $thismonth . '_monthlycheckoutsummary.xls"');   
    header('Cache-Control: max-age=0'); 
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');  
    $objWriter->save('php://output'); 
    exit;    
  }   
  
  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>以下為各倉庫 期間內物料領用月統計表!! </p>     
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
        <th style="text-align:right"><?=$montha[1];?> 對單</th>    
        <th style="text-align:right"><?=$montha[1];?> 雜發</th>  
        <th style="text-align:right"><?=$montha[2];?> 對單</th>    
        <th style="text-align:right"><?=$montha[2];?> 雜發</th> 
        <th style="text-align:right"><?=$montha[3];?> 對單</th>    
        <th style="text-align:right"><?=$montha[3];?> 雜發</th> 
        <th style="text-align:right"><?=$montha[4];?> 對單</th>    
        <th style="text-align:right"><?=$montha[4];?> 雜發</th> 
        <th style="text-align:right"><?=$montha[5];?> 對單</th>    
        <th style="text-align:right"><?=$montha[5];?> 雜發</th> 
        <th style="text-align:right"><?=$montha[6];?> 對單</th>    
        <th style="text-align:right"><?=$montha[6];?> 雜發</th> 
        <th style="text-align:right"><?=$montha[7];?> 對單</th>    
        <th style="text-align:right"><?=$montha[7];?> 雜發</th> 
        <th style="text-align:right"><?=$montha[8];?> 對單</th>    
        <th style="text-align:right"><?=$montha[8];?> 雜發</th> 
        <th style="text-align:right"><?=$montha[9];?> 對單</th>    
        <th style="text-align:right"><?=$montha[9];?> 雜發</th> 
        <th style="text-align:right"><?=$montha[10];?> 對單</th>    
        <th style="text-align:right"><?=$montha[10];?> 雜發</th> 
        <th style="text-align:right"><?=$montha[11];?> 對單</th>    
        <th style="text-align:right"><?=$montha[11];?> 雜發</th>  
        <th style="text-align:right"><?=$montha[12];?> 對單</th>    
        <th style="text-align:right"><?=$montha[12];?> 雜發</th> 
        <th style="text-align:right">合計 對單</th> 
        <th style="text-align:right">合計 雜發</th>       
          
    </tr>
    <?
      $s2 ="select ima01, ima02, ima021, ima25, month011, month012, month021, month022, month031, month032, month041, month042, month051, month052, month061, month062, " .
           "month071, month072, month081, month082, month091, month092, month101, month102, month111, month112, month121, month122 from " .
           "ima_file left join (select sfe07, sum(month011) month011, sum(month012) month012, sum(month021) month021, sum(month022) month022, sum(month031) month031, sum(month032) month032, " .
           "               sum(month041) month041, sum(month042) month042, sum(month051) month051, sum(month052) month052, sum(month061) month061, sum(month062) month062, " .
           "               sum(month071) month071, sum(month072) month072, sum(month081) month081, sum(month082) month082, sum(month091) month091, sum(month092) month092, " .  
           "               sum(month101) month101, sum(month102) month102, sum(month111) month111, sum(month112) month112, sum(month121) month121, sum(month122) month122 " .  
           "from (select sfe07, decode(sfp03,'$montha[1]',out1, 0) month011,  decode(sfp03,'$montha[1]',out2, 0) month012,  decode(sfp03,'$montha[2]',out1, 0)  month021, decode(sfp03,'$montha[2]',out2, 0) month022, " .
           "                    decode(sfp03,'$montha[3]',out1, 0) month031,  decode(sfp03,'$montha[3]',out2, 0) month032,  decode(sfp03,'$montha[4]',out1, 0)  month041, decode(sfp03,'$montha[4]',out2, 0)  month042, " .    
           "                    decode(sfp03,'$montha[5]',out1, 0) month051,  decode(sfp03,'$montha[5]',out2, 0) month052,  decode(sfp03,'$montha[6]',out1, 0)  month061, decode(sfp03,'$montha[6]',out2, 0)  month062, " .    
           "                    decode(sfp03,'$montha[7]',out1, 0) month071,  decode(sfp03,'$montha[7]',out2, 0) month072,  decode(sfp03,'$montha[8]',out1, 0)  month081, decode(sfp03,'$montha[8]',out2, 0)  month082, " .    
           "                    decode(sfp03,'$montha[9]',out1, 0) month091,  decode(sfp03,'$montha[9]',out2, 0) month092,  decode(sfp03,'$montha[10]',out1, 0) month101, decode(sfp03,'$montha[10]',out2, 0)  month102, " .    
           "                    decode(sfp03,'$montha[11]',out1, 0) month111, decode(sfp03,'$montha[11]',out2, 0) month112, decode(sfp03,'$montha[12]',out1, 0) month121, decode(sfp03,'$montha[12]',out2, 0)  month122 " .    
           "       from ( " . 
                         "select sfe07      , to_char(sfp03,'yyyy-mm') sfp03, sfe16 out1, 0 out2 from sfe_file, sfp_file where  sfe02=sfp01 and sfp04='Y' and sfp06='1'   " .
                         "union all " .                                                                                        // sfe08='$imd01' and
                         "select inb04 sfe07, to_char(ina02,'yyyy-mm') sfp03, 0 out1, inb09 out2 from inb_file, ina_file where  inb01=ina01 and inapost='Y' and ina00='1' " .  
           "            ) ) group by sfe07 ) " .                                                                               // inb05='$imd01' and
           "on sfe07=ima01 where substr(ima06,1,1)!='9' and ima35='$imd01' order by ima01 ";
           //and (month011+month012+month021+month022+month031+month032+month041+month042+month051+month052+month061+month062+ " .
           //"month071+month072+month081+month082+month091+month092+month101+month102+month111+month112+month121+month122 )>0 order by ima01 ";                            
      
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
		          <td style="text-align:right"><?=number_format($row2["MONTH011"], 2);?></td>
              <td style="text-align:right"><?=number_format($row2["MONTH012"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH021"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH022"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH031"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH032"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH041"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH042"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH051"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH052"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH061"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH062"], 2);?></td>       
              <td style="text-align:right"><?=number_format($row2["MONTH071"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH072"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH081"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH082"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH091"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH092"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH101"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH102"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH111"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH112"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH121"], 2);?></td>                                 
              <td style="text-align:right"><?=number_format($row2["MONTH122"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH011"]+$row2["MONTH021"]+$row2["MONTH031"]+$row2["MONTH041"]+$row2["MONTH051"]+$row2["MONTH061"]+$row2["MONTH071"]+$row2["MONTH081"]+$row2["MONTH091"]+$row2["MONTH101"]+$row2["MONTH111"]+$row2["MONTH121"] , 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["MONTH012"]+$row2["MONTH022"]+$row2["MONTH032"]+$row2["MONTH042"]+$row2["MONTH052"]+$row2["MONTH062"]+$row2["MONTH072"]+$row2["MONTH082"]+$row2["MONTH092"]+$row2["MONTH102"]+$row2["MONTH112"]+$row2["MONTH122"] , 2);?></td> 
          </tr>
		  <?
			}
      ?>     
</table>    