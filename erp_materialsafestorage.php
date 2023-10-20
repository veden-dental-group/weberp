<?php
  session_start();
  $pagetitle = "資材部 &raquo; 算出12個月內最高6個月的平均量";
  include("_data.php");
  //auth("erp_materialsafestorage.php");  
  
  function average_max_6($n1,$n2,$n3,$n4,$n5,$n6,$n7,$n8,$n9,$n10,$n11,$n12){

    //傳入12個數字 傳回最大的6個數字的平均值
    $total[0]=$n1;
    $total[1]=$n2;
    $total[2]=$n3;
    $total[3]=$n4;
    $total[4]=$n5;
    $total[5]=$n6;
    $total[6]=$n7;
    $total[7]=$n8;
    $total[8]=$n9;
    $total[9]=$n10;
    $total[10]=$n11;
    $total[11]=$n12;
    rsort($total);

    $average=0;
    for($i=0;$i<6;$i++){
        $average += $total[$i];
    }
    return($average/6);
  }
  
  
  if (is_null($_GET['thismonth'])) {
    $thismonth = date("Y-m",strtotime("-1 month")) ;     
  } else {      
    $thismonth = $_GET['thismonth'] ;   
  }                                                                            
   
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

    $filename='templates/materialsafestorage.xls';
        
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
                                                            
    $osheet     ->setCellValue('F2', $montha[1])
                ->setCellValue('G2', $montha[2]) 
                ->setCellValue('H2', $montha[3]) 
                ->setCellValue('I2', $montha[4]) 
                ->setCellValue('J2', $montha[5]) 
                ->setCellValue('K2', $montha[6]) 
                ->setCellValue('L2', $montha[7]) 
                ->setCellValue('M2', $montha[8]) 
                ->setCellValue('N2', $montha[9]) 
                ->setCellValue('O2', $montha[10]) 
                ->setCellValue('P2', $montha[11]) 
                ->setCellValue('Q2', $montha[12]) ;
    
    
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
                         "select inb04 sfe07, to_char(ina02,'yyyy-mm') sfp03, inb09 out1, 0 out2 from inb_file, ina_file where  inb01=ina01 and inapost='Y' and ina00='1' " .  
           "            ) ) group by sfe07 ) " .                                                                               // inb05='$imd01' and
           "on sfe07=ima01 where substr(ima06,1,1)!='9' and ima35!='3005' and ima35!='5001' and ima35!='9999' order by ima01 ";
           
           //and (month011+month012+month021+month022+month031+month032+month041+month042+month051+month052+month061+month062+ " .
           //"month071+month072+month081+month082+month091+month092+month101+month102+month111+month112+month121+month122 )>0 order by ima01 ";       

    $erp_sql2 = oci_parse($erp_conn,$s2 );
    oci_execute($erp_sql2);   
    $y=5;         
    $i=1;  
    while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {      
          $average6  = round(average_max_6($row2["MONTH011"],$row2["MONTH021"],$row2["MONTH031"],$row2["MONTH041"],$row2["MONTH051"],$row2["MONTH061"],$row2["MONTH071"],$row2["MONTH081"],$row2["MONTH091"],$row2["MONTH101"],$row2["MONTH111"],$row2["MONTH121"]),2);
          $average12 = round(($row2["MONTH011"]+$row2["MONTH021"]+$row2["MONTH031"]+$row2["MONTH041"]+$row2["MONTH051"]+$row2["MONTH061"]+$row2["MONTH071"]+$row2["MONTH081"]+$row2["MONTH091"]+$row2["MONTH101"]+$row2["MONTH111"]+$row2["MONTH121"])/12 , 2);
          
          $osheet     ->setCellValue('A'. $y, $i)
                      ->setCellValue('B'. $y, $row2["IMA01"])
                      ->setCellValue('C'. $y, $row2["IMA02"])
                      ->setCellValue('D'. $y, $row2["IMA021"])
                      ->setCellValue('E'. $y, $row2["IMA25"])
                      ->setCellValue('F'. $y, round($row2["MONTH011"],2))   
                      ->setCellValue('G'. $y, round($row2["MONTH021"],2))  
                      ->setCellValue('H'. $y, round($row2["MONTH031"],2)) 
                      ->setCellValue('I'. $y, round($row2["MONTH041"],2)) 
                      ->setCellValue('J'. $y, round($row2["MONTH051"],2)) 
                      ->setCellValue('K'. $y, round($row2["MONTH061"],2)) 
                      ->setCellValue('L'. $y, round($row2["MONTH071"],2)) 
                      ->setCellValue('M'. $y, round($row2["MONTH081"],2)) 
                      ->setCellValue('N'. $y, round($row2["MONTH091"],2))
                      ->setCellValue('O'. $y, round($row2["MONTH101"],2)) 
                      ->setCellValue('P'. $y, round($row2["MONTH111"],2)) 
                      ->setCellValue('Q'. $y, round($row2["MONTH121"],2)) 
                      ->setCellValue('R'. $y, $average6)
                      ->setCellValue('S'. $y, $average12);
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
    $osheet ->setTitle('期間內物料月安全庫存數');
                                                            
    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
    $objPHPExcel->setActiveSheetIndex(0);

    // Redirect output to a client’s web browser (Excel5)
    header('Content-Type: application/vnd.ms-excel');  
    header('Content-Disposition: attachment;filename="'. $thismonth . '_monthlysafestorage.xls"');   
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
        <th style="text-align:right"><?=$montha[1];?> </th>  
        <th style="text-align:right"><?=$montha[2];?> </th> 
        <th style="text-align:right"><?=$montha[3];?> </th>    
        <th style="text-align:right"><?=$montha[4];?> </th>    
        <th style="text-align:right"><?=$montha[5];?> </th>    
        <th style="text-align:right"><?=$montha[6];?> </th>    
        <th style="text-align:right"><?=$montha[7];?> </th>    
        <th style="text-align:right"><?=$montha[8];?> </th>   
        <th style="text-align:right"><?=$montha[9];?> </th>    
        <th style="text-align:right"><?=$montha[10];?></th>   
        <th style="text-align:right"><?=$montha[11];?></th>    
        <th style="text-align:right"><?=$montha[12];?></th> 
        <th style="text-align:right">6月平均</th> 
        <th style="text-align:right">12月平均</th>       
          
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
                         "select inb04 sfe07, to_char(ina02,'yyyy-mm') sfp03, inb09 out1, 0 out2 from inb_file, ina_file where  inb01=ina01 and inapost='Y' and ina00='1' " .  
           "            ) ) group by sfe07 ) " .                                                                               // inb05='$imd01' and
           "on sfe07=ima01 where substr(ima06,1,1)!='9'  and ima35!='3005' and ima35!='5001' and ima35!='9999' order by ima01 ";
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
              <td style="text-align:right"><?=number_format($row2["MONTH021"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH031"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH041"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH051"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH061"], 2);?></td>       
              <td style="text-align:right"><?=number_format($row2["MONTH071"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["MONTH081"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH091"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["MONTH101"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["MONTH111"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["MONTH121"], 2);?></td>  
              <td style="text-align:right"><?=number_format(average_max_6($row2["MONTH011"],$row2["MONTH021"],$row2["MONTH031"],$row2["MONTH041"],$row2["MONTH051"],$row2["MONTH061"],$row2["MONTH071"],$row2["MONTH081"],$row2["MONTH091"],$row2["MONTH101"],$row2["MONTH111"],$row2["MONTH121"]), 2);?></td> 
              <td style="text-align:right"><?=number_format(($row2["MONTH011"]+$row2["MONTH021"]+$row2["MONTH031"]+$row2["MONTH041"]+$row2["MONTH051"]+$row2["MONTH061"]+$row2["MONTH071"]+$row2["MONTH081"]+$row2["MONTH091"]+$row2["MONTH101"]+$row2["MONTH111"]+$row2["MONTH121"])/12 , 2);?></td> 
              
          </tr>
		  <?
			}
      ?>     
</table>    