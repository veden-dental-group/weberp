<?php
  session_start();
  $pagetitle = "資材部 &raquo; 每月材料領用一覽表";
  include("_data.php");
  auth("erp_monthlyusage.php");  
  
  if (is_null($_GET['thismonth'])) {
    $thismonth = date("Y-m",strtotime("-1 month")) ;     
  } else {      
    $thismonth = $_GET['thismonth'] ;   
  }    
  $previousmonth=date('Y-m',strtotime("-1 month", strtotime($thismonth."-01")));
  $imd01=$_GET['imd01']; 
  
  if ($_GET["submit"]=="匯出") {   

    $filename='templates/monthlyusage.xls';
        
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
                ->setCellValue('R2', $thismonth);   
                      
    $y=4;         
    $i=1;  
    
    if ($imd01!='F000'){ //不是F線     
             //上個月庫存 
        $sql="(select imk01 code, imk09, 0 normalin, 0 backin, 0 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from imk_file                   where imk02='$imd01' and (trim(to_char(imk05,'9999'))|| '-'||trim(to_char(imk06,'00'))='$previousmonth') " .
             "union all " . //雜收
             "select inb04 code, 0 imk09, 0 normalin, 0 backin, inb09 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from inb_file, ina_file    where inb05='$imd01' and inb01=ina01 and inapost='Y' and ina00='3' and to_char(ina02,'yyyy-mm')='$thismonth' " .  
             "union all " . //雜發     
             "select inb04 code, 0 imk09, 0 normalin, 0 backin, 0 zain, 0 dioin, 0 normalout, inb09 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from inb_file, ina_file    where inb05='$imd01' and inb01=ina01 and inapost='Y' and ina00='1' and to_char(ina02,'yyyy-mm')='$thismonth' " . 
             "union all " .  //按單發
             "select sfe07 code, 0 imk09, 0 normalin, 0 backin, 0 zain, 0 dioin,  sfe16 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from sfe_file, sfp_file   where sfe08='$imd01' and sfe02=sfp01 and sfp04='Y' and sfp06='1' and to_char(sfp03,'yyyy-mm')='$thismonth' " .
             "union all " .  //退料
             "select sfe07 code, 0 imk09, 0 normalin, sfe16 backin, 0 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from sfe_file, sfp_file    where sfe08='$imd01' and sfe02=sfp01 and sfp04='Y' and sfp06='8' and to_char(sfp03,'yyyy-mm')='$thismonth' " . 
             "union all " .  //調出
             "select imn03 code, 0 imk09, 0 normalin, 0 backin, 0 zain, 0 dioin, 0 normalout, 0 zaout, imn10 dioout, 0 noorderout, 0 borrowin, 0 borrowout from imn_file, imm_file    where imn04='$imd01' and imn01=imm01 and imn04='$imd01' and imm03='Y' and to_char(imm02,'yyyy-mm')='$thismonth' " . 
             "union all " .  //調入
             "select imn03 code, 0 imk09, 0 normalin, 0 backin, 0 zain, imn22 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from imn_file, imm_file    where imn15='$imd01' and imn01=imm01 and imn15='$imd01' and imm03='Y' and to_char(imm02,'yyyy-mm')='$thismonth' " .
             "union all " . //無訂單出貨 
             "select ogb04 code, 0 imk09, 0 normalin, 0 backin, 0 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, ogb12 noorderout, 0 borrowin, 0 borrowout from oga_file,ogb_file where ogb09='$imd01' and ogb01=oga01 and oga905 is null and ogapost='Y' and to_char(oga02,'yyyy-mm')='$thismonth' " .
             "union all " . //同業借入 
             "select imp03 code, 0 imk09, 0 normalin, 0 backin, 0 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, imp04 borrowin, 0 borrowout from imp_file,imo_file where imp11='$imd01' and imp01=imo01 and imopost='Y' and to_char(imo02,'yyyy-mm')='$thismonth' " .
             "union all " . //歸還同業
             "select imq05 code, 0 imk09, 0 normalin, 0 backin, 0 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, imq07 borrowout from imq_file,imr_file where imq08='$imd01' and imq01=imr01 and imrpost='Y' and to_char(imr09,'yyyy-mm')='$thismonth' " .
             "union all " .   //入庫
             "select rvv31 code, 0 imk09, rvvud07 normalin, 0 backin,  0 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from rvv_file, rvu_file where rvv32='$imd01' and rvv01=rvu01 and rvv03='1' and to_char(rvu03,'yyyy-mm')='$thismonth') "; 
        $s2 ="select ima01, ima02, ima25, sum(imk09) imk09, sum(normalin) normalin, sum(backin) backin, sum(zain) zain, sum(dioin) dioin,  sum(normalout) normalout, sum(zaout) zaout, sum(dioout) dioout, sum(noorderout) noorderout, sum(borrowin) borrowin, sum(borrowout) borrowout from " .
             "ima_file, $sql b where substr(ima06,1,1)!='9'  and code=ima01 group by (ima01,ima02, ima25) order by ima01 ";
    } else {
        $sql="(select imk01 code, imk09, 0 normalin, 0 backin, 0 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from imk_file                   where imk02='F000' and (trim(to_char(imk05,'9999'))|| '-'||trim(to_char(imk06,'00'))='$previousmonth') " .
             "union all " . 
             "select inb04 code, 0 imk09, 0 normalin, 0 backin, inb09 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from inb_file, ina_file    where inb05='F000' and inb01=ina01 and inapost='Y' and ina00='3' and to_char(ina02,'yyyy-mm')='$thismonth' " .  
             "union all " .      
             "select inb04 code, 0 imk09, 0 normalin, 0 backin, 0 zain, 0 dioin, 0 normalout, inb09 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from inb_file, ina_file    where inb05='F000' and inb01=ina01 and inapost='Y' and ina00='1' and to_char(ina02,'yyyy-mm')='$thismonth' " . 
             "union all " .  
             "select sfe07 code, 0 imk09, 0 normalin, 0 backin, 0 zain, 0 dioin,  sfe16 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from sfe_file, sfp_file   where sfe08='F000' and sfe02=sfp01 and sfp04='Y' and sfp06='1' and to_char(sfp03,'yyyy-mm')='$thismonth' " .
             "union all " .  
             "select sfe07 code, 0 imk09, 0 normalin, sfe16 backin, 0 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from sfe_file, sfp_file    where sfe08='F000' and sfe02=sfp01 and sfp04='Y' and sfp06='8' and to_char(sfp03,'yyyy-mm')='$thismonth' " . 
             "union all " .  
             "select imn03 code, 0 imk09, 0 normalin, 0 backin, 0 zain, 0 dioin, 0 normalout, 0 zaout, imn10 dioout, 0 noorderout, 0 borrowin, 0 borrowout from imn_file, imm_file    where imn04='F000' and imn01=imm01 and imn04='$imd01' and imm03='Y' and to_char(imm02,'yyyy-mm')='$thismonth' " . 
             "union all " .  
             "select imn03 code, 0 imk09, 0 normalin, 0 backin, 0 zain, imn22 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from imn_file, imm_file    where imn15='F000' and imn01=imm01 and imn15='$imd01' and imm03='Y' and to_char(imm02,'yyyy-mm')='$thismonth' " .
             "union all " . //無訂單出貨 
             "select ogb04 code, 0 imk09, 0 normalin, 0 backin, 0 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, ogb12 noorderout, 0 borrowin, 0 borrowout from oga_file,ogb_file where ogb09='F000' and ogb01=oga01 and oga905 is null and ogapost='Y' and to_char(oga02,'yyyy-mm')='$thismonth' " .
             "union all " . //同業借入 
             "select imp03 code, 0 imk09, 0 normalin, 0 backin, 0 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, imp04 borrowin, 0 borrowout from imp_file,imo_file where imp11='F000' and imp01=imo01 and imopost='Y' and to_char(imo02,'yyyy-mm')='$thismonth' " .
             "union all " . //歸還同業
             "select imq05 code, 0 imk09, 0 normalin, 0 backin, 0 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, imq07 borrowout from imq_file,imr_file where imq08='F000' and imq01=imr01 and imrpost='Y' and to_char(imr09,'yyyy-mm')='$thismonth' " .
             "union all " .  
             "select rvv31 code, 0 imk09, rvvud07 normalin, 0 backin,  0 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from rvv_file, rvu_file where rvv32='F000' and rvv01=rvu01 and rvv03='1' and to_char(rvu03,'yyyy-mm')='$thismonth') "; 
             
        $s2 ="select ima01, ima02, ima25, sum(imk09) imk09, sum(normalin) normalin, sum(backin) backin, sum(zain) zain, sum(dioin) dioin,  sum(normalout) normalout, sum(zaout) zaout, sum(dioout) dioout, sum(noorderout) noorderout, sum(borrowin) borrowin, sum(borrowout) borrowout from " .
             "ima_file, $sql b where substr(ima06,1,1)!='9' and code=ima01 group by (ima01,ima02, ima25) order by ima01 ";       
    }  
    $erp_sql2 = oci_parse($erp_conn,$s2 );
    oci_execute($erp_sql2);   
    while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {     
          $objPHPExcel->setActiveSheetIndex(0)
                      ->setCellValue('A'. $y, $i)
                      ->setCellValue('B'. $y, $row2["IMA01"])
                      ->setCellValue('C'. $y, $row2["IMA02"])
                      ->setCellValue('D'. $y, $row2["IMK09"]+$row2["NORMALIN"]+$row2["BACKIN"]+$row2["DIOIN"]+$row2["BORROWIN"]+$row2["ZAIN"]-$row2["NORMALOUT"]-$row2["ZAOUT"]-$row2["DIOOUT"]-$row2["NOORDEROUT"]-$row2["BORROWOUT"])   
                      ->setCellValue('E'. $y, $row2["IMK09"])   
                      ->setCellValue('F'. $y, $row2["IMA25"])      
                      ->setCellValue('G'. $y, $row2["NORMALIN"])   
                      ->setCellValue('H'. $y, $row2["BACKIN"])   
                      ->setCellValue('I'. $y, $row2["ZAIN"])  
                      ->setCellValue('J'. $y, $row2["DIOIN"])  
                      ->setCellValue('K'. $y, $row2["BORROWIN"])   
                      ->setCellValue('L'. $y, $row2["NORMALIN"]+$row2["BACKIN"]+$row2["ZAIN"]+$row2["DIOIN"]+$row2["BORROWIN"])   
                      ->setCellValue('M'. $y, $row2["NORMALOUT"])    
                      ->setCellValue('N'. $y, $row2["ZAOUT"])
                      ->setCellValue('O'. $y, $row2["DIOOUT"]) 
                      ->setCellValue('P'. $y, $row2["NOORDEROUT"])  
                      ->setCellValue('Q'. $y, $row2["BORROWOUT"])   
                      ->setCellValue('R'. $y, $row2["NORMALOUT"]+$row2["DIOOUT"]+$row2["ZAOUT"]+$row2["NOORDEROUT"]+$row2["BORROWOUT"]);       
          $y++;
          $i++;
    }                      
    
    //total
    //$unit='=sum(G14:G' . ($y-1) . ')'; 
    //$objPHPExcel->setActiveSheetIndex(0)
    //            ->setCellValue('A'.($y+1), 'Total') 
    //            ->setCellValue('G'.($y+1), '=sum(G14:G' . ($y-1) . ')')
    //            ->setCellValue('J'.($y+1), '=sum(J14:J' . ($y-1) . ')');  
                       
    // Rename sheet
    $objPHPExcel->getActiveSheet()->setTitle('每月領用一覽表');
                                                            
    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
    $objPHPExcel->setActiveSheetIndex(0);

    // Redirect output to a client’s web browser (Excel5)
    header('Content-Type: application/vnd.ms-excel');  
    header('Content-Disposition: attachment;filename="'. $thismonth . '_monthlyusage.xls"');   
    header('Cache-Control: max-age=0'); 
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');  
    $objWriter->save('php://output'); 
    exit;    
  }   
  
  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>以下為各倉庫的物料領用一覽表!! </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
        月份:   
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
        <th>單位</th>   
        <th style="text-align:right">期初庫存</th>  
        <th style="text-align:right">驗收入庫</th> 
        <th style="text-align:right">退料</th>
        <th style="text-align:right">雜收</th>
        <th style="text-align:right">調進</th>  
        <th style="text-align:right">同業借入</th>    
        <th style="text-align:right">小計</th>        
        <th style="text-align:right">按單發料</th>    
        <th style="text-align:right">雜發</th> 
        <th style="text-align:right">調出</th>    
        <th style="text-align:right">無訂出</th>  
        <th style="text-align:right">還給同業</th>  
        <th style="text-align:right">小計</th>      
        <th style="text-align:right">目前庫存</th>   
    </tr>
    <?
      if ($imd01!='F000'){ //不是F線     
          $sql="(select imk01 code, imk09, 0 normalin, 0 backin, 0 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout  from imk_file                 where imk02='$imd01' and (trim(to_char(imk05,'9999'))|| '-'||trim(to_char(imk06,'00'))='$previousmonth') " .
               "union all " . 
               "select inb04 code, 0 imk09, 0 normalin, 0 backin, inb09 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from inb_file, ina_file    where inb05='$imd01' and inb01=ina01 and inapost='Y' and ina00='3' and to_char(ina02,'yyyy-mm')='$thismonth' " .  
               "union all " .      
               "select inb04 code, 0 imk09, 0 normalin, 0 backin, 0 zain, 0 dioin, 0 normalout, inb09 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from inb_file, ina_file    where inb05='$imd01' and inb01=ina01 and inapost='Y' and ina00='1' and to_char(ina02,'yyyy-mm')='$thismonth' " . 
               "union all " .  
               "select sfe07 code, 0 imk09, 0 normalin, 0 backin, 0 zain, 0 dioin,  sfe16 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from sfe_file, sfp_file   where sfe08='$imd01' and sfe02=sfp01 and sfp04='Y' and sfp06='1' and to_char(sfp03,'yyyy-mm')='$thismonth' " .
               "union all " .  
               "select sfe07 code, 0 imk09, 0 normalin, sfe16 backin, 0 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from sfe_file, sfp_file    where sfe08='$imd01' and sfe02=sfp01 and sfp04='Y' and sfp06='8' and to_char(sfp03,'yyyy-mm')='$thismonth' " . 
               "union all " .  
               "select imn03 code, 0 imk09, 0 normalin, 0 backin, 0 zain, 0 dioin, 0 normalout, 0 zaout, imn10 dioout, 0 noorderout, 0 borrowin, 0 borrowout from imn_file, imm_file    where imn04='$imd01' and imn01=imm01 and imn04='$imd01' and imm03='Y' and to_char(imm02,'yyyy-mm')='$thismonth' " . 
               "union all " .  
               "select imn03 code, 0 imk09, 0 normalin, 0 backin, 0 zain, imn22 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from imn_file, imm_file    where imn15='$imd01' and imn01=imm01 and imn15='$imd01' and imm03='Y' and to_char(imm02,'yyyy-mm')='$thismonth' " .
               "union all " . //無訂單出貨 
               "select ogb04 code, 0 imk09, 0 normalin, 0 backin, 0 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, ogb12 noorderout, 0 borrowin, 0 borrowout from oga_file,ogb_file where ogb09='$imd01' and ogb01=oga01 and oga905 is null and ogapost='Y' and to_char(oga02,'yyyy-mm')='$thismonth' " .
               "union all " . //同業借入 
               "select imp03 code, 0 imk09, 0 normalin, 0 backin, 0 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, imp04 borrowin, 0 borrowout from imp_file,imo_file where imp11='$imd01' and imp01=imo01 and imopost='Y' and to_char(imo02,'yyyy-mm')='$thismonth' " .
               "union all " . //歸還同業
               "select imq05 code, 0 imk09, 0 normalin, 0 backin, 0 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, imq07 borrowout from imq_file,imr_file where imq08='$imd01' and imq01=imr01 and imrpost='Y' and to_char(imr09,'yyyy-mm')='$thismonth' " .
               "union all " .  
               "select rvv31 code, 0 imk09, rvvud07 normalin, 0 backin,  0 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from rvv_file, rvu_file where rvv32='$imd01' and rvv01=rvu01 and rvv03='1' and to_char(rvu03,'yyyy-mm')='$thismonth') "; 
          $s2 ="select ima01, ima02, ima25, sum(imk09) imk09, sum(normalin) normalin, sum(backin) backin, sum(zain) zain, sum(dioin) dioin,  sum(normalout) normalout, sum(zaout) zaout, sum(dioout) dioout, sum(noorderout) noorderout, sum(borrowin) borrowin, sum(borrowout) borrowout from " .
               "ima_file, $sql b where substr(ima06,1,1)!='9'  and code=ima01 group by (ima01,ima02, ima25) order by ima01 ";
      } else {
          $sql="(select imk01 code, imk09, 0 normalin, 0 backin, 0 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from imk_file                   where imk02='F000' and (trim(to_char(imk05,'9999'))|| '-'||trim(to_char(imk06,'00'))='$previousmonth') " .
               "union all " . 
               "select inb04 code, 0 imk09, 0 normalin, 0 backin, inb09 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from inb_file, ina_file    where inb05='F000' and inb01=ina01 and inapost='Y' and ina00='3' and to_char(ina02,'yyyy-mm')='$thismonth' " .  
               "union all " .      
               "select inb04 code, 0 imk09, 0 normalin, 0 backin, 0 zain, 0 dioin, 0 normalout, inb09 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from inb_file, ina_file    where inb05='F000' and inb01=ina01 and inapost='Y' and ina00='1' and to_char(ina02,'yyyy-mm')='$thismonth' " . 
               "union all " .  
               "select sfe07 code, 0 imk09, 0 normalin, 0 backin, 0 zain, 0 dioin,  sfe16 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from sfe_file, sfp_file   where sfe08='F000' and sfe02=sfp01 and sfp04='Y' and sfp06='1' and to_char(sfp03,'yyyy-mm')='$thismonth' " .
               "union all " .  
               "select sfe07 code, 0 imk09, 0 normalin, sfe16 backin, 0 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from sfe_file, sfp_file    where sfe08='F000' and sfe02=sfp01 and sfp04='Y' and sfp06='8' and to_char(sfp03,'yyyy-mm')='$thismonth' " . 
               "union all " .  
               "select imn03 code, 0 imk09, 0 normalin, 0 backin, 0 zain, 0 dioin, 0 normalout, 0 zaout, imn10 dioout, 0 noorderout, 0 borrowin, 0 borrowout from imn_file, imm_file    where imn04='F000' and imn01=imm01 and imn04='$imd01' and imm03='Y' and to_char(imm02,'yyyy-mm')='$thismonth' " . 
               "union all " .  
               "select imn03 code, 0 imk09, 0 normalin, 0 backin, 0 zain, imn22 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from imn_file, imm_file    where imn15='F000' and imn01=imm01 and imn15='$imd01' and imm03='Y' and to_char(imm02,'yyyy-mm')='$thismonth' " .
               "union all " . //無訂單出貨 
               "select ogb04 code, 0 imk09, 0 normalin, 0 backin, 0 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, ogb12 noorderout, 0 borrowin, 0 borrowout from oga_file,ogb_file where ogb09='F000' and ogb01=oga01 and oga905 is null and ogapost='Y' and to_char(oga02,'yyyy-mm')='$thismonth' " .
               "union all " . //同業借入 
               "select imp03 code, 0 imk09, 0 normalin, 0 backin, 0 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, imp04 borrowin, 0 borrowout from imp_file,imo_file where imp11='F000' and imp01=imo01 and imopost='Y' and to_char(imo02,'yyyy-mm')='$thismonth' " .
               "union all " . //歸還同業
               "select imq05 code, 0 imk09, 0 normalin, 0 backin, 0 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, imq07 borrowout from imq_file,imr_file where imq08='F000' and imq01=imr01 and imrpost='Y' and to_char(imr09,'yyyy-mm')='$thismonth' " .
               "union all " .  
               "select rvv31 code, 0 imk09, rvvud07 normalin, 0 backin,  0 zain, 0 dioin, 0 normalout, 0 zaout, 0 dioout, 0 noorderout, 0 borrowin, 0 borrowout from rvv_file, rvu_file where rvv32='F000' and rvv01=rvu01 and rvv03='1' and to_char(rvu03,'yyyy-mm')='$thismonth') "; 
               
          $s2 ="select ima01, ima02, ima25, sum(imk09) imk09, sum(normalin) normalin, sum(backin) backin, sum(zain) zain, sum(dioin) dioin,  sum(normalout) normalout, sum(zaout) zaout, sum(dioout) dioout, sum(noorderout) noorderout, sum(borrowin) borrowin, sum(borrowout) borrowout from " .
               "ima_file, $sql b where substr(ima06,1,1)!='9' and code=ima01 group by (ima01,ima02,ima25) order by ima01 ";       
      }                                                                                                                                           
      
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
              <td><?=$row2["IMA25"];?></td>   
		          <td style="text-align:right"><?=number_format($row2["IMK09"], 2);?></td>
              <td style="text-align:right"><?=number_format($row2["NORMALIN"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["BACKIN"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["ZAIN"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["DIOIN"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["BORROWIN"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["NORMALIN"]+$row2["BACKIN"]+$row2["ZAIN"]+$row2["DIOIN"]+$row2["BORROWIN"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["NORMALOUT"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["ZAOUT"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["DIOOUT"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["NOORDEROUT"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["BORROWOUT"], 2);?></td>      
              <td style="text-align:right"><?=number_format($row2["NORMALOUT"]+$row2["ZAOUT"]+$row2["DIOOUT"]+$row2["NOORDEROUT"]+$row2["BORROWOUT"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["IMK09"]+$row2["NORMALIN"]+$row2["BACKIN"]+$row2["ZAIN"]+$row2["DIOIN"]+$row2["BORROWIN"]-$row2["NORMALOUT"]-$row2["ZAOUT"]-$row2["DIOOUT"]-$row2["NOORDEROUT"]-$row2["BORROWOUT"], 2);?></td> 
          </tr>
		  <?
			}
      ?>     
</table>    