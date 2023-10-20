<?php
  session_start();
  $pagetitle = "車間作業 &raquo; 期間內未刷卡工單";
  include("_data.php");   
  //auth("erp_casedailynocheckin.php");  
  
  if (is_null($_GET['bdate'])) {
    $bdate=date('Y-m-d', strtotime("-1 days"));   
  } else {
    $bdate=$_GET['bdate'];
  }
                       

  if ($_GET["submit"]=="匯出") {   
    $filename='templates/casenocheckin.xls';
        
    error_reporting(E_NONE);  
    require_once 'classes/PHPExcel.php'; 
    require_once 'classes/PHPExcel/IOFactory.php';  
    $objReader = PHPExcel_IOFactory::createReader('Excel5');
    $objPHPExcel = $objReader->load($filename);  
    $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
    $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4); 
    // Set properties
    $objPHPExcel ->getProperties()->setCreator('Frank' )
                 ->setLastModifiedBy('Frank')
                 ->setTitle('Frank')
                 ->setSubject('Frank')
                 ->setDescription('Frank')
                 ->setKeywords('Frank')
                 ->setCategory('Frank');                           
                      
    $y=4;   
    $s2= "select gem02, sfb01, sfb22, sfbud02, occ02, ima02 from
            (select sfb01, sfb22, sfb05, sfb82, SFBUD02 from sfb_file
            where  sfb01 in ( select tc_ogb002 from tc_ogb_file, tc_oga_file where tc_ogb001=tc_oga001 and tc_oga002 = to_date('$bdate','yy/mm/dd') )
            and sfb01 not in ( select tc_srg001 from tc_srg_file where tc_srg007 is not null or tc_srg010 is not null or tc_srg013 is not null)
            and sfb82 !='6AZ000'
            and sfb05 not like '1Z%' 
            and sfb05 not like '2Z%' ), gem_file, oea_file, occ_file, ima_file
            where sfb82=gem01 
            and sfb22=oea01 and ta_oea004='1'
            and oea04=occ01
            and sfb05=ima01
            order by gem01" ;
    $erp_sql2 = oci_parse($erp_conn,$s2 );
    oci_execute($erp_sql2);   
    while ($row = oci_fetch_array($erp_sql2, OCI_ASSOC)) {     
          $objPHPExcel->setActiveSheetIndex(0)
                      ->setCellValue('A'. $y, $bdate)
                      ->setCellValue('B'. $y, $row["GEM02"])
                      ->setCellValue('C'. $y, $row["SFB01"])
                      ->setCellValue('D'. $y, $row["SFB22"])   
                      ->setCellValue('E'. $y, $row["SFBUD02"])   
                      ->setCellValue('F'. $y, $row["OCC02"])   
                      ->setCellValue('G'. $y, $row["IMA02"]);   
          $y++;              
    }                      
                       
    // Rename sheet
    $objPHPExcel->getActiveSheet()->setTitle('未刷卡工單');
                                                            
    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
    $objPHPExcel->setActiveSheetIndex(0);

    // Redirect output to a client’s web browser (Excel5)
    header('Content-Type: application/vnd.ms-excel');  
    header('Content-Disposition: attachment;filename="casenocheckin_'.$bdate.'.xls"');   
    header('Cache-Control: max-age=0'); 
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');  
    $objWriter->save('php://output'); 
    exit;    
  } 
  
  
  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>期間內未刷卡工單 </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
         出貨日期:   
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> 
        &nbsp;&nbsp;      
 <!--       製處:   
        <select name="gem01" id="gem01">  
          <option value="">全部製處</option>
        <?
          $s1= "select gem01,gem02 from gem_file where (gem01 like '69%' or gem01 like '6A%') and substr(gem01,3,1)!='0' order by gem01";
          $erp_sql1 = oci_parse($erp_conn1,$s1 );
          oci_execute($erp_sql1);  
          while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
              echo "<option value=" . $row1["GEM01"];  
              if ($_GET["gem01"] == $row1["GEM01"]) echo " selected";                  
              echo ">" . $row1['GEM01'] ."--" .$row1["GEM02"] . "</option>"; 
          }   
        ?>
        </select> &nbsp;  &nbsp; &nbsp;  &nbsp;   -->
        <input type="submit" name="submit" id="submit" value="查詢">  &nbsp;&nbsp;   &nbsp;&nbsp;      
        <input type="submit" name="submit" id="submit" value="匯出">             
      </td></tr>
    </table>
  </div>
</form>


<? if (is_null($_GET['submit'])) die ; ?>   
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
        <th width="16">&nbsp;</th> 
        <th>日期</th>  
        <th>製處</th> 
        <th>工單號</th>  
        <th>訂單號</th>    
        <th>RX #</th>    
        <th>客戶</th>
        <th>產品</th>  
           
    </tr>
    <?
      
      $s2= "select gem02, sfb01, sfb22, sfbud02, occ02, ima02 from
            (select sfb01, sfb22, sfb05, sfb82, SFBUD02 from sfb_file
            where  sfb01 in ( select tc_ogb002 from tc_ogb_file, tc_oga_file where tc_ogb001=tc_oga001 and tc_oga002 = to_date('$bdate','yy/mm/dd') )
            and sfb01 not in ( select tc_srg001 from tc_srg_file where tc_srg007 is not null or tc_srg010 is not null or tc_srg013 is not null)
            and sfb82 !='6AZ000'
            and sfb05 not like '1Z%' 
            and sfb05 not like '2Z%' ), gem_file, oea_file, occ_file, ima_file
            where sfb82=gem01 
            and sfb22=oea01 and ta_oea004='1'
            and oea04=occ01
            and sfb05=ima01
            order by gem01" ;

      $erp_sql2 = oci_parse($erp_conn,$s2 );
      oci_execute($erp_sql2);  
      $bgkleur = "ffffff";  
      $i=0;      
      $total=0;      
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {     
          $i++;                                                   
          ?>
            <tr bgcolor="#<?=$bgkleur;?>">
                <td><?=$i;?></td>      
                <td><?=$row2['GEM02'];?></td>  
                <td><?=$row2['SFB01'];?></td>
                <td><?=$row2['SFB22'];?></td> 
                <td><?=$row2['SFBUD02'];?></td>   
                <td><?=$row2['OCC02'];?></td>  
                <td><?=$row2['IMA02'];?></td>                                                   
            </tr>
          <?   
      }
      ?>   
</table>   