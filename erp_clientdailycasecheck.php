<?php
  session_start();
  $pagetitle = "業務部 &raquo; 每日到貨分類清單檢查";
  include("_data.php");
  include("_erp.php");
  // auth("erp_clientdailycasecheck.php");  
  
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
  
  //客戶   
  if ($_GET['occ01']!=''){
    //$occfilter=" and occ01 in (select occ01 from occ_file where occ07='" . $_GET['occ01'] . "') ";
    $occfilter=" and occ01 = '" . $_GET['occ01'] . "' ";
  } else {
    $occfilter='';
  }
  
  //新/重/修
  if ($_GET['taoea004']!='') {
    $taoeafilter=" and ta_oea004='" . $_GET['taoea004'] . "' ";
  } else {
    $taoeafilter='';
  }

  //製處
  if ($_GET['gem01']!='') {
    $gemfilter=" and sfb82='" . $_GET['gem01'] . "' ";
  } else {
    $gemfilter='';
  }
  
  //產品分類
  
  if ($_GET['azf01']!='') {
    $azf01=$_GET['azf01'];
    if (strlen($azf01)==2) {
        $azffilter=" and substr(ima10,1,2)='$azf01' "; 
    } else if (strlen($azf01)==3) {
        $azffilter=" and substr(ima10,1,3)='$azf01' "; 
    } else {
        $azffilter=" and ima10='$azf01' ";
    }     
  } else {
    $azffilter='';
  }

  if ($_GET["submit"]=="NOSD匯出") {   
    
    $filename='templates/nosddailycases.xls';
        
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
           
    $y=2; 
    $i=1;        
    $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);  
    $edate1=substr($edate,2,2).substr($edate,5,2).substr($edate,8,2);
      
    $s2="select to_char(oea02,'mm/dd/yy') oea02, sfb01, sfb05, ima1002, oea01, ta_oea004, ta_oea006, ta_oea046, ta_oea047, ta_oea048, sfb08, sfb82, gem02
         from sfb_file, oea_file, ima_file, occ_file, gem_file 
         where sfb22=oea01 and (oea02 between to_date('$bdate1','yy/mm/dd') and to_date('$edate1','yy/mm/dd')) and sfb82=gem01 
         and sfb05=ima01 and oea04=occ01 $occfilter $gemfilter order by oea02, ta_oea006 ";
    $erp_sql2 = oci_parse($erp_conn,$s2 );
    oci_execute($erp_sql2);   
    while ($row = oci_fetch_array($erp_sql2, OCI_ASSOC)) {     
        if ($row['TA_OEA004']=='1') {
          $taoea='New';
        } else if ($row['TA_OEA004']=='2') {
          $taoea='Redo'; 
        } else if ($row['TA_OEA004']=='3') {
          $taoea='Modify';
        } else {
          $taoea='';
        } 
        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A'. $y, $i)
                    ->setCellValue('B'. $y, $row["OEA02"])
                    ->setCellValue('C'. $y, $row["TA_OEA006"]. ' ')
                    ->setCellValue('D'. $y, $row["IMA1002"])   
                    ->setCellValue('E'. $y, $row["SFB08"])  
                    ->setCellValue('F'. $y, $taoea)
                    ->setCellValue('G'. $y, $row["TA_OEA046"]. $row["TA_OEA047"].$row["TA_OEA048"]);
        $y++;
        $i++;
    }   
    $objPHPExcel->setActiveSheetIndex(0)                 
                ->setCellValue('A'. $y, '-- end of data --');                 
    
 
    // Rename sheet
    $objPHPExcel->getActiveSheet()->setTitle('Daily Cases');
                                                            
    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
    $objPHPExcel->setActiveSheetIndex(0);

    // Redirect output to a client’s web browser (Excel5)
    header('Content-Type: application/vnd.ms-excel');  
    header('Content-Disposition: attachment;filename="dailaycases_' . $bdate .'_' . $edate . '.xls"');   
    header('Cache-Control: max-age=0'); 
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');  
    $objWriter->save('php://output'); 
    exit;    
  }           
   
  if ($_GET["submit"]=="一般匯出") {   
    
    $filename='templates/normaldailycases.xls';
        
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
           
    $y=2; 
    $i=1;        
    $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);  
    $edate1=substr($edate,2,2).substr($edate,5,2).substr($edate,8,2);
      
    $s2="select to_char(oea02,'mm/dd/yy') oea02, sfb01, sfb05, ima02, oea01, ta_oea002, ta_oea004, ta_oea006,  ta_oea046, ta_oea047, ta_oea048, sfb08, occ01, occ02, azf01, azf03, gem01, ta_oea001, ta_oea003, gem02 
        from sfb_file, oea_file, ima_file, occ_file, azf_file, gem_file
        where sfb22=oea01 and (oea02 between to_date('$bdate1','yy/mm/dd') and to_date('$edate1','yy/mm/dd')) and sfb82=gem01 
        and sfb05=ima01 and ima10=azf01 and oea04=occ01 $azffilter $occfilter $taoeafilter $gemfilter order by oea02, ta_oea006, sfb01, sfb05 ";
    $erp_sql2 = oci_parse($erp_conn,$s2 );
    oci_execute($erp_sql2);   
    while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {  
        if ($row2['TA_OEA004']=='1') {
          $taoea='新做';
        } else if ($row2['TA_OEA004']=='2') {
          $taoea='重做'; 
        } else if ($row2['TA_OEA004']=='3') {
          $taoea='修改';
        } else {
          $taoea='錯誤';
        };
        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A'. $y, $i)
                    ->setCellValue('B'. $y, $row2["OEA02"])
                    ->setCellValue('C'. $y, $row2["TA_OEA006"])
                    ->setCellValue('D'. $y, $row2["SFB01"])   
                    ->setCellValue('E'. $y, $row2["OEA01"])   
                    ->setCellValue('F'. $y, $row2["SFB05"])    
                    ->setCellValue('G'. $y, $row2["IMA02"])  
                    ->setCellValue('H'. $y, $row2["SFB08"])  
                    ->setCellValue('I'. $y, $taoea)   
                    ->setCellValue('J'. $y, $row2["AZF01"])   
                    ->setCellValue('K'. $y, $row2["AZF03"])   
                    ->setCellValue('L'. $y, $row2["OCC01"])   
                    ->setCellValue('M'. $y, $row2["OCC02"])   
                    ->setCellValue('N'. $y, $row2['TA_OEA001'])
                    ->setCellValue('O'. $y, $row2['TA_OEA003'])
                    ->setCellValue('P'. $y, $row2['GEM01'])
                    ->setCellValue('Q'. $y, $row2['GEM02'])
                    ->setCellValue('R'. $y, $row2["TA_OEA046"]. $row2["TA_OEA047"]. $row2["TA_OEA048"])
                    ->setCellValue('S'. $y, $row2['TA_OEA002']);
        $y++;
        $i++;
    }   
    $objPHPExcel->setActiveSheetIndex(0)                 
                ->setCellValue('A'. $y, '-- end of data --');                 
    
 
    // Rename sheet
    $objPHPExcel->getActiveSheet()->setTitle('Daily Cases');
                                                            
    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
    $objPHPExcel->setActiveSheetIndex(0);

    // Redirect output to a client’s web browser (Excel5)
    header('Content-Type: application/vnd.ms-excel');  
    header('Content-Disposition: attachment;filename="dailaycases_' . $bdate .'_' . $edate . '.xls"');   
    header('Cache-Control: max-age=0'); 
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');  
    $objWriter->save('php://output'); 
    exit;    
  }   
  
  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>客戶每日到貨分類顆數檢查 </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
         到貨日期:   
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> ~~
        <input name="edate" type="text" id="edate" size="12" maxlength="12" value=<?=$edate;?> onfocus="new WdatePicker()"> &nbsp;&nbsp;    
        客戶:
        <select name="occ01" id="occ01">  
            <option value=''>全部</option>
            <?
              //$s1= "select distinct a.occ07 occ07, b.occ02 occ02 from occ_file a , occ_file b where a.occ07=b.occ01 order by a.occ07 ";
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
        業績分類:
        <select name="azf01" id="azf01">  
            <option value=''>全部</option>
            <option value='931' <? if ($_GET['azf01']=='931') echo " selected";?> >一般烤瓷牙小計</option> 
            <option value='932' <? if ($_GET['azf01']=='932') echo " selected";?>>Captek小計</option> 
            <option value='933' <? if ($_GET['azf01']=='933') echo " selected";?>>鋼牙小計</option> 
            <option value='934' <? if ($_GET['azf01']=='934') echo " selected";?>>全瓷小計</option> 
            <option value='93'  <? if ($_GET['azf01']=='93')  echo " selected";?>>固定小計</option> 
            <option value='941' <? if ($_GET['azf01']=='941') echo " selected";?>>金屬床小計</option> 
            <option value='942' <? if ($_GET['azf01']=='942') echo " selected";?>>彈性床小計</option>   
            <option value='943' <? if ($_GET['azf01']=='943') echo " selected";?>>樹脂床小計</option>   
            <option value='944' <? if ($_GET['azf01']=='944') echo " selected";?>>活動臨時假牙小計</option>   
            <option value='945' <? if ($_GET['azf01']=='945') echo " selected";?>>咬合板小計</option>   
            <option value='946' <? if ($_GET['azf01']=='946') echo " selected";?>>透明支架小計</option>  
            <option value='94'  <? if ($_GET['azf01']=='94')  echo " selected";?>>活動小計</option>    
            <?
              $s1= "select azf01, azf03 from azf_file where azf02='E' and azfacti='Y' order by azf01 ";
              $erp_sql1 = oci_parse($erp_conn,$s1 );
              oci_execute($erp_sql1);  
              while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                  echo "<option value=" . $row1["AZF01"];  
                  if ($_GET["azf01"] == $row1["AZF01"]) echo " selected";                  
                  echo ">" . $row1['AZF01'] ."--" .$row1["AZF03"] . "</option>"; 
              }   
            ?>
        </select>
        部門別:
        <select name="gem01" id="gem01">  
            <option value=''>全部</option>
            <?
              $s1= "select gem01, gem02 from gem_file where (substr(gem01,1,2)='69' or substr(gem01,1,2)='6A') and substr(gem01,4,3)='000' and gemacti='Y' order by gem01 ";
              $erp_sql1 = oci_parse($erp_conn,$s1 );
              oci_execute($erp_sql1);  
              while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                  echo "<option value=" . $row1["GEM01"];  
                  if ($_GET["gem01"] == $row1["GEM01"]) echo " selected";                  
                  echo ">" . $row1['GEM01'] ."--" .$row1["GEM02"] . "</option>"; 
              }   
            ?>
        </select>
        新做/重做/修改:
        <select name="taoea004" id="taoea004">  
            <option value=''  <? if ($_GET['taoea004']=='')  echo " selected";?>>全部</option> 
            <option value='2' <? if ($_GET['taoea004']=='2') echo " selected";?>>重做</option> 
            <option value='3' <? if ($_GET['taoea004']=='3') echo " selected";?>>修改</option> 
        </select>
        &nbsp;&nbsp;      
        <input type="submit" name="submit" id="submit" value="查詢">  &nbsp;&nbsp;   &nbsp;&nbsp;             
        <input type="submit" name="submit" id="submit" value="NOSD匯出">  &nbsp;&nbsp;   &nbsp;&nbsp; 
        <input type="submit" name="submit" id="submit" value="一般匯出">  &nbsp;&nbsp;   &nbsp;&nbsp; 
      </td></tr>
    </table>
  </div>
</form>


<? if (is_null($_GET['submit'])) die ; ?>   
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
        <th width="16">&nbsp;</th> 
        <th>到貨日期</th>  
        <th>RX #</th>    
        <th>工單編號</th>     
        <th>訂單編號</th>
        <th>產品編號</th>  
        <th>產品名稱</th> 
        <th>顆/床數</th> 
        <th>新/重/修</th> 
        <th>分類代號</th> 
        <th>分類名稱</th> 
        <th>客戶代號</th>   
        <th>客戶名稱</th> 
        <th>診所</th>
        <th>患者</th> 
        <th>製處</th>  
        <th>比色</th>
    </tr>
    <?
      $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);  
      $edate1=substr($edate,2,2).substr($edate,5,2).substr($edate,8,2);
      
      $s2="select to_char(oea02,'mm/dd/yy') oea02, sfb01, sfb05, ima02, oea01, ta_oea004, ta_oea006,ta_oea046, ta_oea047, ta_oea048, sfb08, occ01, occ02, azf01, azf03, gem01, ta_oea001, ta_oea003, gem02 
           from sfb_file, oea_file, ima_file, occ_file, azf_file, gem_file
           where sfb22=oea01 and (oea02 between to_date('$bdate1','yy/mm/dd') and to_date('$edate1','yy/mm/dd')) and sfb82=gem01 
           and sfb05=ima01 and ima10=azf01 and oea04=occ01 $azffilter $occfilter $taoeafilter $gemfilter order by oea02, ta_oea006, sfb01, sfb05 ";
      $erp_sql2 = oci_parse($erp_conn,$s2 );
      oci_execute($erp_sql2);  
      $bgkleur = "ffffff";  
      $i=0;      
      $total=0;      
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {     
          $i++;
          $total+=$row2['SFB08'];
          if ($row2['TA_OEA004']=='1') {
            $taoea='新做';
          } else if ($row2['TA_OEA004']=='2') {
            $taoea='重做'; 
          } else if ($row2['TA_OEA004']=='3') {
            $taoea='修改';
          } else {
            $taoea='錯誤';
          }
          ?>
            <tr bgcolor="#<?=$bgkleur;?>">
                <td><?=$i;?></td>      
                <td><?=$row2['OEA02'];?></td>  
                <td><?=$row2['TA_OEA006'];?></td>
                <td><?=$row2['SFB01'];?></td> 
                <td><?=$row2['OEA01'];?></td> 
                <td><?=$row2['SFB05'];?></td>   
                <td><?=$row2["IMA02"];?></td>  
                <td><?=$row2["SFB08"];?></td>  
                <td><?=$taoea;?></td>  
                <td><?=$row2["AZF01"];?></td> 
                <td><?=$row2["AZF03"];?></td>
                <td><?=$row2["OCC01"];?></td> 
                <td><?=$row2["OCC02"];?></td>
                <td><?=$row2["TA_OEA001"];?></td>
                <td><?=$row2["TA_OEA003"];?></td>  
                <td><?=$row2['GEM01'].'  '.$row2["GEM02"];?></td> 
                <td><?=$row2["TA_OEA046"]. $row2["TA_OEA047"]. $row2["TA_OEA048"];;?></td>
            </tr>
          <?   
      }
      ?>   
      <tr bgcolor="#<?=$bgkleur;?>">
          <td colspan="4">&nbsp;</td> 
          <td><?=$total;?></td>  
          <td colspan="3">&nbsp;</td>      
      </tr> 
</table>   
