<?php
  session_start();
  $pagetitle = "業務部 &raquo; Delay Report";
  include("_data.php");
  include("_erp.php");
  auth("clientdelayreport.php");  
  
  if (is_null($_GET['bdate'])) {
    $bdate =  date('Y-m-d',strtotime('-3 days'));
  } else {
    $bdate=$_GET['bdate'];
  }                              
  
  if (is_null($_GET['occ01'])) {
    $occ01 =  '';
  } else {
    $occ01=$_GET['occ01'];
  }  
  
    if ($_GET["submit"]=="匯出") {   
        error_reporting(E_ALL);  
        require_once 'classes/PHPExcel.php'; 
        require_once 'classes/PHPExcel/IOFactory.php';  
        $objPHPExcel = new PHPExcel();
        // Set properties
        $objPHPExcel ->getProperties()->setCreator( 'DelayReport')
                     ->setLastModifiedBy('DelayReport')
                     ->setTitle('DelayReport')
                     ->setSubject('DelayReport')
                     ->setDescription('DelayReport')
                     ->setKeywords('DelayReport')
                     ->setCategory('DelayReport');
        
        // Add some data      
                        
        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A1', 'Delay Report');                    
        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A3', 'No.')                   
                    ->setCellValue('B3', 'Case No.')
                    ->setCellValue('C3', 'Received Date')   
                    ->setCellValue('D3', 'Due Date')   
                    ->setCellValue('E3', 'Product Code')  
                    ->setCellValue('F3', 'Product Description')  
                    ->setCellValue('G3', 'Qty.')  
                    ->setCellValue('H3', 'Status')  
                    ->setCellValue('I3', 'Comment'); 

        $y=4; 
        $qtytotal=0;  
        $i=0;   
        
        $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);  
        $s2= "select * from " .
             "(select ta_oea006 rxno, to_char(oea02,'yyyy/mm/dd') indate, to_char(ta_oea005,'yyyy/mm/dd') duedate, oea04, oea01, oea02, oga16 from oea_file left join oga_file on oea01=oga16) ". 
             "where oea02 <= to_date('$bdate1','yymmdd') and oea04='$occ01' and oga16 is null";
        $erp_sql2 = oci_parse($erp_conn,$s2 );
        oci_execute($erp_sql2);  
        while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {   
              $oea01=$row2['OEA01'];        
              $msg='';    
              //$status=whereisthecase($oea01);
              $s="select to_char(oga02,'yyyy/mm/dd') oga02 from oga_file where oga16='$oea01'";
              $erp_sql = oci_parse($erp_conn,$s );
              oci_execute($erp_sql);  
              $row = oci_fetch_array($erp_sql, OCI_ASSOC);
              if (is_null($row['OGA02'])) {
                //檢查有秤重 但未產生出貨單
                $s="select to_char(tc_oga002,'yyyy/mm/dd') tc_oga002 from sfb_file, tc_oga_file,tc_ogb_file " . 
                   "where sfb22='$oea01' and sfb01=tc_ogb002 and tc_ogb001=tc_oga001 ";
                $erp_sql = oci_parse($erp_conn,$s );
                oci_execute($erp_sql);  
                $row = oci_fetch_array($erp_sql, OCI_ASSOC);
                if (is_null($row['TC_OGA002'])) {
                  //檢查有無入庫  
                  $s="select to_char(sfu02,'yyyy/mm/dd') sfu02 from sfb_file, sfu_file, sfv_file " . 
                     "where sfb22='$oea01' and sfb01=sfv11 and sfv01=sfu01 ";
                  $erp_sql = oci_parse($erp_conn,$s );
                  oci_execute($erp_sql);  
                  $row = oci_fetch_array($erp_sql, OCI_ASSOC);
                  if (is_null($row['SFU02'])) {
                    //判斷在哪一道工序   tc_srg_file沒有訂單號碼
                    $s="select * from tc_srg_file, sfb_file where tc_srg005 is not null and tc_srg001=sfb01 and sfb22='$oea01' ";
                    $erp_sql = oci_parse($erp_conn,$s );
                    oci_execute($erp_sql);  
                    $row = oci_fetch_array($erp_sql, OCI_ASSOC);
                    if (is_null($row['TC_SRG001'])) {       //找不到在製量 有可能是有工單已入庫 或無訂單
                      $s="select * from tc_srg_file where tc_srg002='$rxno' "; //找尋有無報工
                      $erp_sql = oci_parse($erp_conn,$s );
                      oci_execute($erp_sql);  
                      $row = oci_fetch_array($erp_sql, OCI_ASSOC);
                      if (is_null($row['TC_SRG001'])) { //無待報工記錄 檢查有無訂單
                        $s="select * from oea_file where ta_oea006='$rxno'";
                        $erp_sql = oci_parse($erp_conn,$s );
                        oci_execute($erp_sql);  
                        $row = oci_fetch_array($erp_sql, OCI_ASSOC);
                        if (is_null($row['OEA02'])) {
                          //檢查VD210有無訂單 但未審核     
                          $s="select * from oea_file where ta_oea006='$rxno'";
                          $erp_sql = oci_parse($erp_conn2,$s );
                          oci_execute($erp_sql);  
                          $row = oci_fetch_array($erp_sql, OCI_ASSOC);
                          if (is_null($row['OEA02'])) {
                            $msg.= $rxno . " VD210 尚未有多角貿易訂單.";  
                          } else {
                            $msg= $rxno . " VD210 有訂單,但尚未審核 抛轉至VD110中.";  
                          }                    
                          
                        } else {
                          $msg=$rxno . " 有訂單但未轉工單, 請執行 asfp304 轉派工單.";                      
                        }  
                      } else { 
                        $msg=$rxno . " 已於入庫完畢, 但未秤重, 請執行 asft620 中的 審核及庫存過帳 功能."; 
                      }        
                    } else {    //找到在製量 判斷在哪裡
                      $ecb01=$row['TC_SRG003'];
                      $ecb03=$row['TC_SRG004'];
                      $tc_srg010=$row['TC_SRG010'];  
                      $s="select ecb06, ecb17 from ecb_file where ecb01='$ecb01' and ecb03='$ecb03' ";
                      $erp_sql = oci_parse($erp_conn,$s );
                      oci_execute($erp_sql);  
                      $row1 = oci_fetch_array($erp_sql, OCI_ASSOC);
                      if (is_null($row1['ECB06'])) {
                        $msg = $rxno . " 在車間製作中, 但產品工序設定錯誤, 無法定位, 請洽資訊室.";    
                      } else {
                        $ecb06=$row1['ECB06'];
                        $ecb17=$row1['ECB17'];
                        //查出哪一個製處 若沒有入站的人員工號 要回派工單去找製處
                        if (is_null($row['TC_SRG009'])) {     
                            $sfb01=$row['TC_SRG001'];
                            $s="select sfb82 from sfb_file where sfb01='$sfb01' "; 
                            $erp_sql = oci_parse($erp_conn,$s );
                            oci_execute($erp_sql);  
                            $rowsfb = oci_fetch_array($erp_sql, OCI_ASSOC);
                            $gem01=$rowsfb['SFB82'];  
                            $s="select gem02 from gem_file where gem01='$gem01' "; 
                            $erp_sql = oci_parse($erp_conn,$s );
                            oci_execute($erp_sql);  
                            $rowgem = oci_fetch_array($erp_sql, OCI_ASSOC);
                            $gem02=$rowgem['GEM02']; 
                        } else {                    
                            $gen01=$row['TC_SRG009'];
                            $s="select gem02 from gen_file, gem_file where gen01='$gen01' and gen03=gem01 ";
                            $erp_sql = oci_parse($erp_conn,$s );
                            oci_execute($erp_sql);  
                            $row3 = oci_fetch_array($erp_sql, OCI_ASSOC);  
                            $gem02=$row3['GEM02']; 
                        } 
                                     
                        if (is_null($row['TC_SRG007'])){   //有在製量 沒有入站 表示上站出 本站未進
                          $msg = $rxno . " 在 $gem02 -- $ecb06 -- $ecb17 的上一站刷出 但未入站." ;
                        } else {
                          //判斷出站沒 未出站表示, 但有出站 則表示QC中
                          if (is_null($tc_srg010)) {
                            $msg = $rxno . " 在 $gem02 -- $ecb06 -- $ecb17 製作中." ;
                          } else {                                                   
                            $msg = $rxno . " 在 $gem02 -- $ecb17 PQC中." ; 
                          }
                        }
                      }
                    }                   
                  } else {  
                    $msg=$rxno . " 已於 " . $row['SFU02'] ." 入庫完畢, 但未秤重."; 
                  }
                  
                } else {    
                  $msg=$rxno . " 已於 " . $row['TC_OGA002'] ." 秤重完畢, 但未產生出貨單.";  
                } 
              } else {
                  $msg=$rxno . " 已於 " . $row['OGA02'] ." 出貨完畢!!";                
              } 
                                           
              //使用rxno 去ogb找product
              $s3="select oeb04, oeb06, oeb12, ima1002 from oeb_file, ima_file where oeb01='$oea01' and oeb04=ima01";
              $erp_sql3 = oci_parse($erp_conn,$s3 );
              oci_execute($erp_sql3);
              while ($row3 = oci_fetch_array($erp_sql3, OCI_ASSOC)) {
                  $i++;  
                  $objPHPExcel->setActiveSheetIndex(0)
                              ->setCellValue('A'. $y, $i)
                              ->setCellValue('B'. $y, $row2["RXNO"])
                              ->setCellValue('C'. $y, $row2["INDATE"])   
                              ->setCellValue('D'. $y, $row2["DUEDATE"])   
                              ->setCellValue('E'. $y, $row3["OEB04"])   
                              ->setCellValue('F'. $y, $row3["OEB06"])   
                              ->setCellValue('G'. $y, $row3["OEB12"])  
                              ->setCellValue('H'. $y, $row3["IMA1002"]) 
                              ->setCellValue('I'. $y, $msg); 
                  $y++; 
              }    
        }
        //total
        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A'.($y+1), 'Total');  
                                                   
        $objPHPExcel->setActiveSheetIndex(0)        
                    ->setCellValue('G'.($y+1), '=sum(g3:g' . ($y-1). ")");                 
        
        // Rename sheet
        $objPHPExcel->getActiveSheet()->setTitle('DelayReport');
                                                                
        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $objPHPExcel->setActiveSheetIndex(0);

        // Redirect output to a client’s web browser (Excel5)
        header('Content-Type: application/vnd.ms-excel');  
        header('Content-Disposition: attachment;filename="DelayReport.xls"');    
        header('Cache-Control: max-age=0'); 
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');  
        $objWriter->save('php://output'); 
        exit;    
  }  
  
  function whereisthecase($oea01){ 
      //檢查有無出貨
      $s="select to_char(oga02,'yyyy/mm/dd') oga02 from oga_file where oga16='$oea01'";
      $erp_sql = oci_parse($erp_conn,$s );
      oci_execute($erp_sql);  
      $row = oci_fetch_array($erp_sql, OCI_ASSOC);
      if (is_null($row['OGA02'])) {
        //檢查有秤重 但未產生出貨單
        $s="select to_char(tc_oga02,'yyyy/mm/dd') tc_oga02 from sfb_file, tc_oga_file,tc_ogb_file " . 
           "where sfb22='$oea01' and sfb01=tc_ogb002 and tc_ogb001=tc_oga001 ";
        $erp_sql = oci_parse($erp_conn,$s );
        oci_execute($erp_sql);  
        $row = oci_fetch_array($erp_sql, OCI_ASSOC);
        if (is_null($row['TC_OGA02'])) {
          //檢查有無入庫  
          $s="select to_char(sfu02,'yyyy/mm/dd') sfu02 from sfb_file, sfu_file, sfv_file " . 
             "where sfb22='$oea01' and sfb01=sfv11 and sfv01=sfu01 ";
          $erp_sql = oci_parse($erp_conn,$s );
          oci_execute($erp_sql);  
          $row = oci_fetch_array($erp_sql, OCI_ASSOC);
          if (is_null($row['SFU02'])) {
            //判斷在哪一道工序   tc_srg_file沒有訂單號碼
            $s="select * from tc_srg_file, sfb_file where tc_srg005 is not null and tc_srg001=sfb01 and sfb22='$oea01' ";
            $erp_sql = oci_parse($erp_conn,$s );
            oci_execute($erp_sql);  
            $row = oci_fetch_array($erp_sql, OCI_ASSOC);
            if (is_null($row['TC_SRG001'])) {       //找不到在製量 有可能是有工單已入庫 或無訂單
              $s="select * from tc_srg_file where tc_srg002='$rxno' "; //找尋有無報工
              $erp_sql = oci_parse($erp_conn,$s );
              oci_execute($erp_sql);  
              $row = oci_fetch_array($erp_sql, OCI_ASSOC);
              if (is_null($row['TC_SRG001'])) { //無待報工記錄 檢查有無訂單
                $s="select * from oea_file where ta_oea006='$rxno'";
                $erp_sql = oci_parse($erp_conn,$s );
                oci_execute($erp_sql);  
                $row = oci_fetch_array($erp_sql, OCI_ASSOC);
                if (is_null($row['OEA02'])) {
                  //檢查VD210有無訂單 但未審核     
                  $s="select * from oea_file where ta_oea006='$rxno'";
                  $erp_sql = oci_parse($erp_conn2,$s );
                  oci_execute($erp_sql);  
                  $row = oci_fetch_array($erp_sql, OCI_ASSOC);
                  if (is_null($row['OEA02'])) {
                    $msg.= $rxno . " VD210 尚未有多角貿易訂單.<br>";  
                  } else {
                    $msg.= $rxno . " VD210 有訂單,但尚未審核 抛轉至VD110中.<br>";  
                  }                    
                  
                } else {
                  $msg.=$rxno . " 有訂單但未轉工單, 請執行 asfp304 轉派工單.<br>";                      
                }  
              } else { 
                $msg.=$rxno . " 已於入庫完畢, 但未秤重, 請執行 asft620 中的 審核及庫存過帳 功能.<br>"; 
              }        
            } else {    //找到在製量 判斷在哪裡
              $ecb01=$row['TC_SRG003'];
              $ecb03=$row['TC_SRG004'];
              $tc_srg010=$row['TC_SRG010'];  
              $s="select ecb06, ecb17 from ecb_file where ecb01='$ecb01' and ecb03='$ecb03' ";
              $erp_sql = oci_parse($erp_conn,$s );
              oci_execute($erp_sql);  
              $row1 = oci_fetch_array($erp_sql, OCI_ASSOC);
              if (is_null($row1['ECB06'])) {
                $msg .= $rxno . " 在車間製作中, 但產品工序設定錯誤, 無法定位, 請洽資訊室.<br>";    
              } else {
                $ecb06=$row1['ECB06'];
                $ecb17=$row1['ECB17'];
                
                $gen01=$row['TC_SRG009'];//查出哪一個製處
                $s="select gem02 from gen_file, gem_file where gen01= $gen01 and gen03=gem01 ";
                $erp_sql = oci_parse($erp_conn,$s );
                oci_execute($erp_sql);  
                $row3 = oci_fetch_array($erp_sql, OCI_ASSOC);  
                $gem02=$row3['GEM02'];                     
                if (is_null($row['TC_SRG007'])){   //有在製量 沒有入站 表示上站出 本站未進
                  $msg .= $rxno . " 在 $gem02 -- $ecb06 -- $ecb17 的上一站刷出 但未入站.<br>" ;
                } else {
                  //判斷出站沒 未出站表示, 但有出站 則表示QC中
                  if (is_null($tc_srg010)) {
                    $msg .= $rxno . " 在 $gem02 -- $ecb06 -- $ecb17 製作中.<br>" ;
                  } else {                                                   
                    $msg .= $rxno . " 在 $gem02 -- $ecb17 PQC中.<br>" ; 
                  }
                }
              }
            }                   
          } else {  
            $msg.=$rxno . " 已於 " . $row['SFU02'] ." 入庫完畢, 但未秤重, 請執行 asft620 中的 審核及庫存過帳 功能.<br>"; 
          }
          
        } else {    
          $msg.=$rxno . " 已於 " . $row['TC_OGA02'] ." 秤重完畢, 但未產生出貨單, 請執行 csft998 中的 生成出貨單 功能.<br>";  
        } 
      } else {
          $msg.=$rxno . " 已於 " . $row['OGA02'] ." 出貨完畢!!<br>";                
      }  
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
<p>客戶Delay Report -- 由出貨查起!! </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
         到貨日期:   
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> 
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
        <th>RX #</th>  
        <th>Received Date</th>  
        <th>Due Date</th> 
        <th>Product Code</th>
        <th>Product Description</th>   
        <th>Qty.</th>    
        <th>Status</th>
        <th>Comment</th> 
    </tr>
    <?
      $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);  
      $s2= "select * from " .
           "(select ta_oea006 rxno, to_char(oea02,'yyyy/mm/dd') indate, to_char(ta_oea005,'yyyy/mm/dd') duedate, oea04, oea01, oea02, oga16 from oea_file left join oga_file on oea01=oga16) ". 
           "where oea02 <= to_date('$bdate1','yymmdd') and oea04='$occ01' and oga16 is null order by indate,rxno";
      $erp_sql2 = oci_parse($erp_conn,$s2 );
      oci_execute($erp_sql2);  
      $bgkleur = "ffffff";  
      $i=0;
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {  
          $oea01=$row2['OEA01'];   
          $msg=findcasewithoea01($oea01, $erp_conn1, $erp_conn2);   
          //使用rxno 去ogb找product
          $s3="select oeb04, oeb06, oeb12, ima1002 from oeb_file, ima_file where oeb01='$oea01' and oeb04=ima01";
          $erp_sql3 = oci_parse($erp_conn,$s3 );
          oci_execute($erp_sql3);
          while ($row3 = oci_fetch_array($erp_sql3, OCI_ASSOC)) {
              $i++;                
          ?>
              <tr bgcolor="#<?=$bgkleur;?>">
                  <td><?=$i;?></td>
                  <td><?=$row2["RXNO"];?></td>  
                  <td><?=$row2["INDATE"];?></td>  
                  <td><?=$row2["DUEDATE"];?></td>
                  <td><?=$row3["OEB04"];?></td>  
                  <td><?=$row3["OEB06"];?></td>  
                  <td><?=$row3["OEB12"];?></td>
                  <td><?=$row3["IMA1002"];?></td>  
                  <td><?=$msg;?></td>   
                </tr>
          <?
          }
      }
      ?>    
</table>   