<?php
  session_start();
  $pagetitle = "業務部 &raquo; Where is the CASE";
  include("_data.php");
  auth("whereismycasesv1.php");  
  
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
<p>由出貨查起!! </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
        RX #:   
        <input name="rxno" type="text" id="rxno" size="100" value="<?=$_GET['rxno'];?>">    (若有多個 RX#  請用 , 隔開 )           
        <input type="submit" name="submit" id="submit" value="查詢">  &nbsp;&nbsp;   &nbsp;&nbsp;          
      </td></tr>
    </table>
  </div>
</form>
<? if (is_null($_GET['submit'])) die ; ?>   
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">   
      <?
        $rxarray=explode(',', $_GET['rxno']);
        $max=count($rxarray);
        $msg='';
        for($i=0; $i<$max; $i++){
            $rxno=$rxarray[$i]; 
            //檢查有無出貨
            $s="select to_char(oga02,'yyyy/mm/dd') oga02 from oga_file where ogaud02='$rxno'";
            $erp_sql = oci_parse($erp_conn,$s );
            $status=oci_execute($erp_sql);  
            $row = oci_fetch_array($erp_sql, OCI_ASSOC);
            if (is_null($row['OGA02'])) {
              //檢查有秤重 但未產生出貨單
              $s="select to_char(tc_oga002,'yyyy/mm/dd') tc_oga002 from sfb_file, tc_oga_file,tc_ogb_file " . 
                 "where sfbud02='$rxno' and sfb01=tc_ogb002 and tc_ogb001=tc_oga001 ";
              $erp_sql = oci_parse($erp_conn,$s );
              oci_execute($erp_sql);  
              $row = oci_fetch_array($erp_sql, OCI_ASSOC);
              if (is_null($row['TC_OGA002'])) {
                //檢查有無入庫  
                $s="select to_char(sfu02,'yyyy/mm/dd') sfu02 from sfb_file, sfu_file, sfv_file " . 
                   "where sfbud02='$rxno' and sfb01=sfv11 and sfv01=sfu01 ";
                $erp_sql = oci_parse($erp_conn,$s );
                oci_execute($erp_sql);  
                $row = oci_fetch_array($erp_sql, OCI_ASSOC);
                if (is_null($row['SFU02'])) {
                  //判斷在哪一道工序
                  $s="select * from tc_srg_file where tc_srg002='$rxno' and  tc_srg005 is not null ";
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
                    
                      $s="select gem02 from gen_file, gem_file where gen01='$gen01' and gen03=gem01 ";
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
                $msg.=$rxno . " 已於 " . $row['TC_OGA002'] ." 秤重完畢, 但未產生出貨單, 請執行 csft998 中的 生成出貨單 功能.<br>";  
              } 
            } else {
                $msg.=$rxno . " 已於 " . $row['OGA02'] ." 出貨完畢!!<br>";                
            }
        }
           
      ?>
      <tr bgcolor="#ffffff">
        <td><?=$msg;?></td>      
      </tr>  
</table>   
