<?php
  session_start();
  $pagetitle = "業務部 &raquo; Where is the CASE";
  include("_data.php");
  auth("whereismycases.php");  
  
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
<p>WHERE IS MY CASE!! </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
         RX #:   
        <input name="rxno" type="text" id="rxno" size="150" value="<?=$_GET['rxno'];?>">    (若有多個 RX#  請用 , 隔開 )           
        <input type="submit" name="submit" id="submit" value="查詢">  &nbsp;&nbsp;   &nbsp;&nbsp;          
        </td>
      </tr>
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
            //檢查VD210有無訂單      
            $s="select oea01 from oea_file where ta_oea006='$rxno'";
            $erp_sql = oci_parse($erp_conn2,$s );
            oci_execute($erp_sql);  
            $row = oci_fetch_array($erp_sql, OCI_ASSOC);
            if (is_null($row['OEA01'])) {
                $msg.= $rxno . " VD210 尚未有多角貿易訂單.<br>";  
            } else {
                //檢查VD110有無訂單
                $s="select oea01 from oea_file where ta_oea006='$rxno'";
                $erp_sql = oci_parse($erp_conn,$s );
                oci_execute($erp_sql);  
                $row = oci_fetch_array($erp_sql, OCI_ASSOC);
                if (is_null($row['OEA01'])) {               
                    $msg.= $rxno . " VD210 有訂單,但尚未審核 抛轉至VD110中.<br>"; 
                } else {
                    //檢查VD110 有無抛轉工單
                    $s="select sfb01 from sfb_file where sfbud02='$rxno'";
                    $erp_sql = oci_parse($erp_conn,$s );
                    oci_execute($erp_sql);  
                    $row = oci_fetch_array($erp_sql, OCI_ASSOC);
                    if (is_null($row['SFB01'])) { 
                        $msg.= $rxno . " VD110 有訂單,但尚未抛轉工單, 請執行 asfp304 訂單抛轉工單.<br>"; 
                    } else {
                        //檢查有無在製量
                        $s="select * from tc_srg_file where tc_srg002='$rxno' and  tc_srg005 is not null ";  
                        $erp_sql = oci_parse($erp_conn,$s );
                        oci_execute($erp_sql);  
                        $row = oci_fetch_array($erp_sql, OCI_ASSOC);
                        //無在製量
                        if (is_null($row['TC_SRG002'])) {
                            //檢查有無入庫  
                            $s="select to_char(sfu02,'yyyy/mm/dd') sfu02 from sfb_file, sfu_file, sfv_file " . 
                               "where sfbud02='$rxno' and sfb01=sfv11 and sfv01=sfu01 ";
                            $erp_sql = oci_parse($erp_conn,$s );
                            oci_execute($erp_sql);  
                            $row = oci_fetch_array($erp_sql, OCI_ASSOC);
                            if (is_null($row['SFU02'])) {
                                $msg .= $rxno . " 已製作完畢, 但尚未入庫, 請執行 asft620 入庫.<br>";  
                            } else {
                                //檢查有無秤重 但未產生出貨單
                                $s="select to_char(tc_oga02,'yyyy/mm/dd') tc_oga02 from sfb_file, tc_oga_file,tc_ogb_file " . 
                                   "where sfbud02='$rxno' and sfb01=tc_ogb002 and tc_ogb001=tc_oga001 ";
                                $erp_sql = oci_parse($erp_conn,$s );
                                oci_execute($erp_sql);  
                                $row = oci_fetch_array($erp_sql, OCI_ASSOC);
                                if (is_null($row['TC_OGA02'])) {
                                    $msg.=$rxno . " 已於 " . $row['SFU02'] ." 入庫完畢, 但未秤重, 請執行 asft620 中的 審核及庫存過帳 功能.<br>";   
                                } else {
                                    //檢查有秤重 但未產生出貨單
                                    $s="select to_char(tc_oga02,'yyyy/mm/dd') tc_oga02 from sfb_file, tc_oga_file,tc_ogb_file " . 
                                       "where sfbud02='$rxno' and sfb01=tc_ogb002 and tc_ogb001=tc_oga001 ";
                                    $erp_sql = oci_parse($erp_conn,$s );
                                    oci_execute($erp_sql);  
                                    $row = oci_fetch_array($erp_sql, OCI_ASSOC);
                                    if (is_null($row['TC_OGA02'])) {
                                        $msg.=$rxno . " 已於 " . $row['TC_OGA02'] ." 秤重完畢, 但未產生出貨單, 請執行 csft998 中的 生成出貨單 功能.<br>"; 
                                    } else {
                                        $s="select to_char(oga02,'yyyy/mm/dd') oga02 from oga_file where ogaud02='$rxno'";
                                        $erp_sql = oci_parse($erp_conn,$s );
                                        oci_execute($erp_sql);  
                                        $row = oci_fetch_array($erp_sql, OCI_ASSOC);
                                        if (is_null($row['OGA02'])) {
                                            $msg.=$rxno . " 有秤重, 有出貨單, 但無記錄, 請連絡資訊室檢查!!<br>";
                                        } else {
                                            $msg.=$rxno . " 已於 " . $row['OGA02'] ." 出貨完畢!!<br>"; 
                                        }
                                    } 
                                }   
                            }   
                        } else {
                        //有在製量 判斷在哪裡
                            $ecb01=$row['TC_SRG003'];
                            $ecb03=$row['TC_SRG004'];
                            $gen01=$row['TC_SRG009'];   
                            //查出哪一個製處
                            $s="select gem02 from gen_file, gem_file where gen01='$gen01' and gen03=gem01 ";
                            $erp_sql = oci_parse($erp_conn,$s );
                            oci_execute($erp_sql);  
                            $row3 = oci_fetch_array($erp_sql, OCI_ASSOC);  
                            //若無製處 表示本站未入站 找到最近的一站 檢查它的製作人工序
                            if (is_null($row3['GEM02']))  {                                         
                                $s="select tc_srg009 from tc_srg_file where tc_srg002='$rxno' and tc_srg004 < '$ecb03' "; 
                                $erp_sql = oci_parse($erp_conn,$s );
                                oci_execute($erp_sql);  
                                $row4 = oci_fetch_array($erp_sql, OCI_ASSOC); 
                                if (is_null($row4['TC_SRG009'])) {
                                } else{
                                    $gen01=$row4['TC_SRG009']; 
                                    $s="select gem02 from gen_file, gem_file where gen01= $gen01 and gen03=gem01 ";
                                    $erp_sql = oci_parse($erp_conn,$s );
                                    oci_execute($erp_sql);  
                                    $row5 = oci_fetch_array($erp_sql, OCI_ASSOC); 
                                    $gem02=$row5['GEM02']; 
                                }
                            } else {
                                $gem02=$row3['GEM02']; 
                            }
                            //找出在哪道工序
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
                    }  
                } 
            }   
        }   
      ?>
      <tr bgcolor="#ffffff">
        <td><?=$msg;?></td>      
      </tr>  
</table>   
