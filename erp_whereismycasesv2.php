<?php
  session_start();
  $pagetitle = "業務部 &raquo; Where is the CASE v2";
  include("_data.php");
  auth("erp_whereismycasesv2.php");

  $bdate = $_GET['bdate'] ? $_GET['bdate'] : date('Y-m-d', strtotime("-60 days"));

  function findcasewithrxno($rxno, $erp_conn1, $erp_conn2, $bdate){
      //檢查vd110有無訂單 
      $msg='';
      $soea110="select oea01, to_char(oea02,'mm-dd-yyyy') oea02, ta_oea004, occ01, occ02, ta_oea006 from oea_file, occ_file where ta_oea006 like '%$rxno%' " .
               "and oea04=occ01 and oea02 > to_date('$bdate','yyyy-mm-dd') ";
      $erp_sqloea110 = oci_parse($erp_conn1,$soea110 );
      oci_execute($erp_sqloea110);  
      $nrowoea110 = oci_fetch_all($erp_sqloea110, $results);    
      for ($i = 0; $i < $nrowoea110; $i++) {   //vd110可能有N筆相同的訂單
          $oea01=$results['OEA01'][$i]; //order no  
          $prefix = "RX:" . $results['TA_OEA006'][$i] . '(訂單:'.$oea01.')' . $results['OCC02'][$i] . " 於 " . $results['OEA02'][$i] . " 錄入, ";
          $msg .= findcasewithoea01($oea01, $prefix, $erp_conn1,$erp_conn2);     
      } 
      return($msg);
  }        

  function findcasewithoea01($oea01, $prefix, $erp_conn1, $erp_conn2){
      //檢查有無出貨  出貨一定一起出貨 所以只要判斷有無訂單號
      $soga="select to_char(oga02,'mm-dd-yyyy') oga02 from oga_file where oga16='$oea01'";
      $erp_sqloga = oci_parse($erp_conn1,$soga );
      oci_execute($erp_sqloga);  
      $rowoga = oci_fetch_array($erp_sqloga, OCI_ASSOC);
      $msg='';
      if (!is_null($rowoga['OGA02'])) {
          $msg=$prefix . " 於 " . $rowoga['OGA02'] ." 出貨完畢!!<br>";   
      } else { 
          //檢查有無工單號  , 配件不用
          $ssfb="select sfb01, ima01, ima02, gem02 from sfb_file, ima_file, gem_file where sfb22='$oea01' and sfb05=ima01 and sfb82=gem01  ";
          $erp_sqlsfb = oci_parse($erp_conn1,$ssfb );
          oci_execute($erp_sqlsfb);  
          $nrowsfb = oci_fetch_all($erp_sqlsfb, $resultsfb); 
          if ($nrowsfb==0){
              $msg=$prefix . "但未轉工單, 請執行 asfp304 轉派工單<br>";
          } else {
              for ($i = 0; $i < $nrowsfb; $i++) {   //可能有N筆相同的訂單的工單
                  $sfb01=$resultsfb['SFB01'][$i];//工單號
                  $gem02=$resultsfb['GEM02'][$i];//製處     
                  $product = $resultsfb['IMA01'][$i] . " " . $resultsfb['IMA02'][$i] ;
                  $where   = findcasewithsfb01($sfb01, $gem02, $erp_conn1, $erp_conn2); 
                  $msga[$i] = $prefix . " " . $product ." " . $where . "<br>"; 
              }       
              for($i=0; $i<count($msga); $i++){  
                $msg .= $msga[$i];
              }  
          }
      }
      return($msg);
  }  
                            // 工單號 製處。 vd110.      vd210
  function findcasewithsfb01($sfb01, $gem02, $erp_conn1, $erp_conn2){ 
      $msg='';
      //沒有出貨單記錄 檢查秤重
      $stcoga="select to_char(tc_oga002,'mm-dd-yyyy') tc_oga002 from tc_ogb_file,tc_oga_file where tc_ogb002='$sfb01' and tc_ogb001=tc_oga001 ";
      $erp_sqltcoga = oci_parse($erp_conn1,$stcoga );
      oci_execute($erp_sqltcoga);  
      $rowtcoga = oci_fetch_array($erp_sqltcoga, OCI_ASSOC);
      if (!is_null($rowtcoga['TC_OGA002'])) {
          $msg=" 於 " . $rowtcoga['TC_OGA002'] ." 掃描秤重, 但未產生出貨單, 請執行 csft998 中的 生成出貨單 功能.";  
      } else {
          //未掃描秤重。union tc_srg + tc_ohf 找到最后一道的位置 如果都沒 就表示沒進站
          $sohf = "select * from 
                      (select * from 
                        (select ecd01, ecd02, to_char(tc_srg007, 'mm-dd-yyyy') indate, tc_srg008 intime, to_char(tc_srg010,'mm-dd-yyyy') outdate, tc_srg011 outtime from tc_srg_file, ecd_file 
                         where tc_srg001='$sfb01' and tc_srg030=ecd01 and tc_srg007 is not null)  
                        union all 
                        (select ecd01, ecd02, to_char(tc_ohf004, 'mm-dd-yyyy') indate, tc_ohf005 intime, to_char(tc_ohf008,'mm-dd-yyyy') outdate, tc_ohf009 outtime from tc_ohf_file, ecd_file 
                         where tc_ohf001='$sfb01' and tc_ohf015=ecd01 )  
                      )
                  order by indate desc, intime desc ";
          $erp_sqlohf = oci_parse($erp_conn1, $sohf);
          oci_execute($erp_sqlohf);
          $rowohf = oci_fetch_array($erp_sqlohf, OCI_ASSOC); //取出第一筆
          if (!is_null($rowohf["ECD01"])){
              if (is_null($rowohf["OUTDATE"])){
                  $msg = $rowohf["ECD02"] . ' ' . $rowohf["INDATE"] . ' '. $rowohf["INTIME"] . " 入站未出站";
              } else {
                  $msg = $rowohf["ECD02"] . ' ' . $rowohf["OUTDATE"] . ' '. $rowohf["OUTTIME"]  ." 已出站";
              }
          } else {
              $msg = "尚未有任何報工記錄";
          }                  
      }   
      return ($msg);
  } 

  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>WHERE IS MY CASE!! </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
		<td bgcolor="#eeeeee">
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value="<?=$bdate;?>" onfocus="new WdatePicker()">之後
        </td>
		<td bgcolor="#eeeeee">
         RX #:   
        <input name="rxno" type="text" id="rxno" size="70" value="<?=$_GET['rxno'];?>">    (若有多個 RX#  請用 , 隔開 )           
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
            $msg.=findcasewithrxno($rxno,$erp_conn1,$erp_conn2, $bdate);
        }   
      ?>
      <tr bgcolor="#ffffff">
        <td><?=$msg;?></td>      
      </tr>  
</table>   
