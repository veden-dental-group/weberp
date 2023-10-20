<?php
  session_start();
  $pagetitle = "車間作業 &raquo; 進站報工";
  include("_data.php");
  auth("orders_checkin.php");;
  date_default_timezone_set('Asia/Taipei');
  //for xajax
  require ('xajax/xajax_core/xajax.inc.php');
  $xajax = new xajax();
  $xajax->configure('javascript URI', 'xajax/');

  Function checkinOrders($v){      
      $o=new xajaxResponse(); 
      
      $erp_db_host = "topprod";
      $erp_db_user = "vd11";
      $erp_db_pass = "vd11";  
      $erp_conn = oci_connect($erp_db_user, $erp_db_pass, $erp_db_host,'AL32UTF8'); 
      
      
      $ecb06=$v['ecb06']; 
      $sfb01=$v['sfb01'];  
      $gen01=$v['gen01'];  
      $new='';
      
      $errmsg='';
      $erp_sql = oci_parse($erp_conn, "select ecd02 from ecd_file where ecd01='$ecb06'");
      oci_execute($erp_sql);  
      $row1 = oci_fetch_array($erp_sql, OCI_ASSOC);
      if (is_null($row1['ECD02'])) {
         $errmsg.="查無此作業編號:'$ecb06'!!&nbsp;&nbsp;";
         $o->assign("ecb06","value",'');    
      }
      
      $erp_sql = oci_parse($erp_conn, "select gen02 from gen_file where gen01='$gen01'");
      oci_execute($erp_sql);  
      $row3 = oci_fetch_array($erp_sql, OCI_ASSOC);
      if (is_null($row3['GEN02'])) {
         $errmsg.="查無此員工編號:'$gen01'!!&nbsp;&nbsp;";
         $o->assign("gen01","value",'');  
      }
      
      $s="select sfb01 from sfb_file where sfb01='$sfb01'";
      $erp_sql = oci_parse($erp_conn, $s);
      oci_execute($erp_sql);  
      $row2 = oci_fetch_array($erp_sql, OCI_ASSOC);
      if (is_null($row2['SFB01'])) {
         $errmsg.="查無此工單號:'$sfb01'!!&nbsp;&nbsp;";
         $o->assign("sfb01","value",'');  
      }
                          
      if ($errmsg!=''){ //有錯誤訊息
          $o->assign('orderstatus','innerHTML',$errmsg);    
      } else {
          //檢查本工單有無在製量 若無則不能報工
          $s= "select tc_srg004 from tc_srg_file, ecb_file where tc_srg001='$sfb01' and tc_srg005 is not null and tc_srg003=ecb01 and tc_srg004=ecb03 and ecb06='$ecb06' ";
          $erp_sql = oci_parse($erp_conn,$s );
          oci_execute($erp_sql);  
          $row4 = oci_fetch_array($erp_sql, OCI_ASSOC);
          if (is_null($row4['TC_SRG004'])){
             //要另外查詢目前應該要在哪一道報工
             
            $errmsg='本工序目前尚無法報工!!';    
            $o->assign('orderstatus','innerHTML',$errmsg);    
          } else {
             $erp_sql = oci_parse($erp_conn, "select * from tc_srg_file where tc_srg001='$sfb01' and tc_srg004='". $row4['TC_SRG004'] ."' ");
             oci_execute($erp_sql);  
             $row5 = oci_fetch_array($erp_sql, OCI_ASSOC);
             if ($row5['TC_SRG018']=='N') {  //非返工
                if (is_null($row5['TC_SRG007'])){
                    
                    $today=date('Y-m-d');
                    $time=date('h:i:s');
                    $s="update tc_srg_file set tc_srg007=to_date('$today','YYYY-MM-DD'), tc_srg008='$time', tc_srg009='$gen01' where tc_srg001='$sfb01' and tc_srg004='". $row4['TC_SRG004'] . "' ";
                    $erp_sql = oci_parse($erp_conn, $s );
                    oci_execute($erp_sql);  
                    $new='Y';                                    
                } else {
                    $errmsg='本工序已報過工!!';    
                    $o->assign('orderstatus','innerHTML',$errmsg); 
                } 
             } else {         //返工
                if (is_null($row5['TC_SRG019'])){
                    $today=date('Y-m-d');
                    $time=date('h:i:s');
                    $erp_sql = oci_parse($erp_conn, "update tc_srg_file set tc_srg019=to_date('$today','YYYY-MM-DD'), tc_srg020='$time', tc_srg021='$gen01' where tc_srg001='$sfb01' and tc_srg004='". $row4['TC_SRG004'] . "' ");
                    oci_execute($erp_sql);                                      
                    $new='Y';
                } else {
                    $errmsg='本工序已報過工!!';    
                    $o->assign('orderstatus','innerHTML',$errmsg);
                } 
             }             
          }        
      }
      if ($new=='Y'){   //有報過工
        $i=$v['orderno'];
        if ($i>10) {
            $ordersvalue='';
            $i=1;
        } else {
            $ordersvalue=$v['ordersvalue'];
            $i++;
        }
        $ordersvalue=$ordersvalue. '<tr><td width="3%">' . '<img src="i/arrow.gif" width="16" height="16">' . '</td>';     
        $ordersvalue=$ordersvalue . '<td width=10%>' . $sfb01               . '</td>' ;   
        $ordersvalue=$ordersvalue . '<td width=10%>' . $ecb06               . '</td>' ; 
        $ordersvalue=$ordersvalue . '<td width=10%>' . $row5['TC_SRG003']   . '</td>' ;  
        $ordersvalue=$ordersvalue . '<td width=10%>' . $row5['TC_SRG018']   . '</td>' ; 
        $ordersvalue=$ordersvalue . '<td width=10%>' . $row5['TC_SRG006']   . '</td>' ; 
        $ordersvalue=$ordersvalue . '<td width=10%>' . $today               . '</td>' ; 
        $ordersvalue=$ordersvalue . '<td width=10%>' . $time                . '</td>' ; 
        $ordersvalue=$ordersvalue . '<td width=10%>' . $row3['GEN02']        . '</td>' ; 
        $ordersvalue=$ordersvalue . '</tr>';         
        $ordershtml= '<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="tabel1">' . $ordersvalue . '</table>'; 
        $o->assign("orderno","value",$i);       
        $o->assign('orderstatus','innerHTML','');   
        $o->assign("ordershtml","innerHTML", $ordershtml); 
        $o->assign("ordersvalue","value",$ordersvalue);   
            
      }
      return $o;   
  }     
      
  $xajax->register(XAJAX_FUNCTION,'checkinOrders');    
  $xajax->processRequest();

  echo '<?xml version="1.0" encoding="UTF-8"?>'; 
  $IsAjax = True;  
                    
  //include("_header.php");
?>

<script language='JavaScript'>  
  checked = false;     
  
  window.onload=setfocusfield;       
  function setfocusfield(){
    document.form1.ecb06.focus();
  }
     
  function checkIn(){          
    xajax_checkinOrders(xajax.getFormValues('form1'));        
    document.form1.sfb01.value='';
    document.form1.sfb01.focus();
    return false;
  }    

</script> 
<script language="javascript" type="text/javascript" src="datepicker/WdatePicker.js"></script> 
<link href="css.css" rel="stylesheet" type="text/css">
<p>請輸入以下資料來報工 </p>
  
<form name="form1" method="POST" action="javascript:void(null)">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr width="100%">
        <td bgcolor="#77C9E9">作業編號: 
          <input name="ecb06" type="text" id="ecb06" size="4" maxlength="4"> &nbsp;&nbsp;  
        </td>
      </tr>
      <tr>
        <td bgcolor="#77C9E9">員工編號:
          <input name="gen01" type="text" id="gen01" size="08" maxlength="08"> &nbsp;&nbsp;    
        </td>  
      </tr>
      <tr>           
        <td bgcolor="#77C9E9">工單編號:
          <input name="sfb01" type="text" id="sfb01" size="15" maxlength="15"> &nbsp;&nbsp;     
        </td> 
      </tr>
      <tr>  
        <td bgcolor="#77C9E9">
          <input name="checkin" type="button" id="checkin" value="報工" onfocus="checkIn()">&nbsp;&nbsp; 
          <span style="color:red" id="orderstatus"></span>              
        </td>
      </tr>
    </table>
  </div>        
  <table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table1" > 
    <tr>                         
      <th width="3%">&nbsp;</th>
      <th width=10%>工單號</th>
      <th width=10%>作業編號</th>  
      <th width=10%>產品編號</th> 
      <th width=10%>返工否</th>  
      <th width=10%>PQC否</th>
      <th width=10%>進站日期</th>
      <th width=10%>進站時間</th>  
      <th width=10%>進站人</th>    
    </tr>    
   </table>
 
   <div id="ordershtml"></div>   
   <input name="ordersvalue" type="hidden">
   <input name="orderno" type="hidden" value='1'>  
   <input name="guid" type="hidden">     
</form>
