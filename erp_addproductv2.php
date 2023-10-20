<?
  
  session_start();
  $pagtitle = "業務部 &raquo; 增加訂單的產品"; 
  include("_data.php");
  include("_erp.php");
  //auth("erp_addproduct.php");      

  if ($_POST["action"] == "save") {
      $oea04=$_POST['oea04'];  //客戶編號  
      $oeb01=$_POST['oea01'];  //訂單號
      $oeb03=$_POST['oeb03'];  //項次
      $oeb04=$_POST['oeb04'];  //品代
      $oeb15=$_POST['oeb15'];  //order date  
      $duedate=$_POST['duedate'];  //new duedate date
      $oebud03=$_POST['oebud03'];    //齒位
      $oebud03a=explode('|', $oebud03);
      $oeb12=count($oebud03a);          //數量

      //取出該產品的中文名稱, 單位(顆/床)
      $sima="select ima02, ima25 from ima_file where ima01='$oeb04'";
      $erp_sqlima = oci_parse($erp_conn2,$sima );
      oci_execute($erp_sqlima); 
      $rowima = oci_fetch_array($erp_sqlima, OCI_ASSOC);
      $errormsg='產品新增完畢, 請至VD110中將訂單轉成工單!!';
      if (is_null($rowima['IMA02'])) {
          $errormsg='查無此產品編號: ' . $oeb04;
      } else if (is_null($duedate)|| $duedate=='') {
          $errormsg='請輸入新的出貨日期';
      } else {
          $ima02=$rowima['IMA02'];  //品名
          $ima25=$rowima['IMA25'];  //單位
                                                       
          //取出該客戶, 該產品的單價, 幣別
          $socc="select xmf07, occ42 from xmf_file, occ_file where xmf01=occ44 and occ01='$oea04' and xmf03='$oeb04' and xmf05<=to_date('$oeb15','yy/mm/dd') order by xmf03, xmf05 desc";  
          $erp_sqlocc = oci_parse($erp_conn2,$socc );
          oci_execute($erp_sqlocc); 
          $rowocc = oci_fetch_array($erp_sqlocc, OCI_ASSOC);
          if (is_null($rowocc['XMF07'])) {
              $errormsg= 'VD210 境外查無單價, 請連絡帳單組輸入後, 再新增. 資料, 客戶:' . $oea04 . ' 產品編號: ' . $oeb04 ;  
          } else {
              $xmf07=floatval($rowocc['XMF07']);    //單價
              $occ42=$rowocc['OCC42'];   
          
              //採購單 要計算本幣的單價, 需要匯率 幣別已取出放在$occ42
              $azj02=substr($oeb15,0,4).substr($oeb15,5,2);
              $sazj="select azj041 from azj_file where azj01='$occ42' and azj02='$azj02' ";  
              $erp_sqlazj = oci_parse($erp_conn2,$sazj );
              oci_execute($erp_sqlazj); 
              $rowazj = oci_fetch_array($erp_sqlazj, OCI_ASSOC);
              if (is_null($rowazj['AZJ041'])) {
                  $errormsg='VD210 境外無幣別匯率, 請連絡財務部輸入後, 再新增. 日期: ' . $oeb05 . ' 匯率:' . $occ42 ;      
              } else {
                  $azj041=floatval($rowazj['AZJ041']);       //匯率                                        
              
                  //取出該客戶, 該產品的製處代號
                  $stcocc="select tc_occ003 from tc_occ_file, ima_file where tc_occ001='$oea04' and tc_occ002=ima06 and ima01='$oeb04'";  
                  $erp_sqltcocc = oci_parse($erp_conn2,$stcocc );
                  oci_execute($erp_sqltcocc); 
                  $rowtcocc = oci_fetch_array($erp_sqltcocc, OCI_ASSOC);
                  if (is_null($rowtcocc['TC_OCC003'])) {
                      $errormsg='查無該客戶:' . $oea04 . ' 產品編號: ' . $oeb04 . ' 的製處' ;  
                  } else {
                      $tc_occ003=$rowtcocc['TC_OCC003'];  
                  
                      $oeb14= $oeb12*$xmf07;             //總價
                      //寫入vd210的oeb 訂單單身
                      $soeb="insert into oeb_file values " .
                            "('$oeb01', '$oeb03', '$oeb04', '$ima25', 1, '$ima02',NULL,NULL,'9999',NULL,NULL,NULL,$oeb12, $xmf07, $oeb14, $oeb14, to_date('$oeb15','yy/mm/dd'), to_date('$oeb15','yy/mm/dd'), " .
                            "0,NULL,'N',NULL,NULL,NULL,0,0,0,0,'N',NULL,NULL,NULL,0,NULL,NULL,NULL,0,'N',NULL,NULL,NULL,'$ima25',1,NULL,NULL,NULL,NULL,'$ima25', $oeb12,NULL,NULL,1,NULL,NULL,100,NULL,NULL,NULL, " .
                            "NULL,NULL,'N',0,NULL,NULL,NULL,0,0,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'$oebud03','$tc_occ003',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL, " .
                            "NULL,1,NULL,NULL,0,2,'VD210','VD200')"; 
                      $erp_sqloeb = oci_parse($erp_conn2,$soeb );
                      oci_execute($erp_sqloeb);
                      
                      //寫入vd210 的pmn 採購單單身
                      $pmn30=$xmf07*$azj041;
                      $spmn="insert into pmn_file values " .
                            "('$oeb01','TRI', $oeb03,NULL,'$oeb04', '$ima02',NULL,NULL,'$ima25','$ima25',1,NULL,'N',1,NULL,NULL, 0, 'Y','Y','2',NULL,$oeb12, NULL, '$oeb01', $oeb03, $pmn30, $xmf07, $xmf07, NULL, " .
                            "to_date('$oeb15','yy/mm/dd'),to_date('$oeb15','yy/mm/dd'),to_date('$oeb15','yy/mm/dd'),to_date('$oeb15','yy/mm/dd'), NULL, 'Y', NULL, NULL, 0, 0, 0, $pmn30, NULL, NULL, 0, 0, NULL, 0, NULL, 0, NULL, 0, 0, NULL, " .
                            "NULL, '$oeb04', 1, 'N', 'N', '1',NULL, NULL, NULL, NULL, NULL, NULL, '$ima25', 1, NULL, NULL, NULL, NULL, '$ima25', $oeb12, $oeb14, $oeb14, NULL, NULL, $xmf07, NULL, NULL, NULL, NULL, " .
                            "NULL, NULL, NULL, NULL, '$oebud03', '$tc_occ003', NULL, NULL,$oeb12, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 'N', NULL, ' ', NULL, NULL, NULL, NULL, 'VD210', 'VD200', NULL)"; 
                      $erp_sqlpmn = oci_parse($erp_conn2,$spmn );
                      oci_execute($erp_sqlpmn);
                      
                      //寫入vd210 的tc_ext 應計費用別單身       
                      $sext="insert into tc_ext_file values " .
                            "('$oeb01',$oeb03,'$oeb04',$oeb12,'$ima25',$xmf07,$oeb14,'4',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL)"; 
                      $erp_sqlext = oci_parse($erp_conn2,$sext );
                      oci_execute($erp_sqlext);
                      
                      //寫入vd110的oeb 訂單單身
                      $soeb="insert into oeb_file values " .
                            "('$oeb01', '$oeb03', '$oeb04', '$ima25', 1, '$ima02',NULL,NULL,'9999',NULL,NULL,NULL,$oeb12, $xmf07, $oeb14, $oeb14, to_date('$oeb15','yy/mm/dd'), to_date('$oeb15','yy/mm/dd'), " .
                            "$xmf07,NULL,'N',NULL,NULL,NULL,0,0,0,0,'N',NULL,NULL,NULL,0,NULL,NULL,NULL,0,'N',NULL,NULL,NULL,'$ima25',1,NULL,NULL,NULL,NULL,'$ima25', $oeb12,NULL,NULL,1,NULL,NULL,100,NULL,NULL,NULL, " .
                            "NULL,NULL,'N',NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'$oebud03','$tc_occ003',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL, " .
                            "NULL,' ',NULL,NULL,0,1,'VD110','VD100')"; 
                      $erp_sqloeb = oci_parse($erp_conn1,$soeb );
                      oci_execute($erp_sqloeb);
                      
                      //寫入vd110 的tc_ext 應計費用別單身       
                      $sext="insert into tc_ext_file values " .
                            "('$oeb01',$oeb03,'$oeb04',$oeb12,'$ima25',$xmf07,$oeb14,'4',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL)"; 
                      $erp_sqlext = oci_parse($erp_conn1,$sext );
                      oci_execute($erp_sqlext);

                      // 更新訂單出貨日期
                      $sduedate = "update oea_file set ta_oea005 = to_date('$duedate','yy/mm/dd') where oea01='$oeb01'";
                      $erp_sqlduedate1 = oci_parse($erp_conn1,$sduedate );
                      oci_execute($erp_sqlduedate1);
                      $erp_sqlduedate2 = oci_parse($erp_conn2,$sduedate );
                      oci_execute($erp_sqlduedate2);
                  }
              }   
          }       
      }  
       
      msg ($errormsg);
      forward("erp_addproductv2.php");
  }
  if (is_null($_GET['oea01'])){
      $oea01='B531-';
  } else {
      $oea01=$_GET['oea01'];
  }
  include("_header.php");
?>
<script language='JavaScript'>
checked = false;
function checkedAll () {
  if (checked == false) {
    checked = true
  }else{
    checked = false
  }
  
  for (var i = 0; i < document.form1.elements.length; i++) {
    var e = document.form1.elements[i];
        if (e.type == 'checkbox' && e.disabled==false) {
            e.checked = checked
        }                                                          
  }  
}
</script>
<link href="css.css" rel="stylesheet" type="text/css">
<p>增加產品到訂單中 *** 限未掃描秤重的工單 *** </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">
            訂單號碼:   
            <input name="oea01" type="text" id="oea01" size="15" maxlength="15" value="<?=$oea01;?>"> &nbsp;  &nbsp; 
            <input type="submit" name="Submit2" value="查詢" />
            </div></td>        
        </tr>
    </table>
  </div>
</form>
<?
if (is_null($_GET['Submit2'])) die ; 
//VD210檢查有無訂單
$soea="select to_char(oea02,'yyyy/mm/dd') oea02, to_char(ta_oea005,'yyyy/mm/dd') ta_oea005, ta_oea006, oea04, occ02, ta_oea011, oea99 from oea_file, occ_file " .
      "where oea01='$oea01' and oea03=occ01 "; 
$erp_sqloea = oci_parse($erp_conn2,$soea );
oci_execute($erp_sqloea); 
$rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC);
if (is_null($rowoea['OEA02'])) {
    msg('查無此訂單號: ' . $oea01 );
    forward('erp_addproduct.php');
}

//檢查VD110有無出貨
$soga="select tc_oga001, to_char(tc_oga002,'yyyy/mm/dd') tc_oga002 from tc_oga_file, tc_ogb_file where tc_oga001=tc_ogb001 and tc_ogb003='$oea01' "; 
$erp_sqloga = oci_parse($erp_conn1,$soga );
oci_execute($erp_sqloga); 
$rowoga = oci_fetch_array($erp_sqloga, OCI_ASSOC);
if (!is_null($rowoga['TC_OGA001'])) {
    msg('訂單號 '. $oea01 . ' 已於 ' . $rowoga['TC_OGA002'] . ' 秤重, 秤重單號: ' . $rowoga['TC_OGA001'] );
    forward('erp_addproduct.php'); 
}

if ($rowoea['TYPE']=='1'){
    $type='Crown and Bridge Work';
} else if ($rowoea['TYPE']=='2'){  
    $type='Removable Prosthesis';  
} else {
    $type='Combination Work'; 
}                            

//取出該訂單的產品 寫入陣列中
$oeb03  =array();
$oeb04  =array();
$oeb06  =array();
$oebud03=array();
$oeb12  =array();
$oeb05  =array();
$i=1;
$soeb="select oeb03, oeb04, oeb06, oebud03, oeb12, oeb05 from oeb_file where oeb01='$oea01'  "; 
$erp_sqloeb = oci_parse($erp_conn2,$soeb);                                                
oci_execute($erp_sqloeb); 
while ($rowoeb = oci_fetch_array($erp_sqloeb, OCI_ASSOC)) { 
    $oeb03[$i]  =$rowoeb['OEB03'];
    $oeb04[$i]  =$rowoeb['OEB04'];
    $oeb06[$i]  =$rowoeb['OEB06'];
    $oebud03[$i]=$rowoeb['OEBUD03'];
    $oeb12[$i]  =$rowoeb['OEB12'];
    $oeb05[$i]  =$rowoeb['OEB05'];
    $i++;
}

//取出該訂單的最高項次
$smoeb="select max(oeb03) oeb03 from oeb_file where oeb01='$oea01'  "; 
$erp_sqlmoeb = oci_parse($erp_conn2,$smoeb);                                                
oci_execute($erp_sqlmoeb); 
$rowmoeb = oci_fetch_array($erp_sqlmoeb, OCI_ASSOC);
$moeb03=intval($rowmoeb['OEB03'])+1;
?>
<form action="<?=$PHP_SELF;?>" method="post" name="form1">
訂單內容:
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
  <tr>                     
    <th>Order Date</th>
    <th>Due Date</th> 
    <th>RX # </th>
    <th>送貨客戶</th>  
    <th>Type</th>
  </tr>
  <tr>  
    <td><?=$rowoea['OEA02'];?></td>
    <td><?=$rowoea['TA_OEA005'];?></td>    
    <td><?=$rowoea['TA_OEA006'];?></td>  
    <td><?=$rowoea['OEA04'].' -- '.$rowoea['OCC02'];?></td>    
    <td><?=$type;?></td>    
  </tr>
</table>
<br>
產品清單:
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4"> 
  <tr>                        
    <th>項次</th>
    <th>品代</th> 
    <th>品名</th>
    <th>齒位</th>  
    <th>數量</th>
  </tr>
  <?
  for ($j=1;$j<$i;$j++) {
  ?>
  <tr>  
    <td><?=$oeb03[$j];?></td>   
    <td><?=$oeb04[$j];?></td>   
    <td><?=$oeb06[$j];?></td>   
    <td><?=$oebud03[$j];?></td>   
    <td><?=$oeb12[$j] . ' ' . $oeb15[$j];?></td>   
  </tr>
  <? } ?>
</table>
<br>
新增產品:
<table>   
  <tr>
      <td bgcolor="#FF66FF" class="witbold">項次: </td>
      <td><input name="oeb03" type="text" id="oeb03" readonly="true" value="<?=$moeb03;?>"></td>
    </tr>
  <tr>
    <td bgcolor="#FF66FF" class="witbold">產品編號: *</td>
    <td><input name="oeb04" type="text" id="oeb04" size="10"> </td>
  </tr>
  <tr>
    <td bgcolor="#FF66FF" class="witbold">齒位/床位(收費用):</td>
    <td><input name="oebud03" type="text" id="oebud03" size="50">  (歐洲齒位, 個齒位間請用 | 隔開)</td>
  </tr>
    <tr>
        <td bgcolor="#FF66FF" class="witbold">新的Due Date:</td>
        <td>
            <input name="duedate" type="text" id="duedate" size="12" maxlength="12" onfocus="new WdatePicker()">
        </td>
    </tr>
    <tr>
    <td>&nbsp;</td>
    <td><input type="hidden" name="oea01"  value="<?=$_GET['oea01'];?>">     
        <input type="hidden" name="oeb15"  value="<?=$rowoea['OEA02'];?>">      
        <input type="hidden" name="oea04"  value="<?=$rowoea['OEA04'];?>">       
        <input type="hidden" name="action" value="save">
        <input type="submit" name="Submit" value="新增">        
    </td>
  </tr>
</table>  
</form>
