<?
  session_start();
  $pagtitle = "IT &raquo; 產生工作日期";
  include("_data.php");

  //auth("erp_createworkday.php");
   
  $bdate = date('Y-m-d');
  $edate = date('Y-m-d');
    
  if ($_GET["action"] == "generate") {
    $bdate=$_GET['bdate'];
    $edate=$_GET['edate'];
    //tc_frd001  日期文字版 index
    //tc_frd002  是否週五日
    //tc_frd003  是否國定假日
    //tc_frd011  日期  search
    //tc_frd013  'A' 用此欄位來分辨儲存的資料用在何處
    //先清空現有的所有假日
    $s1="delete from tc_frd_file 
        where tc_frd013='A'
        and tc_frd011 between to_date('$bdate', 'yy/mm/dd') and to_date('$edate', 'yy/mm/dd')";
    $erp_sql1 = oci_parse($erp_conn1,$s1);
    oci_execute($erp_sql1);

    $tdate = $bdate;
    while ($tdate<=$edate) {
      echo $tdate;
      $weekday = date('w',strtotime($tdate));
      $isFriday='N';
      if($weekday==0 || $weekday==5) {
        $isFriday='Y';
      }
      echo $weekday;
      $s3="insert into tc_frd_file values (
        '$tdate', '$isFriday', 'N',
        to_date('$tdate', 'yy/mm/dd'), NULL, 'A',
        NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 
        NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL
      )";
      $erp_sql3 = oci_parse($erp_conn1,$s3 );
      oci_execute($erp_sql3);
      $tdate = date('Y-m-d', strtotime('+1 days', strtotime($tdate)));
    }

    msg('設定成功');
    forward("erp_createworkday.php");
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
<p>產生工作日期 </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">
          起訖日期:
          <input name="bdate" type="text" id="bdate" size="12" maxlength="12" onfocus="new WdatePicker()" value=<?=$bdate;?>>&nbsp; &nbsp ～～
          <input name="edate" type="text" id="edate" size="12" maxlength="12" onfocus="new WdatePicker()" value=<?=$edate;?>>&nbsp; &nbsp;
          <input type="submit" name="Submit2" value="送出" />
          <input type="hidden" name="action" value="generate">
        </div></td>
      </tr>
    </table>
  </div>
</form>

<? if (is_null($_GET['Submit2'])) die ; ?>
<form action="<?=$PHP_SELF;?>" method="post" name="form1">
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
  <tr>
    <th width="16">&nbsp;</th>
    <th>Order No.</th>
    <th>RX #</th>
    <th>Order Date</th>
    <th>Due Date</th>
    <th>New Order Date</th>
    <th>New Due Date</th>
    <th width="16"><input type="checkbox" value="checkbox" name='checkall' onclick='checkedAll();'></th>
  </tr>
  <?
      if ($rxno=='') {
          $rxnofilter='';
      } else  {
          $rxnofilter=" and ta_oea006 like '$rxno%' ";
      }

      if ($occ01=='') {
          $occfilter='';
      } else  {
          $occfilter=" and oea04='$occ01' ";
      }

      if ($orderdate==''){
          $orderdatefilter='';
      } else {
          $orderdate1=substr($orderdate,2,2).substr($orderdate,5,2).substr($orderdate,8,2);
          $orderdatefilter=" and oea02=to_date('$orderdate1','yy/mm/dd') " ;
      }

      $soea="select oea01, to_char(ta_oea005,'yyyy-mm-dd') ta_oea005, to_char(oea02,'yyyy-mm-dd') oea02, ta_oea006 from vd210.oea_file " .
            "where 1=1 " . $occfilter . $rxnofilter . $orderdatefilter . " order by oea01 ";
      $erp_sqloea = oci_parse($erp_conn2,$soea );
      oci_execute($erp_sqloea);
      while ($rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC)) {
      ?>
	    <tr bgcolor="#FFFFFF">
		      <td><img src="i/arrow.gif" width="16" height="16">
              <input type="hidden" name="oeaarray[]" value="<?=$rowoea['OEA01'];?>">
          </td>
			    <td><?=$rowoea["OEA01"];?></td>
          <td><?=$rowoea["TA_OEA006"];?></td>
          <td><?=$rowoea["OEA02"];?></td>
          <td><?=$rowoea["TA_OEA005"];?></td>
          <td><?=$neworderdate;?></td>
          <td><?=$newduedate;?></td>
          <td width=16><input name="risok<?=$rowoea['OEA01'];?>" type="checkbox" id="risok" value="Y" </td>
          <input name="oea02<?=$rowoea['OEA01'];?>" type="hidden" value="<?=$rowoea["OEA02"];?>">
          <input name="taoea005<?=$rowoea['OEA01'];?>" type="hidden" value="<?=$rowoea["TA_OEA005"];?>">
      </tr>
      <?
      }
  ?>
  <tr>
    <td>&nbsp;</td>
    <td>
        <input type="hidden" name="neworderdate"  value="<?=$neworderdate;?>">
        <input type="hidden" name="newduedate"    value="<?=$newduedate;?>">
        <input type="hidden" name="action"        value="save">
        <input type="submit" name="Submit"        value="更新">
    </td>
  </tr>
</table>
</form>
