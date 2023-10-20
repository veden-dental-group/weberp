<?
  session_start();
  $pagtitle = "IT &raquo; 新增sfb_file";
  include("_data.php");
  include("_erp.php");
  //auth("erp_add_sfb_file.php");


  if ($_POST["action"] == "save") {
      foreach ($_POST["ssfb01"] as $ssfb01){
          if ($_POST["risok" . $ssfb01] == "Y") {

              $fsfb01=$_POST['fsfb01'.$ssfb01];

              //copy $tsfb01 的值, 放到 $ssfb01裡
              // 但要保留訂單日期 顆數 出貨日期
              if ($fsfb01!='') { //有可以拷貝的工單號才做
                  //加備料檔
                  // 先看原來有無需要備料的
                  /*
                  $s1="select sfa01 from sfa_file where sfa01='$fsfb01'";
                  $erp_sql1 = oci_parse($erp_conn1,$s1 );
                  oci_execute($erp_sql1);
                  $row1 = oci_fetch_array($erp_sql1, OCI_ASSOC);
                  if ($row1['SFA01']==$fsfb01) {
                      $s2="insert into vd110.sfa_file " .
                            "( select '$ssfb01', sfa02, sfa03, sfa04, sfa05, sfa06, sfa061, sfa062, sfa063, sfa064, sfa065, sfa066, sfa07, sfa08, " .
                            "sfa09, sfa10, sfa11, sfa12, sfa13, sfa14, sfa15, sfa16, sfa161, sfa25, sfa26, sfa27, sfa28, sfa29, sfa30, sfa31, " .
                            "sfa91, sfa92, sfa93, sfa94, sfa95, sfa96, sfa97, sfa98, sfa99, sfa100, sfaacti, sfa32, sfaud01, sfaud02, sfaud03, " .
                            "sfaud04, sfaud05, '20150507_frank_insert', sfaud07, sfaud08, sfaud09, sfaud10, sfaud11, sfaud12, sfaud13, sfaud14, sfaud15, sfa36, " .
                            "sfaplant, sfalegal, sfa012, sfa013 " .
                            "from sfa_file where sfa01='$fsfb01')";
                      $erp_sql2 = oci_parse($erp_conn1,$s2 );
                     // oci_execute($erp_sql2);
                  }
                  */

                  //加工單
                  //先取出訂單編號, 項次, 數量, 到貨日期, 出貨日期
                  $s1="select to_char(tc_oga002, 'yyyy-mm-dd') tc_oga002,  tc_ogb003, tc_ogb004, tc_ogb006, tc_ogb011, to_char(oga02, 'yyyy-mm-dd') oga02 " .
                      "from tc_ogb_file, tc_oga_file, oga_file " .
                      "where tc_oga001=tc_ogb001 and tc_ogb003=oga16 and " .
                      "tc_ogb002='$ssfb01' ";
                  $erp_sql1 = oci_parse($erp_conn1,$s1 );
                  oci_execute($erp_sql1);
                  $row1     = oci_fetch_array($erp_sql1, OCI_ASSOC);
                  $no       = $row1['TC_OGB003'];
                  $sn       = $row1['TC_OGB004'];
                  $qty      = $row1['TC_OGB006'];
                  $rx       = $row1['TC_OGB011'];
                  $outdate  = $row1['OGA02'];
                  $indate   = $row1['TC_OGA002'];

                  $s2="insert into vd110.sfb_file " .
                        "( select '$ssfb01', sfb02, sfb03, sfb04, sfb05, sfb06, sfb07, to_date('$indate','yy/mm/dd'), $qty, $qty, $qty, " .
                        "sfb10, sfb11, sfb111, sfb12, sfb121, sfb122, to_date('$indate','yy/mm/dd'), sfb14, to_date('$outdate','yy/mm/dd'), " .
                        "sfb16, sfb17, sfb18, sfb19, sfb20, sfb21, '$no', '$sn', sfb222, sfb23, sfb24, to_date('$indate','yy/mm/dd'), " .
                        "to_date('$indate','yy/mm/dd'), sfb26, sfb27, sfb271, sfb28, sfb29, sfb30, sfb31, sfb32, sfb33, sfb34, sfb35, " .
                        "sfb36, sfb37, to_date('$outdate','yy/mm/dd'), sfb39, sfb40, sfb41, sfb42, to_date('$indate','yy/mm/dd'), sfb82, " .
                        "sfb85, sfb86, sfb87, sfb88, sfb91, sfb92, sfb93, sfb94, sfb95, sfb96, sfb97, sfb98, sfb99, sfb100, sfb101, " .
                        "sfbacti, sfbuser, sfbgrup, sfbmodu, to_date('$indate','yy/mm/dd'), sfb1001, sfb1002, sfb1003, sfb102, sfb50, sfb51, sfb103, " .
                        "sfbud01, '$rx', sfbud03, '20150507_frank_insert', sfbud05, sfbud06, sfbud07, sfbud08, sfbud09, sfbud10, sfbud11, " .
                        "sfbud12, sfbud13, sfbud14, sfbud15, sfb43, sfb44, sfbmksg, " .
                        "sfbplant, sfblegal, sfboriu, sfborig, sfb104 " .
                        "from sfb_file where sfb01='$fsfb01')";
                  $erp_sql2 = oci_parse($erp_conn1,$s2 );
                  oci_execute($erp_sql2);

              }
          }
      }          
      msg ('更新完畢');
      forward("erp_add_sfb_file.php");
  }

  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<form action="<?=$PHP_SELF;?>" method="post" name="form1">
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
  <tr>
    <th width="16">&nbsp;</th>
    <th>舊工單號碼</th>
    <th>舊客戶編號</th>
    <th>舊訂單號碼</th>
    <th>舊項次</th>
    <th>舊訂單日期</th>
    <th>舊品代</th>
    <th>舊數量</th>
    <th>舊出貨日期</th>
    <th width="16">&nbsp;</th>

    <th>新工單號碼</th>
    <th>新訂單號碼</th>
    <th>新項次</th>
    <th>新訂單日期</th>
    <th>新數量</th>
    <th width="16">&nbsp;</th>
  </tr>
  <?
      $x     = $_GET['x'];
      $xx    = 0;
      $sfb01 = $_GET['sfb01'];
      // 先取得該工單的相關資料
      $s1 ="select tc_ogb002, tc_ogb003, tc_ogb004, tc_ogb011, tc_ogb005, tc_ogb006, tc_oga004, to_char(tc_oga002,'yyyy-mm-dd') tc_oga002, to_char(oea02,'yyyy-mm-dd') oea02
            from tc_ogb_file, tc_oga_file, oea_file
            where tc_oga001=tc_ogb001 and tc_ogb003=oea01 and
            tc_ogb002 in
           -- ( 'A311-1409020159', 'A311-1409020160' )
            ( select sfv11
              from sfv_file, sfu_file
              where sfv01=sfu01 and sfu02 between to_date('140901','yy/mm/dd') and to_date('140930','yy/mm/dd')
              and not exists ( select 1 from sfb_file where sfb01=sfv11)
            )
            order by tc_ogb003, tc_ogb004 ";

      $erp_sql1 = oci_parse($erp_conn1,$s1 );
      oci_execute($erp_sql1);
      while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
          $xx++;
          $ssfb01 = $row1['TC_OGB002'];
          $sima01 = $row1['TC_OGB005'];
          $soea01 = $row1['TC_OGB003'];
          $soea011= $row1['TC_OGB004'];
          $soea02 = $row1['OEA02'];
          $soga02 = $row1['TC_OGA002'];
          $socc01 = $row1['TC_OGA004'];
          $sqty   = $row1['TC_OGB006'];

          $s2="select sfb01, sfb22, sfb221, to_char(sfb81,'yyyy-mm-dd') sfb81, sfb08 " .
              " from sfb_file " .
              " where sfb05='$sima01' ".
             // " and sfb08>=$sqty " .
              //" and sfb221=$soea011 " .
              //" and sfb81 >= to_date('2014/09/01','yy/mm/dd') and sfb81 <= to_date('2014/09/30','yy/mm/dd') ".
              " and sfb04=8 " ;
          $erp_sql2 = oci_parse($erp_conn1,$s2 );
          oci_execute($erp_sql2);
          $x=1;
          $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);
      ?>
    	    <tr bgcolor="#FFFFFF">
    		      <td><?=$xx;?>
                  <input type="hidden" name="fsfb01<?=$ssfb01;?>" value="<?=$row2['SFB01'];?>">
                  <input type="hidden" name="ssfb01[]" value="<?=$ssfb01;?>">
              </td>
              <td><?=$ssfb01;?></td>
              <td><?=$socc01;?></td>
              <td><?=$soea01;?></td>
              <td><?=$soea011;?></td>
              <td><?=$soea02;?></td>
              <td><?=$sima01;?></td>
              <td><?=$sqty;?></td>
              <td><?=$soga02;?></td>
              <td><img src="i/arrow.gif" width="16" height="16">
    			    <td><?=$row2["SFB01"];?></td>
              <td><?=$row2["SFB22"];?></td>
              <td><?=$row2["SFB221"];?></td>
              <td><?=$row2["SFB81"];?></td>
              <td><?=$row2["SFB08"];?></td>
              <td width=16><input name="risok<?=$ssfb01;?>" type="checkbox" id="risok" value="Y"> </td>
          </tr>

        <?
      }
  ?>
  <tr>
    <td>&nbsp;</td>
    <td>
        <input type="hidden" name="action" value="save">
        <input type="submit" name="Submit" value="更新">
    </td>
  </tr>
</table>
</form>
