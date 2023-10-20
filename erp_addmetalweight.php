<?
  session_start();
  $pagtitle = "業務部 &raquo; 輸入研磨後重量"; 
  include("_data.php");
  include("_erp.php");
  auth("erp_addmetalweight.php");
  
  if (is_null($_GET['bdate'])) {
    $bdate=date('Y-m-d');
  } else {
    $bdate=$_GET['bdate'];
  }
  $occ01=$_GET['occ01'];
  
  if (is_null($_GET['clienttype'])) {
    $clienttype='1';
  } else {
    $clienttype=$_GET['clienttype'];
  }
                        
  if ($clienttype=='1') {  
    $occfilter = " and tc_oga004='$occ01' ";     
  } else if ($clienttype=='3') {      
    $occfilter = " and tc_oga004 like 'T183%' ";
  } else {  
    $occfilter = " and oea04 like 'H101%' "; 
  }    
     
  
  //客戶金屬重量加損公司如下:
  //一種金屬只會用在鋼牙或磁牙 不會兩者共用同一種金屬
  // 1:鋼牙用金屬  2/B/C/NF:為磁牙用
  // CNC B金屬: 每顆1.1g 不管用多少
  // CNC C金屬: 每顆0.9g 不管用多少
  // UC 磁牙加收研磨後10%
  // PLS NF >= 1*每顆
  // 其餘任何客戶 鋼牙(1)任何貴金屬加收 15%  磁牙(非1)則收鑄造後重量
  // 研磨後 = 各次研磨後重 - 退料 - 退料損耗
  // 鑄造後 = 各次鑄造後重 - 退料 - 退料損耗    
  
  
  //另外在tc_dex013 放置超重的重量
  //1.  瓷 +咬金 /瓷牙（种植体）/开面冠     2.5g/颗
  //2.  钢牙（种植体）/瓷牙+post 连体         3.5g/颗
  //3.  钢牙 onlay     1.5g/颗
  //4.  钢牙inlay      1.0g/颗
  //5.   普通烤瓷牙    2.0g/颗
  //6.   钢牙          3.0g/颗

  
     
  if ($_POST["action"] == "save") {     
      $isdelay=$_POST['isdelay'];       
      $msg=''; 
      //取出客戶是否為CNC, UC    
      $socc= "select occud10 from occ_file where occ01='$occ01'"; 
      $erp_sqlocc = oci_parse($erp_conn1,$socc );
      oci_execute($erp_sqlocc); 
      $rowocc=oci_fetch_array($erp_sqlocc, OCI_ASSOC);
      $occtype=$rowocc['OCCUD10']; //1:其他 2:CNC 3:UC         
      foreach ($_POST["sfearray"] as $sfe){ 
          $sfe01=$_POST['sfe01'.$sfe];      
          $sfe02=$_POST['sfe02'.$sfe];  
          $sfe07=$_POST['sfe07'.$sfe]; 
          $sfe28=$_POST['sfe28'.$sfe];  
          $sfb05=$_POST['sfb05'.$sfe];    
          $tasfe002=floatval($_POST['tasfe002'.$sfe]);
          $outqty=floatval($_POST['outqty'.$sfe]);
          $imaud10=$_POST['imaud10'.$sfe]; 
          //先將要修改的資料記錄下來                                    
          $queryap   = "insert into erp_addmetalweight_updaterecord ( bdate, occ01, sfe02, sfe28, weight, username, ip ) values (
                      '" . safetext($bdate)                   . "',     
                      '" . safetext($occ01)                   . "',  
                      '" . safetext($sfe02)                   . "',      
                      '" . safetext($sfe28)                   . "',     
                      '" . safetext($tasfe002)                . "',       
                      '" . safetext($_SESSION['account'])     . "',       
                      '" . safetext($_SERVER['REMOTE_ADDR'])  . "')";
          $resultap= mysql_query($queryap) or die ('31 erp_addmetalweight_updaterecord added error. ' .mysql_error());
          //如果研磨後比鑄造後還重 就以鑄造後
          if($tasfe002>$outqty) {
            $msg.="$sfe02 研磨後>鑄造後 取鑄造後重量!!";
            $tasfe002=$outqty;
          } 
          
          //客戶自備金 以鑄造後重量為重量 其中 EZ:1K01000002  全球:1K02000004
          if ($sfe07=='1K01000002' or $sfe07=='1K02000004') {
              $tasfe002=$outqty;                               
          }        
          
          $ssfe1= "update sfe_file set ta_sfe002=$tasfe002 where sfe02='$sfe02' and sfe28='$sfe28' and sfe07='$sfe07'";
          $erp_sqlsfe1=oci_parse($erp_conn1,$ssfe1);
          $rs1=oci_execute($erp_sqlsfe1);
          if ($rs1) {
              $msg.="更新成功!!";  
          } else {
              $msg.="VD110 sfe_file發料檔 $sfe02 $sfe28 更新失敗!!";  
          }  
                                                                                                                                           
          //合計同一張工單的所有研磨後重量  退料量  損耗量      
          //研磨後重量 ta_sfe002   -   //退料要再扣掉sfe16           -   //退料要扣掉損耗   ta_sfe003
          $ssfe2="select sfe01, sfe07, 
                  sum(decode(sfe06,'4',0,sfe16) - decode(sfe06,'4',sfe16,0) - decode(sfe06,'4',decode(ta_sfe003,null,0,ta_sfe003),0)) sfe16,
                  sum(decode(ta_sfe002,null,0,ta_sfe002) - decode(sfe06,'4',decode(sfe16, null,0,sfe16),0) - decode(sfe06,'4',decode(ta_sfe003,null,0,ta_sfe003),0)) ta_sfe 
                  from sfe_file where sfe01='$sfe01' and sfe07='$sfe07' group by sfe01, sfe07";
          $erp_sqlsfe2=oci_parse($erp_conn1,$ssfe2);
          $rs2=oci_execute($erp_sqlsfe2);
          while ($rowsfe2=oci_fetch_array($erp_sqlsfe2, OCI_ASSOC)) {  //同一張工單 可能有兩種金屬以上 
              $tasfe=$rowsfe2['TA_SFE']; 
              $sfe16=$rowsfe2['SFE16']; 
              //$sfe07=$rowsfe2['SFE07'];
          
              $ssfb="select sfb05, sfb08, sfb22, sfb221, oea04 from sfb_file, oea_file where sfb01='$sfe01' and sfb22=oea01 ";
              $erp_sqlsfb=oci_parse($erp_conn1,$ssfb);
              $rs3=oci_execute($erp_sqlsfb);
              $rowsfb=oci_fetch_array($erp_sqlsfb, OCI_ASSOC); 
              $sfb08=$rowsfb['SFB08'];   //顆數
              $sfb22=$rowsfb['SFB22']; 
              $sfb221=$rowsfb['SFB221']; 
              
              //把原始資料寫起來 供檢杳用
              $tcdex020='客:'. $occtype . ' 金:' . $imaud10 . ' 重:' . $tasfe . ' 顆:' . $sfb08;   
              
              if  ($occtype=='1') {  //PLS 
                  if ($imaud10=='5') {     // 5: PLS CNC 每顆至少1g 
                      //$tasfe *= 1.15;
                      //$tasfe=max($tasfe,$sfb08);  
                      $tasfe=$sfb08;
                  } else {                        //其餘收15%
                      $tasfe *= 1.15;    
                  }                             
              } else if ($occtype=='2') {          //CNC
                  if ($imaud10=='3') {            //3:B 要找到顆數 每顆以0.9收                        
                      $tasfe= $sfb08 * 1.1;   
                  } else if ($imaud10=='4') {     // 4:C 要找到顆數 每顆以1.1收     
                      $tasfe= $sfb08 * 0.9;
                  } else if ($imaud10=='5') {     // 5: NF CNC 每顆至少1g 
                      //$tasfe *= 1.15;  
                      //$tasfe=max($tasfe,$sfb08);  
                      $tasfe=$sfb08; 
                  } else {                        //其餘收15%
                      $tasfe *= 1.15;    
                  }
              } else if ($occtype=='3') {   //UC 磁牙收10% 因此先判斷鋼牙 剩下為磁牙 
                  if ($imaud10=='1') {
                      $tasfe *= 1.15;   
                  } else {
                      $tasfe *= 1.10;     
                  }    
              } else {                      //其餘都加收15%
                 //如果客戶是U118001 (EZ) 且品代為16122 則每顆一律減去0.4g
                  if ($rowsfb['OEA04']=='U118001') {
                      if ($tasfe > 0) {
                          if ($rowsfb['SFB05']=='16122') {   
                              $tasfe -= ($sfb08*0.4);  
                          } else {
                            //有可能U118001 同一張訂單做連體 則發料會一起發 要在另一張工單裡扣除 
                            //若同一張訂單有做16122 但沒有發料的 也要扣除
                              $sm="select count(*) is16122 from sfb_file, sfa_file where sfb22='$sfb22' and sfb05='16122' and sfb01=sfa01 and sfa05 = 0 ";
                              $erp_sqlsm=oci_parse($erp_conn1,$sm);
                              $rsm=oci_execute($erp_sqlsm);  
                              if ($rsm['IS16122']>0) {
                                  $tasfe -= ($sfb08*0.4);   
                              }
                          }
                      }
                  }           
                                 
                  $tasfe *= 1.15;               
              }
                                                      
              //如果研磨後重量為0 則以鑄造後重量當收費重量??
              if ($tasfe<=0) $tasfe=$sfe16;
              
              //客戶自備金 以鑄造後重量為重量 其中 EZ:1K01000002  全球:1K02000004
              if ($sfe07=='1K01000002' or $sfe07=='1K02000004') {
                $tasfe=$sfe16;                               
              }       
              
              
                 
              //四捨五入
              $tasfe=round($tasfe,2);
              
              //取出本產品的標準重量是多少 只算TA的 以加速程式運行
              //另外在tc_dex013 放置超重的重量
              //1.  瓷 +咬金 /瓷牙（种植体）/开面冠     2.5g/颗
              //2.  钢牙（种植体）/瓷牙+post 连体         3.5g/颗
              //3.  钢牙 onlay     1.5g/颗
              //4.  钢牙inlay      1.0g/颗
              //5.   普通烤瓷牙    2.0g/颗
              //6.   钢牙          3.0g/颗
              $overweight=0;
              if (substr($occ01,0,4)=='U121') {     
                  $sima="select imaud11 from ima_file where ima01='$sfb05' ";
                  $erp_sqlima=oci_parse($erp_conn1,$sima);
                  $rsima=oci_execute($erp_sqlima);
                  $rowima=oci_fetch_array($erp_sqlima, OCI_ASSOC);
                  $imaud11=$rowima['IMAUD11'];
                  switch ($imaud11) {
                      case 1:
                          $overweight = $sfb08* 2.5;
                          break;                          
                      case 2:
                          $overweight = $sfb08* 3.5;  
                          break;
                      case 3:
                          $overweight = $sfb08* 1.5; 
                          break;                          
                      case 4:
                          $overweight = $sfb08* 1; 
                          break;
                      case 5:
                          $overweight = $sfb08* 2; 
                          break;
                      case 6:
                          $overweight = $sfb08* 3; 
                          break;
                      default:
                  }                                                                              
              }    
              
              
                                                        
              
              //要寫到tc_dex_file要依訂單項次的金屬來寫
              //取出本訂單的出貨單
              $soga="select oga01, to_char(oga02,'yy/mm/dd') oga02,oga04,ogaud02, oga16 from oga_file where oga16='$sfb22' ";
              $erp_sqloga=oci_parse($erp_conn1,$soga);
              $rs4=oci_execute($erp_sqloga);
              $rowoga=oci_fetch_array($erp_sqloga, OCI_ASSOC);  
              $oga01=$rowoga['OGA01'];     
              $oga02=$rowoga['OGA02'];      
              $oga04=$rowoga['OGA04'];   
              $ogaud02=$rowoga['OGAUD02'];  
                       
              //vd110 更新金屬重量時 以出貨單號 訂單項次 金屬品名
              //目前發現資材會晚入帳 所以要先檢查 若沒有tc_dex則新增一筆資料進去
              //110和210要分別新增
              $stcdex = "select tc_dex001 from tc_dex_file where tc_dex001='$oga01' and tc_dex002='$sfb221' and tc_dex003='$sfe07' "; 
              $erp_sqltcdex=oci_parse($erp_conn1,$stcdex);
              $rs5=oci_execute($erp_sqltcdex);
              $rowtcdex=oci_fetch_array($erp_sqltcdex, OCI_ASSOC);
              if (is_null($rowtcdex['TC_DEX001'])) {
                  if ($isdelay=='Y') {
                      //新增tc_dex_file時, tc_dex006是單價 要取得單價
                      //$occ01 客戶代碼
                      $sxmf="select xmf07 from xmf_file, occ_file where xmf03='$sfe07' and xmf01=occ44 and occ01='$occ01' order by xmf05 desc ";
                      $erp_sqlxmf=oci_parse($erp_conn1,$sxmf);
                      $rsxmf=oci_execute($erp_sqlxmf);
                      $rowxmf=oci_fetch_array($erp_sqlxmf, OCI_ASSOC);  
                      $price=$rowxmf['XMF07'];
                      if(is_null($price)) {
                              $price=0;
                          } else {
                              $price=floatval($price);
                          }
                       
                      ///INSERT INTO tc_dex_file VALUES (l_tc_ext.*,l_oga02,l_oga04,l_ta_oea006,'','','','','','','','','')
                      $stcdex1 = "insert into tc_dex_file values ('$oga01','$sfb221','$sfe07',$tasfe,'G',$price,$tasfe*$price,'2','VD110','VD110','Y', ".
                                 "to_date('$oga02','yy/mm/dd'), to_date('$oga02','yy/mm/dd'),'07050005','660000','','07050005','660000',to_date('$oga02','yy/mm/dd'),'$oga04','$ogaud02','',$overweight,'','','','','','','erp_addmetalweight')";
                      $erp_sqltcdex1=oci_parse($erp_conn1,$stcdex1);
                      $rs6=oci_execute($erp_sqltcdex1);
                      if ($rs6) {
                          $msg.="新增成功!!";  
                      } else {
                          $msg.="VD110 tc_dex_file $oga01 重量新增失敗!!";  
                      }
                  } else {
                          $msg.='VD110 ' . $oga01.'--'.$sfb221."重量更新失敗!!";  
                  }     
              }  else {
                  $stcdex2 = "update tc_dex_file set tc_dex004=$tasfe, tc_dex007=$tasfe*tc_dex006, tc_dex020='$tcdex020', tc_dex013=$overweight where tc_dex001='$oga01' and tc_dex002='$sfb221' and tc_dex003='$sfe07' ";
                  $erp_sqltcdex2=oci_parse($erp_conn1,$stcdex2);
                  $rs7=oci_execute($erp_sqltcdex2);
                  if ($rs7) {
                      $msg.="更新成功!!";  
                  } else {
                      $msg.="VD110 tc_dex_file $oga01 重量更新失敗!!";  
                  }   
              }
              
              //vd210
              //大部份都是還沒審核過帳 抛到210去 基本上不能新增到vd210
              //但若vd210有出貨單 就要新增到vd210
              $soga="select oga01, to_char(oga02,'yy/mm/dd') oga02,oga04,ogaud02, oga16 from oga_file where oga16='$sfb22' ";
              $erp_sqloga=oci_parse($erp_conn2,$soga);
              $rs4=oci_execute($erp_sqloga);
              $rowoga=oci_fetch_array($erp_sqloga, OCI_ASSOC);  
              $oga01=$rowoga['OGA01'];     
              $oga02=$rowoga['OGA02'];      
              $oga04=$rowoga['OGA04'];   
              $ogaud02=$rowoga['OGAUD02'];  
               
              if (is_null($oga01)) {  //若210找不到出貨就不用動作
              } ELSE {        
                  //vd110 更新金屬重量時 以出貨單號 訂單項次 金屬品名
                  //目前發現資材會晚入帳 所以要先檢查 若沒有tc_dex則新增一筆資料進去
                  //110和210要分別新增
                  $stcdex = "select tc_dex001 from tc_dex_file where tc_dex001='$oga01' and tc_dex002='$sfb221' and tc_dex003='$sfe07' "; 
                  $erp_sqltcdex=oci_parse($erp_conn2,$stcdex);
                  $rs5=oci_execute($erp_sqltcdex);
                  $rowtcdex=oci_fetch_array($erp_sqltcdex, OCI_ASSOC);
                  if (is_null($rowtcdex['TC_DEX001'])) {
                      if ($isdelay=='Y') {
                          //新增tc_dex_file時, tc_dex006是單價 要取得單價
                          //$occ01 客戶代碼
                          $sxmf="select xmf07 from xmf_file, occ_file where xmf03='$sfe07' and xmf01=occ44 and occ01='$occ01' order by xmf05 desc ";
                          $erp_sqlxmf=oci_parse($erp_conn1,$sxmf);
                          $rsxmf=oci_execute($erp_sqlxmf);
                          $rowxmf=oci_fetch_array($erp_sqlxmf, OCI_ASSOC);  
                          $price=$rowxmf['XMF07'];
                          if(is_null($price)) {
                              $price=0;
                          } else {
                              $price=floatval($price);
                          }
                          ///INSERT INTO tc_dex_file VALUES (l_tc_ext.*,l_oga02,l_oga04,l_ta_oea006,'','','','','','','','','')
                          $stcdex1 = "insert into tc_dex_file values ('$oga01','$sfb221','$sfe07',$tasfe,'G',$price,$tasfe*$price,'2','VD110','VD110','Y', ".
                                     "to_date('$oga02','yy/mm/dd'), to_date('$oga02','yy/mm/dd'),'07050005','660000','','07050005','660000',to_date('$oga02','yy/mm/dd'),'$oga04','$ogaud02','',$overweight,'','','','','','','erp_addmetalweight')";
                          $erp_sqltcdex1=oci_parse($erp_conn2,$stcdex1);
                          $rs6=oci_execute($erp_sqltcdex1);
                          if ($rs6) {
                              $msg.="新增成功!!";  
                          } else {
                              $msg.="VD210 tc_dex_file $oga01 重量更新失敗!!";  
                          }     
                      } else {
                              $msg.='VD210 ' .$oga01.'--'.$sfb221."重量更新失敗!!";                         
                      }
                  }  else {
                      $stcdex2 = "update tc_dex_file set tc_dex004=$tasfe, tc_dex007=$tasfe*tc_dex006, tc_dex020='$tcdex020', tc_dex013=$overweight where tc_dex001='$oga01' and tc_dex002='$sfb221' and tc_dex003='$sfe07' ";
                      $erp_sqltcdex2=oci_parse($erp_conn2,$stcdex2);
                      $rs7=oci_execute($erp_sqltcdex2);
                      if ($rs7) {
                          $msg.="更新成功!!";  
                      } else {
                          $msg.="VD210 tc_dex_file $oga01 重量更新失敗!!";  
                      }   
                  }
              }
          }
  } 
      msg($msg);     
      forward("erp_addmetalweight.php");                                                              
  }
  
  include("_header.php");
?>  
<link href="css.css" rel="stylesheet" type="text/css">
<p>輸入研磨後金屬重量. </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">
            出貨日期:   
            <input name="bdate" type="text" id="bdate" size="12" readonly="true" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> &nbsp;&nbsp; 
            客戶: 
            <select name="occ01" id="occ01">  
                <?
                  $s1= "select occ01,occ02 from occ_file order by occ01 ";
                  $erp_sql1 = oci_parse($erp_conn1,$s1 );
                  oci_execute($erp_sql1);  
                  while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                      echo "<option value=" . $row1["OCC01"];  
                      if ($_GET["occ01"] == $row1["OCC01"]) echo " selected";                  
                      echo ">" . $row1['OCC01'] ."--" .$row1["OCC02"] . "</option>"; 
                  }   
                ?>
            </select>  
            &nbsp;&nbsp;  
            客戶種類: 
              <input name="clienttype" type="radio" value="1" id="clienttype1" <?if($clienttype=='1') echo " checked";?>><label for="clienttype1">一般客戶 </label>&nbsp; 
              <input name="clienttype" type="radio" value="3" id="clienttype3" <?if($clienttype=='3') echo " checked";?>><label for="clienttype3">澳門客戶群 </label>   
              <input name="clienttype" type="radio" value="4" id="clienttype3" <?if($clienttype=='4') echo " checked";?>><label for="clienttype4">HK客戶群 </label>&nbsp;&nbsp;&nbsp;&nbsp;  
            是否倉庫發料比出貨晚:
            <input name="isdelay" type=checkbox id="isdelay" value="Y" <?if ($_GET['isdelay']=='Y') echo ' checked';?>></input>
            &nbsp;&nbsp;  
            <input type="submit" name="Submit2" value="送出" />
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
    <th>RX #</th>   
    <th>工單號碼</th>  
    <th>金屬代號</th>      
    <th>金屬名稱</th> 
    <th>領用日期</th>  
    <th style="text-align:right" >退料量</th>   
    <th style="text-align:right" >退料損耗</th>  
    <th style="text-align:right" >鑄造後重量</th>  
    <th style="text-align:right" >研磨後重量</th>                                                                            
  </tr>            
  <?      
      //發料 
      $ssfe1="select sfbud02, sfe01, sfe02, sfe28, sfb05, to_char(sfe04,'mm-dd-yyyy') sfe04, sfe07, ima02, imaud10, sfe16 outqty, ta_sfe002, 0 inqty, 0 lostqty, 1 sfetype from sfe_file, sfb_file, ima_file " . 
             "where sfe06='1' and sfe01=sfb01 and sfe07=ima01 and imaud10 is not null " .
             "and sfe01 in ( select tc_ogb002 from tc_ogb_file, tc_oga_file,oga_file where tc_ogb001=tc_oga001 $occfilter and tc_oga002=to_date('$bdate','yy/mm/dd') and tc_ogb003=oga16 ) "; 
      //退料                                                                                                                    
      $ssfe2="select sfbud02, sfe01, sfe02, sfe28, sfb05, to_char(sfe04,'mm-dd-yyyy') sfe04, sfe07, ima02, imaud10, 0 outqty , 0 ta_sfe002, sfe16 inqty, ta_sfe003 lostqty, 2 sfetype from sfe_file, sfb_file, ima_file " . 
             "where sfe06='4' and sfe01=sfb01 and sfe07=ima01 and imaud10 is not null " .
             "and sfe01 in ( select tc_ogb002 from tc_ogb_file, tc_oga_file, oga_file where tc_ogb001=tc_oga001 $occfilter and tc_oga002=to_date('$bdate','yy/mm/dd') and tc_ogb003=oga16 ) "; 
      $ssfe= "select sfbud02, sfe01, sfe02, sfe28, sfb05, sfe04, sfe07, ima02, imaud10, outqty, ta_sfe002, inqty, lostqty, sfetype from  " .
               "( select * from ($ssfe1) union all ($ssfe2) ) " .
             "order by sfbud02, sfetype, sfe04";
      $erp_sqlsfe = oci_parse($erp_conn1,$ssfe );
      oci_execute($erp_sqlsfe); 
      while ($rowsfe = oci_fetch_array($erp_sqlsfe, OCI_ASSOC)) { 
        $key=$rowsfe['SFE02'].$rowsfe['SFE28'];   
        //要算出金屬重量計算方式
        if ($rowsfe['IMAUD10']==1) {
            $metaltype='鋼牙';
        } else if ($rowsfe['IMAUD10']==2) {
            $metaltype='CNC B';
        } else  if ($rowsfe['IMAUD10']==3) {
            $metaltype='CNC C';
        } else if ($rowsfe['IMAUD10']==4) {
            $metaltype='UC 磁牙';
        } else if ($rowsfe['IMAUD10']==5) {
            $metaltype='NF';
        } else {
            $metaltype='未設定';
        }  
        $tasfe002=$rowsfe['TA_SFE002'];
      ?>    
	    <tr bgcolor="#FFFFFF"> 
		      <td><img src="i/arrow.gif" width="16" height="16">  
              <input type="hidden" name="sfearray[]"      value="<?=$key;?>"> 
              <input type="hidden" name="sfe01<?=$key;?>" value="<?=$rowsfe['SFE01'];?>">   
              <input type="hidden" name="sfe02<?=$key;?>" value="<?=$rowsfe['SFE02'];?>">    
              <input type="hidden" name="sfe07<?=$key;?>" value="<?=$rowsfe['SFE07'];?>">        
              <input type="hidden" name="sfe28<?=$key;?>" value="<?=$rowsfe['SFE28'];?>">       
              <input type="hidden" name="sfb05<?=$key;?>" value="<?=$rowsfe['SFB05'];?>">  
              <input type="hidden" name="outqty<?=$key;?>" value="<?=$rowsfe['OUTQTY'];?>">     
              <input type="hidden" name="imaud10<?=$key;?>" value="<?=$rowsfe['IMAUD10'];?>">      
          </td>             
			    <td><?=$rowsfe["SFBUD02"];?></td>
          <td><?=$rowsfe["SFE01"];?></td>   
          <td><?=$rowsfe["SFE07"];?></td> 
          <td><?=$rowsfe["IMA02"];?></td>  
          <td><?=$rowsfe["SFE04"];?></td>  
          <td style="text-align:right" ><?=number_format($rowsfe["INQTY"],2,'.',',');?></td>    
          <td style="text-align:right" ><?=number_format($rowsfe["LOSTQTY"],2,'.',',');?></td>   
          <td style="text-align:right" > <?=number_format($rowsfe["OUTQTY"],2,'.',',');?> </td>
          <? if ($rowsfe['SFETYPE']==1) { ?>            
                <td style="text-align:right" ><input name="tasfe002<?=$key;?>" type="text"  id="tasfe002<?=$key;?>" value=<?=number_format($tasfe002,2,'.',',');?> style="text-align:right" onmouseover="this.focus()" onfocus="this.select()" onkeypress="return numberOnly(event,'sfe<?=$key;?>')" ></td>  
          <? } else { ?>
                <td>&nbsp;&nbsp; 
                    <input type="hidden" name="tasfe002<?=$key;?>" value=0>  
                </td>   
          <? } ?>      
      </tr> 
      <?  
      }   
  ?>    
  <tr>
    <td>&nbsp;</td>
    <td><input type="hidden" name="bdate"  value="<?=$bdate;?>">    
        <input type="hidden" name="occ01"  value="<?=$occ01;?>">           
        <input type="hidden" name="isdelay"  value="<?=$_GET['isdelay'];?>">            
        <input type="hidden" name="action" value="save">
        <input type="submit" name="Submit" value="更新">        
    </td>
  </tr>
</table>  
</form>
