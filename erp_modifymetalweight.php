<?
  session_start();
  $pagtitle = "業務部 &raquo; 修改研磨後重量"; 
  include("_data.php");
  include("_erp.php");
  auth("erp_modifymetalweight.php");
  
  
  $rxno=$_GET['rxno'];
     
  //修改後的重量要回寫到invoice中   
  if ($_POST["action"] == "save") {            
      $msg=''; 
             
      foreach ($_POST["sfearray"] as $sfe)   {
          if ($_POST["risok" . $sfe] == "Y") {                   
              $sfe01=$_POST['sfe01'.$sfe];      
              $sfe02=$_POST['sfe02'.$sfe];  
              $sfe28=$_POST['sfe28'.$sfe];  
              $tasfe002=$_POST['tasfe002'.$sfe];
              $imaud10=$_POST['imaud10'.$sfe]; 
              //先將要修改的資料記錄下來                                    
              $queryap   = "insert into erp_addmetalweight_updaterecord ( bdate, occ01, sfe02, sfe28, weight, username, ip ) values (
                          '" . 'modify'                           . "',     
                          '" . safetext($rxno)                    . "',  
                          '" . safetext($sfe02)                   . "',      
                          '" . safetext($sfe28)                   . "',     
                          '" . safetext($tasfe002)                . "',       
                          '" . safetext($_SESSION['account'])     . "',       
                          '" . safetext($_SERVER['REMOTE_ADDR'])  . "')";
              $resultap= mysql_query($queryap) or die ('31 erp_addmetalweight_updaterecord added error. ' .mysql_error());
              
              
              //取出客戶是否為CNC, UC    
              $socc= "select occud10 from occ_file, oea_file, sfb_file  where occ01=oea04 and oea01=sfb22 and sfb01='$sfe01'"; 
              $erp_sqlocc = oci_parse($erp_conn1,$socc );
              oci_execute($erp_sqlocc); 
              $rowocc=oci_fetch_array($erp_sqlocc, OCI_ASSOC);
              $occtype=$rowocc['OCCUD10']; //1:其他 2:CNC 3:UC    
               
              $ssfe1= "update sfe_file set ta_sfe002=$tasfe002 where sfe02='$sfe02' and sfe28='$sfe28'";
              $erp_sqlsfe1=oci_parse($erp_conn1,$ssfe1);
              $rs1=oci_execute($erp_sqlsfe1);
              if ($rs1) {
                  $msg.="VD110 sfe_file發料檔 $sfe02 $sfe28 更新成功!!";  
              } else {
                  $msg.="VD110 sfe_file發料檔 $sfe02 $sfe28 更新失敗!!";  
              }  
                                                                                                                                               
              //合計同一張工單的所有研磨後重量  退料量  損耗量      //研磨後重量                         //退料要再扣掉sfe16                        //退料要扣掉損耗損耗
              $ssfe2="select sfe01, sfe07, sum(decode(ta_sfe002, null, 0, ta_sfe002) - decode(sfe06, '4', decode(sfe16, null, 0, sfe16), 0 ) - decode(sfe06, '4', decode(ta_sfe003, null, 0, ta_sfe003),0) ) ta_sfe from sfe_file where sfe01='$sfe01' group by sfe01, sfe07";
              $erp_sqlsfe2=oci_parse($erp_conn1,$ssfe2);
              $rs2=oci_execute($erp_sqlsfe2);
              while ($rowsfe2=oci_fetch_array($erp_sqlsfe2, OCI_ASSOC)) {  //同一張工單 可能有兩種金屬以上 
                  $tasfe=$rowsfe2['TA_SFE']; 
                  $sfe07=$rowsfe2['SFE07'];
              
                  $ssfb="select sfb08, sfb22, sfb221 from sfb_file where sfb01='$sfe01'";
                  $erp_sqlsfb=oci_parse($erp_conn1,$ssfb);
                  $rs3=oci_execute($erp_sqlsfb);
                  $rowsfb=oci_fetch_array($erp_sqlsfb, OCI_ASSOC); 
                  $sfb08=$rowsfb['SFB08'];  
                  $sfb22=$rowsfb['SFB22']; 
                  $sfb221=$rowsfb['SFB221']; 
                  
                  //把原始資料寫起來 供檢杳用
                  $tcdex020='客:'. $occtype . ' 金:' . $imaud10 . ' 重:' . $tasfe . ' 顆:' . $sfb08;   
                  
                  if ($occtype=='2') {          //CNC
                      if ($imaud10=='3') {      //3:B 要找到顆數 每顆以0.9收                        
                          $tasfe= $sfb08 * 0.9;   
                      } else if ($imaud10=='4') {    // 4:C 要找到顆數 每顆以1.1收     
                          $tasfe= $sfb08 * 1.1;
                      } else {                  //其餘收15%
                          $tasfe *= 1.15;    
                      }
                  } else if ($occtype=='3') {   //UC 磁牙收10% 因此先判斷鋼牙 剩下為磁牙 
                      if ($imaud10=='1') {
                          $tasfe *= 1.15;   
                      } else {
                          $tasfe *= 1.10;     
                      }    
                  } else {                      //其餘都加收15%
                          $tasfe *= 1.15;               
                  }
                  
                  //要寫到tc_dex_file要依訂單項次的金屬來寫
                  //取出本訂單的出貨單
                  $soga="select oga01 from oga_file where oga16='$sfb22' ";
                  $erp_sqloga=oci_parse($erp_conn1,$soga);
                  $rs4=oci_execute($erp_sqloga);
                  $rowoga=oci_fetch_array($erp_sqloga, OCI_ASSOC);  
                  $oga01=$rowoga['OGA01'];         
                           
                  //vd110 更新金屬重量時 以出貨單號 訂單項次 金屬品名
                  $stcdex = "update tc_dex_file set tc_dex004=$tasfe, tc_dex007=$tasfe*tc_dex006, tc_dex020='$tcdex020' where tc_dex001='$oga01' and tc_dex002='$sfb221' and tc_dex003='$sfe07' ";
                  $erp_sqltcdex=oci_parse($erp_conn1,$stcdex);
                  $rs5=oci_execute($erp_sqltcdex);
                  if ($rs5) {
                      $msg.="VD110 tc_dex_file $oga01 重量更新成功!!";  
                  } else {
                      $msg.="VD110 tc_dex_file $oga01 重量更新失敗!!";  
                  } 
                  
                  //
                  //取出vd210本訂單的出貨單
                  $soga="select oga01 from oga_file where oga16='$sfb22' ";
                  $erp_sqloga=oci_parse($erp_conn2,$soga);
                  $rs4=oci_execute($erp_sqloga);
                  $rowoga=oci_fetch_array($erp_sqloga, OCI_ASSOC);  
                  $oga01=$rowoga['OGA01'];   
                  
                  //  //修改後的重量要回寫到VD210 的 invoice中        
                  $stcofb = "update tc_ofb_file set tc_ofb12=$tasfe, tc_ofb14=$tasfe*tc_ofb13, tc_ofb14t=$tasfe*tc_ofb13  where tc_ofb31='$oga01' and tc_ofb32='$sfb221' and tc_ofb04='$sfe07' ";
                  $erp_sqltcofb=oci_parse($erp_conn2,$stcofb);
                  $rs6=oci_execute($erp_sqltcofb);
                  if ($rs6) {
                      $msg.="VD110 tc_ofb_file $oga01 Invoice更新成功!!";  
                  } else {
                      $msg.="VD110 tc_ofb_file $oga01 Invoice更新失敗!!";  
                  }                         
              }         
          }        
      } 
      
      msg($msg);     
      forward("erp_modifymetalweight.php");                                                              
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
            RX #:  
            <input name="rxno" type="text" id="rxno" size="20" value=<?=$_GET['rxno'];?> > &nbsp;&nbsp;      
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
    <th width="16">&nbsp;&nbsp;</th>                                                                                
  </tr>            
  <?      
      //發料 
      $ssfe1="select sfbud02, sfe01, sfe02, sfe28, to_char(sfe04,'mm-dd-yyyy') sfe04, sfe07, ima02, imaud10, sfe16 outqty, ta_sfe002, 0 inqty, 0 lostqty, 1 sfetype from sfe_file, sfb_file, ima_file " . 
             "where sfe06='1' and sfe01=sfb01 and sfbud02='$rxno' and sfe07=ima01 and imaud10 is not null " ;
      //退料                                                                                                                    
      $ssfe2="select sfbud02, sfe01, sfe02, sfe28, to_char(sfe04,'mm-dd-yyyy') sfe04, sfe07, ima02, imaud10, 0 outqty , 0 ta_sfe002, sfe16 inqty, ta_sfe003 lostqty, 2 sfetype from sfe_file, sfb_file, ima_file " . 
             "where sfe06='4' and sfe01=sfb01 and sfbud02='$rxno' and sfe07=ima01 and imaud10 is not null " ;
      $ssfe= "select sfbud02, sfe01, sfe02, sfe28, sfe04, sfe07, ima02, imaud10, outqty, ta_sfe002, inqty, lostqty, sfetype from  " .
               "( select * from ($ssfe1) union all ($ssfe2) ) " .
             "order by sfbud02, sfetype";
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
            $metaltype='其他磁牙';
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
              <input type="hidden" name="sfe28<?=$key;?>" value="<?=$rowsfe['SFE28'];?>">       
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
                <td style="text-align:right" ><input name="tasfe002<?=$key;?>" type="text"  id="tasfe992<?=$key;?>" value=<?=number_format($tasfe002,2,'.',',');?> style="text-align:right" onmouseover="this.focus()" onfocus="this.select()" onkeypress="return numberOnly(event,'sfe<?=$key;?>')" ></td>  
                <td width=16><input name="risok<?=$key;?>" type="checkbox" id="risok<?=$key;?>" value="Y" </td>  
          <? } else { ?>
                <td>&nbsp;&nbsp; 
                    <input type="hidden" name="tasfe002<?=$key;?>" value=0>  
                </td>   
                <td>&nbsp;&nbsp; </td>  
          <? } ?>      
      </tr> 
      <?  
      }   
  ?>    
  <tr>
    <td>&nbsp;</td>
    <td><input type="hidden" name="rxno"  value="<?=$rxno;?>">                   
        <input type="hidden" name="action" value="save">
        <input type="submit" name="Submit" value="更新">        
    </td>
  </tr>
</table>  
</form>
