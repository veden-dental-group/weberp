<?php
function changeteethno($teethno, $type){  
    if ($type==2) {  //1:美洲齒位
        $no[11]='8';  
        $no[12]='7';  
        $no[13]='6';  
        $no[14]='5';  
        $no[15]='4';  
        $no[16]='3';  
        $no[17]='2';  
        $no[18]='1';
        
        $no[21]='9';  
        $no[22]='10';  
        $no[23]='11';  
        $no[24]='12';  
        $no[25]='13';  
        $no[26]='14';  
        $no[27]='15';  
        $no[28]='16';
        
        $no[31]='24';  
        $no[32]='23';  
        $no[33]='22';  
        $no[34]='21';  
        $no[35]='20';  
        $no[36]='19';  
        $no[37]='18';  
        $no[38]='17';
        
        $no[41]='25';  
        $no[42]='26';  
        $no[43]='27';  
        $no[44]='28';  
        $no[45]='29';  
        $no[46]='30';  
        $no[47]='31';  
        $no[48]='32';
        
        $oldteethno=explode('|',$teethno); 
        $newteethno='';
        $j= count($oldteethno);
        for ($i=0; $i<$j; $i++){
          $k=$oldteethno[$i];
          $oldteethno[$i]=$no[$k]; 
        }
        $newteethno=implode('|',$oldteethno);
    } else if ($type==3) {  //美洲
          
        $no[11]='1';  
        $no[12]='2';  
        $no[13]='3';  
        $no[14]='4';  
        $no[15]='5';  
        $no[16]='6';  
        $no[17]='7';  
        $no[18]='8';
        
        $no[21]='1';  
        $no[22]='2';  
        $no[23]='3';  
        $no[24]='4';  
        $no[25]='5';  
        $no[26]='6';  
        $no[27]='7';  
        $no[28]='8';
        
        $no[31]='1';  
        $no[32]='2';  
        $no[33]='3';  
        $no[34]='4';  
        $no[35]='5';  
        $no[36]='6';  
        $no[37]='7';  
        $no[38]='8';
        
        $no[41]='1';  
        $no[42]='2';  
        $no[43]='3';  
        $no[44]='4';  
        $no[45]='5';  
        $no[46]='6';  
        $no[47]='7';  
        $no[48]='8';
        $oldteethno=explode('|',$teethno); 
        $newteethno='';
        $j= count($oldteethno);
        for ($i=0; $i<$j; $i++){
          $k=$oldteethno[$i];
          $oldteethno[$i]=$no[$k]; 
        }
        $newteethno=implode('|',$oldteethno);
    } else {  //沒註明的一律不轉
        $newteethno=$teethno; 
    }        
    return($newteethno);
}


function getgem02($gen01, $erp_conn1){  
    $sgem="select gem02 from gen_file, gem_file where gen01='$gen01' and gen03=gem01 "; 
    $erp_sqlgem = oci_parse($erp_conn1,$sgem );
    oci_execute($erp_sqlgem);  
    $rowgem = oci_fetch_array($erp_sqlgem, OCI_ASSOC);
    return($rowgem['GEM02']);
}   

function findcasewithrxno($rxno, $erp_conn1, $erp_conn2, $bdate=''){
    //檢查vd110有無訂單 
    $msg='';
    $rxno=trim($rxno);
    $soea110="select oea01, to_char(oea02,'mm-dd-yyyy') oea02, ta_oea004, occ01, occ02, ta_oea006, to_char(ta_oea005,'mm-dd-yyyy') ta_oea005 from oea_file, occ_file where ta_oea006 like '%$rxno%' and  oea04=occ01 and oea02 > to_date('$bdate','yyyy-mm-dd') ";
    $erp_sqloea110 = oci_parse($erp_conn1,$soea110 );
    oci_execute($erp_sqloea110);  
    $nrowoea110 = oci_fetch_all($erp_sqloea110, $results);    
    if ($nrowoea110==0) { //VD110沒有 檢查VD210有沒有
        $soea210="select oea01, to_char(oea02,'mm-dd-yyyy') oea02, occ01, occ02, ta_oea006, to_char(ta_oea005,'mm-dd-yyyy') ta_oea005 from oea_file, occ_file where ta_oea006 like '%$rxno%' and oea04=occ01 and oea02 > to_date('$bdate','yyyy-mm-dd') ";
        $erp_sqloea210 = oci_parse($erp_conn2,$soea210 );
        oci_execute($erp_sqloea210);  
        $nrowoea210 = oci_fetch_all($erp_sqloea210, $results); 
        if ($nrowoea210==0){
            $msg.= "RX:" . $rxno . " VD210, VD110均無訂單!!<br>";
        } else {
            for ($i = 0; $i < $nrowoea210; $i++) {   //vd210可能有N筆相同的訂單 都未審核     
                $msg.= "RX:" . $results['TA_OEA006'][$i] . '...' . $results['OCC01'][$i] . ':' . $results['OCC02'][$i] . 
                       "於" . $results['OEA02'][$i] . "錄入, 在VD210有訂單, 單號:" . $results['OEA01'][$i] . ", 但未審核過帳至VD110中.<br>";
            }
        }   
    } else {
        for ($i = 0; $i < $nrowoea110; $i++) {   //vd110可能有N筆相同的訂單
            $ta_oea004=$results['TA_OEA004'][$i];
            $oea01=$results['OEA01'][$i]; //order no
            $prefix = "RX:" . $results['TA_OEA006'][$i] . '(訂單:'.$oea01.')' . $results['OCC02'][$i] . " 於 " . $results['OEA02'][$i] . " 錄入, 預計出貨日期:" . $results['TA_OEA005'][$i] . ",  " ;
            $msg .= findcasewithoea01($oea01, $prefix, $ta_oea004, $erp_conn1,$erp_conn2);
        } 
    }
    return($msg);
}        

function findcasewithoea01($oea01, $prefix, $ta_oea004, $erp_conn1, $erp_conn2){
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
        //$ssfb="select sfb01, ima01, ima02, gem02 from sfb_file, ima_file, gem_file where sfb22='$oea01' and sfb05=ima01 and sfb82=gem01 and ta_ima005='N' ";
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
                $where   = findcasewithsfb01($sfb01,$gem02,$ta_oea004,$erp_conn1,$erp_conn2);
                $msga[$i] = $prefix . " " . $product ." " . $where . "<br>"; 
            }       
            for($i=0; $i<count($msga); $i++){  
              $msg .= $msga[$i];
            }  
        }
    }
    return($msg);
}  

function findcasewithsfb01($sfb01,$gem02,$ta_oea004,$erp_conn1, $erp_conn2){ 
    $msg='';
    //沒有出貨單記錄 檢查秤重
    $stcoga="select to_char(tc_oga002,'mm-dd-yyyy') tc_oga002 from tc_ogb_file,tc_oga_file where tc_ogb002='$sfb01' and tc_ogb001=tc_oga001 ";
    $erp_sqltcoga = oci_parse($erp_conn1,$stcoga );
    oci_execute($erp_sqltcoga);  
    $rowtcoga = oci_fetch_array($erp_sqltcoga, OCI_ASSOC);
    if (!is_null($rowtcoga['TC_OGA002'])) {
        $msg=" 於 " . $rowtcoga['TC_OGA002'] ." 掃描秤重, 但未產生出貨單, 請執行 csft998 中的 生成出貨單 功能.";
    } else {
        //沒有秤重記錄 檢查有無入庫  
        $ssfu="select to_char(sfu02,'mm-dd-yyyy') sfu02 from sfv_file, sfu_file where sfv11='$sfb01' and sfv01=sfu01 ";
        $erp_sqlsfu = oci_parse($erp_conn1,$ssfu );
        oci_execute($erp_sqlsfu);  
        $rowsfu = oci_fetch_array($erp_sqlsfu, OCI_ASSOC);
        if (!is_null($rowsfu['SFU02'])) {
            $msg=" 於 " . $rowsfu['SFU02'] ." 入庫完畢, 但未掃描秤重, 請執行 csft998 掃描秤重作業.";   
        } else { 
           if ($ta_oea004=='3'){ //客戶返工的case 若無入庫, 則只有一道工序 9999
                $msg=" 客戶返修case, 在 $gem02 製作中, 但無法判斷工序0." ;
           } else {
                //檢查工單是否已結案 
                $ssfb="select sfb04, to_char(sfb38,'mm-dd-yyyy') sfb38 from sfb_file where sfb01='$sfb01' ";
                $erp_sqlsfb = oci_parse($erp_conn1,$ssfb );
                oci_execute($erp_sqlsfb);  
                $rowsfb = oci_fetch_array($erp_sqlsfb, OCI_ASSOC);
                if ($rowsfb['SFB04']=='8') {
                    $msg=" 工單已於 ". $rowsfb['SFB38'] . " 結案.";    
                } else {  
                    // 有工單 未入庫 找在哪一道工序有在製量
                    $stcsrg="select tc_srg001, tc_srg002, tc_srg003, tc_srg004, tc_srg005, tc_srg006, to_char(tc_srg007,'mm-dd-yyyy') tc_srg007, tc_srg008, tc_srg009, " .
                            "to_char(tc_srg010,'mm-dd-yyyy') tc_srg010, tc_srg011, tc_srg012, to_char(tc_srg013,'mm-dd-yyyy') tc_srg013, tc_srg014, tc_srg015, tc_srg016, tc_srg018, " .  
                            "to_char(tc_srg019,'mm-dd-yyyy') tc_srg019, tc_srg020, tc_srg021, to_char(tc_srg022,'mm-dd-yyyy') tc_srg022, tc_srg023, tc_srg024, " .
                            "to_char(tc_srg025,'mm-dd-yyyy') tc_srg025, tc_srg026, tc_srg027, tc_srg028, tc_srg029, ecb06, ecb17 " .
                            "from tc_srg_file, ecb_file where tc_srg001='$sfb01' and tc_srg003=ecb01 and tc_srg004=ecb03 and tc_srg005 is not null ";
                    $erp_sqltcsrg = oci_parse($erp_conn1,$stcsrg );
                    oci_execute($erp_sqltcsrg);  
                    $rowtcsrg = oci_fetch_array($erp_sqltcsrg, OCI_ASSOC);   
                    $ecb06=$rowtcsrg['ECB06'];     //目前作業編號
                    $ecb17=$rowtcsrg['ECB17'];     //目前作業名稱    
                    if (is_null($ecb06)){ //車間QC完 但未入庫
                        $msg="  已QC完畢, 等待秤重出貨!!";
                    } else {
                        //找看看有沒有比在製量更小的工序 以判斷是不是第一道工序, 若是第一道工序, 看是否返工 以找上一道的時間
                        $ecb01=$rowtcsrg['TC_SRG003'];   // product code
                        $ecb03=$rowtcsrg['TC_SRG004'];   // sequential code  
                        $secb="select ecb06, ecb17 from ecb_file where ecb01='$ecb01' and ecb03<'$ecb03' ";
                        $erp_sqlecb = oci_parse($erp_conn1,$secb );
                        oci_execute($erp_sqlecb);  
                        $rowecb = oci_fetch_array($erp_sqlecb, OCI_ASSOC);
                        if (is_null($rowecb['ECB06'])) {   //找不到更小工序 表示本站為第一道工序   判斷有無入站 
                            if ($rowtcsrg['TC_SRG018']=="Y"){    //為返工重做  找出由哪裡來
                                if (!is_null($rowtcsrg['TC_SRG022'])) {  //有出站 有在製量 表示在PQC中  顯示出站日期 時間
                                    $gem02=getgem02($rowtcsrg['TC_SRG024'],$erp_conn1);
                                    $outdate=$rowtcsrg['TC_SRG022'];
                                    $outtime=$rowtcsrg['TC_SRG023'];
                                    $msg="  於 $outdate $outtime 在 $gem02 $ecb06 $ecb17 出站 QC中1." ; 
                                } else { 
                                    if (!is_null($rowtcsrg['TC_SRG019'])) {  //有進站 未出站
                                        $gem02=getgem02($rowtcsrg['TC_SRG021'],$erp_conn1);  
                                        $outdate=$rowtcsrg['TC_SRG019'];
                                        $outtime=$rowtcsrg['TC_SRG020'];
                                        $msg="  於 $outdate $outtime 在 $gem02 $ecb06 $ecb17 進站製作中2." ;   
                                    } else {
                                        //未入站 找到最小的QC站 一定由最小的QC站返工
                                        $stcsrg="select tc_srg001, tc_srg002, tc_srg003, tc_srg004, tc_srg005, tc_srg006, to_char(tc_srg007,'mm-dd-yyyy') tc_srg007, tc_srg008, tc_srg009, " .
                                                    "to_char(tc_srg010,'mm-dd-yyyy') tc_srg010, tc_srg011, tc_srg012, to_char(tc_srg013,'mm-dd-yyyy') tc_srg013, tc_srg014, tc_srg015, tc_srg016, tc_srg018, " .  
                                                    "to_char(tc_srg019,'mm-dd-yyyy') tc_srg019, tc_srg020, tc_srg021, to_char(tc_srg022,'mm-dd-yyyy') tc_srg022, tc_srg023, tc_srg024, " .
                                                    "to_char(tc_srg025,'mm-dd-yyyy') tc_srg025, tc_srg026, tc_srg027, tc_srg028, tc_srg029, ecb06, ecb07 " . 
                                                    "from tc_srg_file, ecb_file " . 
                                                    "where tc_srg001='$sfb01' and tc_srg003=ecb01 and tc_srg004=ecb03 and tc_srg006='Y' order by ecb03 asc ";   
                                        $erp_sqltcsrg = oci_parse($erp_conn1,$stcsrg );
                                        oci_execute($erp_sqltcsrg);  
                                        $rowtcsrg = oci_fetch_array($erp_sqltcsrg, OCI_ASSOC);
                                        $outdate=$rowtcsrg['TC_SRG025'];
                                        $outtime=$rowtcsrg['TC_SRG026']; 
                                        $ecb06=$rowtcsrg['ECB06'];     //作業編號
                                        $ecb17=$rowtcsrg['ECB17'];     //作業名稱     
                                        $gem02=getgem02($rowtcsrg['TC_SRG012'],$erp_conn1);       
                                        $msg="  返工case 於 $outdate $outtime 由 $gem02 $ecb06 $ecb17 QC返工3." ;  
                                    }                                   
                                }
                            } else {     //非返工的第一道工序
                                if (!is_null($rowtcsrg['TC_SRG010'])) {  //有出站 有在製量 表示在PQC中
                                    $outdate=$rowtcsrg['TC_SRG010'];
                                    $outtime=$rowtcsrg['TC_SRG011'];
                                    $gem02=getgem02($rowtcsrg['TC_SRG012'],$erp_conn1);   
                                    $msg="  於 $outdate $outtime 在 $gem02 $ecb06 $ecb17 出站 QC中4." ; 
                                } else { 
                                    if (!is_null($rowtcsrg['TC_SRG007'])) {  //有進站 未出站
                                        $outdate=$rowtcsrg['TC_SRG007'];
                                        $outtime=$rowtcsrg['TC_SRG008'];
                                        $gem02=getgem02($rowtcsrg['TC_SRG009'],$erp_conn1);  
                                        $msg="  於 $outdate $outtime 在 $gem02 $ecb06 $ecb17 進站製作中5." ;   
                                    } else { 
                                        $msg="  等待 $gem02 $ecb06 $ecb17 入站中6." ;  
                                    } 
                                }   
                            }                              
                        } else { //有更小工序 表示非第一站 判斷是否返工 要讀不同欄位值     
                            if ($rowtcsrg['TC_SRG018']=="Y"){    //為返工工序  
                                if (!is_null($rowtcsrg['TC_SRG022'])) {  //返工工序 有出站日期 但仍在本站 表示是在PQC中  
                                    $outdate=$rowtcsrg['TC_SRG022'];
                                    $outtime=$rowtcsrg['TC_SRG023'];      
                                    $gem02=getgem02($rowtcsrg['TC_SRG024'],$erp_conn1);     
                                    $msg="  返工case 於 $outdate $outtime 在 $gem02 $ecb06 $ecb17 出站 QC中7." ;  
                                } else {                                   
                                    if (!is_null($rowtcsrg['TC_SRG019'])){ //返工工序 無出站日期 有入站日期  表示在本站製作中 顯示入站日期  
                                        $indate=$rowtcsrg['TC_SRG019'];
                                        $intime=$rowtcsrg['TC_SRG020'];   
                                        $gem02=getgem02($rowtcsrg['TC_SRG021'],$erp_conn1);                 
                                        $msg="  返工case 於 $indate $intime 在 $gem02 $ecb06 $ecb17 入站製作中8." ; 
                                    } else {     //返工工序 無入站日期 要找上一道工序出站時間   可能是剛返工 也可能是上站出站
                                        //找有無比本站小的返工工序 若有 表示由上一道工序而來
                                        $stcsrg1="select tc_srg001, tc_srg002, tc_srg003, tc_srg004, tc_srg005, tc_srg006, to_char(tc_srg007,'mm-dd-yyyy') tc_srg007, tc_srg008, tc_srg009, " .
                                                "to_char(tc_srg010,'mm-dd-yyyy') tc_srg010, tc_srg011, tc_srg012, to_char(tc_srg013,'mm-dd-yyyy') tc_srg013, tc_srg014, tc_srg015, tc_srg016, tc_srg018, " .  
                                                "to_char(tc_srg019,'mm-dd-yyyy') tc_srg019, tc_srg020, tc_srg021, to_char(tc_srg022,'mm-dd-yyyy') tc_srg022, tc_srg023, tc_srg024, " .
                                                "to_char(tc_srg025,'mm-dd-yyyy') tc_srg025, tc_srg026, tc_srg027, tc_srg028, tc_srg029, ecb06, ecb17 " .
                                                "from tc_srg_file, ecb_file " . 
                                                "where tc_srg001='$sfb01' and tc_srg004<'$ecb03' and tc_srg018='Y' and tc_srg003=ecb01 and tc_srg004=ecb03 order by ecb03 desc ";   
                                        $erp_sqltcsrg1 = oci_parse($erp_conn1,$stcsrg1 );
                                        oci_execute($erp_sqltcsrg1);  
                                        $rowtcsrg1 = oci_fetch_array($erp_sqltcsrg1, OCI_ASSOC);
                                        $ecb061=$rowtcsrg1['ECB06'];
                                        $ecb171=$rowtcsrg1['ECB17'];
                                        if (is_null($rowtcsrg1['TC_SRG001'])) {  //沒有更小的返工工序 表示是由最近的QC直接返工來的 要找QC出站的日期/時間
                                            $stcsrg2="select tc_srg001, tc_srg002, tc_srg003, tc_srg004, tc_srg005, tc_srg006, to_char(tc_srg007,'mm-dd-yyyy') tc_srg007, tc_srg008, tc_srg009, " .
                                                    "to_char(tc_srg010,'mm-dd-yyyy') tc_srg010, tc_srg011, tc_srg012, to_char(tc_srg013,'mm-dd-yyyy') tc_srg013, tc_srg014, tc_srg015, tc_srg016, tc_srg018, " .  
                                                    "to_char(tc_srg019,'mm-dd-yyyy') tc_srg019, tc_srg020, tc_srg021, to_char(tc_srg022,'mm-dd-yyyy') tc_srg022, tc_srg023, tc_srg024, " .
                                                    "to_char(tc_srg025,'mm-dd-yyyy') tc_srg025, tc_srg026, tc_srg027, tc_srg028, tc_srg029, ecb06, ecb07 " . 
                                                    "from tc_srg_file, ecb_file " . 
                                                    "where tc_srg001='$sfb01'' and tc_srg004>'$ecb03' and tc_srg003=ecb01 and tc_srg004=ecb03 and tc_srg006='Y' order by ecb03 asc ";   
                                            $erp_sqltcsrg2 = oci_parse($erp_conn1,$stcsrg2 );
                                            oci_execute($erp_sqltcsrg2);  
                                            $rowtcsrg = oci_fetch_array($erp_sqltcsrg2, OCI_ASSOC);
                                            $outdate2=$rowtcsrg2['TC_SRG025'];
                                            $outtime2=$rowtcsrg2['TC_SRG026'];   
                                            $ecb062=$rowtcsrg2['ECB06'];     //作業編號
                                            $ecb172=$rowtcsrg2['ECB17'];     //作業名稱      
                                            $gem02=getgem02($rowtcsrg['TC_SRG024'],$erp_conn1);       
                                            $msg="  返工case 於 $outdate2 $outtime2 由 $gem02 $ecb062 $ecb172 QC返工, 等待 $ecb06 $ecb17 入站9." ;   
                                        } else {  // 有更小工序 找出站日期 直接秀出          
                                            if ($rowtcsrg['TC_SRG006']=='N') { //判斷是由QC出站 或直接出站 
                                                $outdate1=$rowtcsrg1['TC_SRG010'];
                                                $outtime1=$rowtcsrg1['TC_SRG011'];
                                                $gem02=getgem02($rowtcsrg['TC_SRG024'],$erp_conn1);  
                                                $msg="  返工case 於 $outdate1 $outtime1 在 $gem02 $ecb06 $ecb17 出站, 等待 $ecb06 $ecb17 入站a." ;  
                                            } else {
                                                $outdate1=$rowtcsrg1['TC_SRG013'];
                                                $outtime1=$rowtcsrg1['TC_SRG014'];
                                                $gem02=getgem02($rowtcsrg['TC_SRG024'],$erp_conn1);  
                                                $msg="  返工case 於 $outdate1 $outtime1 在 $gem02 $ecb06 $ecb17 QC出站, 等待$ecb06 $ecb17 入站b." ;  
                                            }                                                                                               
                                        }       
                                    }
                                }    
                            }  else {     //不是返工 要判斷有無入站日期 若無要找上一筆的出站日期  
                                if (!is_null($rowtcsrg['TC_SRG010'])) {  //非返工工序 有出站記錄 但仍在本站 表示是在PQC中     
                                    $outdate=$rowtcsrg['TC_SRG010'];
                                    $outtime=$rowtcsrg['TC_SRG011'];  
                                    $gem02=getgem02($rowtcsrg['TC_SRG012'],$erp_conn1);                  
                                    $msg=" 於 $outdate $outtime 在 $gem02 $ecb06 $ecb17 出站 QC中c." ;  
                                } else {    
                                    if (!is_null($rowtcsrg['TC_SRG007'])){ //非返工 無出站日期 有入站日期    
                                        $indate=$rowtcsrg['TC_SRG007'];
                                        $intime=$rowtcsrg['TC_SRG008'];  
                                        $gem02=getgem02($rowtcsrg['TC_SRG009'],$erp_conn1);                    
                                        $msg="於 $indate $intime 在 $gem02 $ecb06 $ecb17 入站製作d." ; 
                                    } else {     //非返工 無入站日期 要找上一道出站工序
                                        $stcsrg3="select tc_srg001, tc_srg002, tc_srg003, tc_srg004, tc_srg005, tc_srg006, to_char(tc_srg007,'mm-dd-yyyy') tc_srg007, tc_srg008, tc_srg009, " .
                                                "to_char(tc_srg010,'mm-dd-yyyy') tc_srg010, tc_srg011, tc_srg012, to_char(tc_srg013,'mm-dd-yyyy') tc_srg013, tc_srg014, tc_srg015, tc_srg016, tc_srg018, " .  
                                                "to_char(tc_srg019,'mm-dd-yyyy') tc_srg019, tc_srg020, tc_srg021, to_char(tc_srg022,'mm-dd-yyyy') tc_srg022, tc_srg023, tc_srg024, " .
                                                "to_char(tc_srg025,'mm-dd-yyyy') tc_srg025, tc_srg026, tc_srg027, tc_srg028, tc_srg029, ecb06, ecb17 " . 
                                                "from tc_srg_file, ecb_file " . 
                                                "where tc_srg001='$sfb01' and tc_srg004<'$ecb03' and tc_srg003=ecb01 and tc_srg004=ecb03 order by ecb06 desc "; 
                                        $erp_sqltcsrg3 = oci_parse($erp_conn1,$stcsrg3 );
                                        oci_execute($erp_sqltcsrg3);  
                                        $rowtcsrg3 = oci_fetch_array($erp_sqltcsrg3, OCI_ASSOC);
                                        $ecb063=$rowtcsrg3['ECB06'];
                                        $ecb173=$rowtcsrg3['ECB17'];                 
                                        if ($rowtcsrg3['TC_SRG006']=='Y') {  //本站要QC 讀PQC出站時間
                                            $outdate3=$rowtcsrg3['TC_SRG013'];
                                            $outtime3=$rowtcsrg3['TC_SRG014'];
                                            $gem02=getgem02($rowtcsrg['TC_SRG012'],$erp_conn1);  
                                            $msg=" 於 $outdate3 $outtime3 在 $gem02 $ecb063 $ecb173 QC出站, 等待$ecb06 $ecb17 入站e." ;  
                                        } else {                                //無PQC 讀正常出站時間
                                            $outdate3=$rowtcsrg3['TC_SRG010'];
                                            $outtime3=$rowtcsrg3['TC_SRG011'];
                                            $gem02=getgem02($rowtcsrg['TC_SRG012'],$erp_conn1);  
                                            $msg=" 於 $outdate3 $outtime3 在 $gem02 $ecb063 $ecb173 出站, 等待$ecb06 $ecb17 入站f." ;  
                                        }  
                                    }
                                }    
                            }                           
                        }
                    }
                } 
            
           //
           }        
        }
        //
    }   
    return ($msg);
} 

?>
