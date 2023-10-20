<?
          session_start(); 
          
          date_default_timezone_set('Asia/Shanghai')   ;
          
          $erp_db_host1 = "topprod";
          $erp_db_user1 = "vd110";
          $erp_db_pass1 = "vd110";  
          $erp_conn1 = oci_connect($erp_db_user1, $erp_db_pass1, $erp_db_host1,'AL32UTF8'); 
           
          $erp_db_host2 = "topprod";
          $erp_db_user2 = "vd210";
          $erp_db_pass2 = "vd210";  
          $erp_conn2 = oci_connect($erp_db_user2, $erp_db_pass2, $erp_db_host2,'AL32UTF8');  
          
          if (is_null($_POST['edate'])) {
              $edate = date('Y-m-d');
          } else {
              $edate=$_POST['edate'];
          }  
          
          //取出15天內的出貨CASES數
          $bdate=date('Y-m-d', strtotime('-15 days', strtotime($edate)));
          
          if (is_null($_POST['occ01'])) {
              $occ01 = 'All';        
          } else {                                                                                
              $occ01=$_POST['occ01'];                                                             
          }  
          
          if ($occ01=='All') {  
              $inoccfilter='';
              $outoccfilter='';  
              $tcofa04filter='';   
          } else {                   
              $inoccfilter=" and oea17='$occ01' ";
              $outoccfilter=" and oea04 in (select occ01 from occ_file where occ07='$occ01') ";   
              $tcofa04filter=" and tc_ofa04 in (select occ01 from occ_file where occ07='$occ01') ";  
          }    
          $month=date('Y-m', strtotime('-0 months', strtotime($edate))); 
          $year=date('Y', strtotime('-0 months', strtotime($edate))); 
          
          $mkdate='0,0,0,'.substr($edate,5,2) .','.substr($edate,8,2).','.substr($edate,0,4);                
          $pmonth=date('Y-m', strtotime('-1 months', mktime($mkdate)));      //要取出上個月的資料
          $ppmonth=date('Y-m', strtotime('-2 months', mktime($mkdate)));   //要取出上上個月的資料     
          
          
          //取出30天 每天的出貨組數
          //計算 $edate出貨幾組   
          $file = "boss1.txt"; 
          $handle = fopen($file, 'w');
          $sout= "select to_char(oga02,'yyyy/mm/dd') oga02, substr(to_char(oga02,'DAY'),3,1) oga021,  count(*) cases from oga_file, oea_file where oga02 between to_date('$bdate','yy/mm/dd') and to_date('$edate','yy/mm/dd') and oga16=oea01 $outoccfilter group by oga02 order by oga02 ";   
          $erp_sqlout = oci_parse($erp_conn1,$sout );
          oci_execute($erp_sqlout);  
          while ($rowout = oci_fetch_array($erp_sqlout, OCI_ASSOC)){   
              $oga02=$rowout['OGA02'];
              $oga021=$rowout['OGA021']; 
              $cases=$rowout['CASES'];
              $data = "$oga02($oga021),$cases\n"; 
              fwrite($handle, $data);
          }
          fclose($handle);
         
        ?>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>      
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        <title>客戶進出貨狀況查詢</title>
        <link rel="stylesheet" href="style.css" type="text/css">
        <script src="../amcharts/amcharts.js" type="text/javascript"></script> 
        <script type="text/javascript" language="javascript" src="scripts.js"></script>                      
        <script type="text/javascript" language="javascript" src="My97DatePicker/WdatePicker.js"></script>    
        <script type="text/javascript" language="javascript" src="calendarDateInput.js"></script>     
        <script type="text/javascript">
            var chart;
            var dataProvider;
            
            window.onload = function() {
              createChart();            
              loadCSV("../boss1.txt");                                    
            }
            function loadCSV(file) {
              if (window.XMLHttpRequest) {
                // IE7+, Firefox, Chrome, Opera, Safari
                var request = new XMLHttpRequest();
              } else {        
                // code for IE6, IE5
                var request = new ActiveXObject('Microsoft.XMLHTTP');
              }
              // load
              request.open('GET', file, false);
              request.send();
              parseCSV(request.responseText);
            }                                    
                                               
            
            // method which parses csv data

            function parseCSV(data){       
                //replace UNIX new lines   
                data = data.replace (/\r\n/g, "\n");  
                //replace MAC new lines  
                data = data.replace (/\r/g, "\n");  
                //split into rows     
                var rows = data.split("\n");            
                // create array which will hold our data:   
                dataProvider = [];           
                // loop through all rows  
                for (var i = 0; i < rows.length; i++){   
                    // this line helps to skip empty rows  
                    if (rows[i]) {                            
                        // our columns are separated by comma  
                        var column = rows[i].split(",");  
                        // column is array now     
                        // first item is date     
                        var date = column[0];                       
                        // second item is value of the second column  
                        var value1 = column[1];                        
                        // create object which contains all these items:  
                        var dataObject = {date:date, value1:value1};    
                        // add object to dataProvider array     
                        dataProvider.push(dataObject); 
                    }    
                }                                  
                // set data provider to the chart  
                chart.dataProvider = dataProvider;                
                // this will force chart to rebuild using new data   
                chart.validateData();  
            }              
            // method which creates chart 
            function createChart(){                       
                // chart variable is declared in the top    
                chart = new AmCharts.AmSerialChart();       
                // here we tell the chart name of category   
                // field in our data provider.            
                // we called it "date" (look at parseCSV method)   
                chart.categoryField = "date";     
                chart.startDuration = 1;   
                //chart.depth3D = 20;
                //chart.angle = 30;

 
                                             
                // AXES
                // category
                var categoryAxis = chart.categoryAxis;
                categoryAxis.labelRotation = 45;
                categoryAxis.gridPosition = "start";     
                categoryAxis.title="15天客戶出貨趨勢圖";
                
                  
                var graph = new AmCharts.AmGraph();        
                // graph should know at what field from data   
                // provider it should get values.   
                // let's assign value1 field for this graph  
                graph.valueField = "value1";  
                graph.type = "column";  
                graph.balloonText = "[[date]]: [[value1]]";
                graph.lineAlpha = 0;
                graph.fillAlphas = 0.8;
                // and add graph to the chart   
                chart.addGraph(graph);              
                // 'chartdiv' is id of a container     
                // where our chart will be      
                chart.write('chartdiv');   
            }
                                          
        </script>
    </head>
    
    <body>
      <tr>
        <form name="form1" id="form1" method="post" action="<?=$PHP_SELF;?>">
             <table width="100%"  border="0" cellspacing="4" cellpadding="0" STYLE="BORDER: 1px solid #666666">
               <tr>
                 <td>日期:  <input name="edate" type="text" id="edate" size="12" maxlength="12" onfocus="new WdatePicker()" value=<?=$edate;?> > &nbsp; 
                     客戶:  <select name="occ01" id="occ01">  
                            <option value="All">全部客戶</option>
                            <?
                              $s1= "select  occ01,  occ02 from occ_file where occ01 in (select distinct occ07 from occ_file)  and substr(occ01,1,1)!='V' order by occ02 ";
                              $erp_sql1 = oci_parse($erp_conn2,$s1 );
                              oci_execute($erp_sql1);  
                              while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                                  echo "<option value=" . $row1["OCC01"];  
                                  if ($occ01 == $row1["OCC01"]) echo " selected";                  
                                  echo ">"  .$row1["OCC02"] . "</option>"; 
                              }   
                            ?>
                            </select> &nbsp; 
                     <input style="font-size: 20px;" type="submit" name="Submit2" value="查詢"></td> &nbsp;           
               </tr>
             </table>
        </form>
      </tr>
      <?
        //計算 $edate到貨幾組        
        $sin= "select sum(cases) cases, sum(case1) case1, sum(case2) case2, sum(case3) case3, sum(case4) case4 from ".
               "(select cases, decode(ta_oea004, '1', cases, 0) case1, decode(ta_oea004, '2', cases, 0) case2, decode(ta_oea004, '3', cases, 0 ) case3, decode(ta_oea004, '4', cases, 0 ) case4  from " . 
                 "(select ta_oea004,  count(*) cases from oea_file where oeaconf='Y' $inoccfilter and oea02=to_date('$edate','yy/mm/dd') group by ta_oea004 ))";   
        $erp_sqlin = oci_parse($erp_conn2,$sin );
        oci_execute($erp_sqlin);  
        $rowin = oci_fetch_array($erp_sqlin, OCI_ASSOC);
        
        //計算 $edate出貨幾組  
        $sout= "select sum(cases) cases, sum(case1) case1, sum(case2) case2, sum(case3) case3, sum(case4) case4 from ".
               "(select cases, decode(ta_oea004, '1', cases, 0) case1, decode(ta_oea004, '2', cases, 0) case2, decode(ta_oea004, '3', cases, 0 ) case3, decode(ta_oea004, '4', cases, 0 ) case4  from " . 
                 "(select ta_oea004,  count(*) cases from oga_file, oea_file where oga02=to_date('$edate','yy/mm/dd') and oga16=oea01 $outoccfilter group by ta_oea004 ))";   
        $erp_sqlout = oci_parse($erp_conn1,$sout );
        oci_execute($erp_sqlout);  
        $rowout = oci_fetch_array($erp_sqlout, OCI_ASSOC);
        
                
        //本月出貨組
        $sout1="select count(*) cases from oga_file, oea_file where to_char(oga02,'yyyy-mm')='$month' and oga02<=to_date('$edate','yy/mm/dd') and oga16=oea01 $outoccfilter ";           
        $erp_sqlout1 = oci_parse($erp_conn1,$sout1);
        oci_execute($erp_sqlout1);  
        $rowout1 = oci_fetch_array($erp_sqlout1, OCI_ASSOC);
        
        //本月出貨顆
        $sout2="select sum(ogb12*imaud07) units from oga_file, ogb_file, oea_file, ima_file where to_char(oga02,'yyyy-mm')='$month' and oga02<=to_date('$edate','yy/mm/dd') and oga16=oea01 $outoccfilter and oga01=ogb01 and ogb04=ima01 ";           
        $erp_sqlout2 = oci_parse($erp_conn1,$sout2);
        oci_execute($erp_sqlout2);  
        $rowout2 = oci_fetch_array($erp_sqlout2, OCI_ASSOC);
        
        //本月出貨金額
        $sout3="select sum(tc_ofb14*tc_ofa24) costs from tc_ofb_file, tc_ofa_file where tc_ofb01=tc_ofa01 $tcofa04filter and to_char(tc_ofa02,'yyyy-mm')='$month' and tc_ofa02<=to_date('$edate','yy/mm/dd') ";
        $erp_sqlout3 = oci_parse($erp_conn2,$sout3);
        oci_execute($erp_sqlout3);  
        $rowout3 = oci_fetch_array($erp_sqlout3, OCI_ASSOC);
        
        
        //本年出貨組
        $sout4="select count(*) cases from oga_file, oea_file where to_char(oga02,'yyyy')='$year' and oga02<=to_date('$edate','yy/mm/dd') and oga16=oea01 $outoccfilter ";           
        $erp_sqlout4 = oci_parse($erp_conn1,$sout4);
        oci_execute($erp_sqlout4);  
        $rowout4 = oci_fetch_array($erp_sqlout4, OCI_ASSOC);
        
        //本月出貨顆
        $sout5="select sum(ogb12*imaud07) units from oga_file, ogb_file, oea_file, ima_file where to_char(oga02,'yyyy')='$year' and oga02<=to_date('$edate','yy/mm/dd') and oga16=oea01 $outoccfilter and oga01=ogb01 and ogb04=ima01 ";           
        $erp_sqlout5 = oci_parse($erp_conn1,$sout5);
        oci_execute($erp_sqlout5);  
        $rowout5 = oci_fetch_array($erp_sqlout5, OCI_ASSOC);
        
        //本月出貨金額
        $sout6="select sum(tc_ofb14*tc_ofa24) costs from tc_ofb_file, tc_ofa_file where tc_ofb01=tc_ofa01 $tcofa04filter and to_char(tc_ofa02,'yyyy')='$year' and tc_ofa02<=to_date('$edate','yy/mm/dd') ";
        $erp_sqlout6 = oci_parse($erp_conn2,$sout6);
        oci_execute($erp_sqlout6);  
        $rowout6 = oci_fetch_array($erp_sqlout6, OCI_ASSOC);
        
        
        //上月出貨組
        $sout7="select count(*) cases from oga_file, oea_file where to_char(oga02,'yyyy-mm')='$pmonth' and oga16=oea01 $outoccfilter ";           
        $erp_sqlout7 = oci_parse($erp_conn1,$sout7);
        oci_execute($erp_sqlout7);  
        $rowout7 = oci_fetch_array($erp_sqlout7, OCI_ASSOC);
        
        //上月出貨顆
        $sout8="select sum(ogb12*imaud07) units from oga_file, ogb_file, oea_file, ima_file where to_char(oga02,'yyyy-mm')='$pmonth' and oga16=oea01 $outoccfilter and oga01=ogb01 and ogb04=ima01 ";           
        $erp_sqlout8 = oci_parse($erp_conn1,$sout8);
        oci_execute($erp_sqlout8);  
        $rowout8 = oci_fetch_array($erp_sqlout8, OCI_ASSOC);
        
        //上月出貨金額
        $sout9="select sum(tc_ofb14*tc_ofa24) costs from tc_ofb_file, tc_ofa_file where tc_ofb01=tc_ofa01 $tcofa04filter and to_char(tc_ofa02,'yyyy-mm')='$pmonth' ";
        $erp_sqlout9 = oci_parse($erp_conn2,$sout9);
        oci_execute($erp_sqlout9);  
        $rowout9 = oci_fetch_array($erp_sqlout9, OCI_ASSOC);
        
        //上上月出貨組
        $souta="select count(*) cases from oga_file, oea_file where to_char(oga02,'yyyy-mm')='$ppmonth' and oga16=oea01 $outoccfilter ";           
        $erp_sqlouta = oci_parse($erp_conn1,$souta);
        oci_execute($erp_sqlouta);  
        $rowouta = oci_fetch_array($erp_sqlouta, OCI_ASSOC);
        
        //上上月出貨顆
        $soutb="select sum(ogb12*imaud07) units from oga_file, ogb_file, oea_file, ima_file where to_char(oga02,'yyyy-mm')='$ppmonth' and oga16=oea01 $outoccfilter and oga01=ogb01 and ogb04=ima01 ";           
        $erp_sqloutb = oci_parse($erp_conn1,$soutb);
        oci_execute($erp_sqloutb);  
        $rowoutb = oci_fetch_array($erp_sqloutb, OCI_ASSOC);
        
        //上上月出貨金額
        $soutc="select sum(tc_ofb14*tc_ofa24) costs from tc_ofb_file, tc_ofa_file where tc_ofb01=tc_ofa01 $tcofa04filter and to_char(tc_ofa02,'yyyy-mm')='$ppmonth' ";
        $erp_sqloutc = oci_parse($erp_conn2,$soutc);
        oci_execute($erp_sqloutc);  
        $rowoutc = oci_fetch_array($erp_sqloutc, OCI_ASSOC);
        
        
        $mcases= number_format($rowout1["CASES"], 1, ".", ",");
        $munits= number_format($rowout2["UNITS"], 1, ".", ",");  
        $mcosts= number_format($rowout3["COSTS"], 1, ".", ",");  
        $ycases= number_format($rowout4["CASES"], 1, ".", ",");
        $yunits= number_format($rowout5["UNITS"], 1, ".", ",");  
        $ycosts= number_format($rowout6["COSTS"], 1, ".", ","); 
        $pmcases= number_format($rowout7["CASES"], 1, ".", ",");
        $pmunits= number_format($rowout8["UNITS"], 1, ".", ",");  
        $pmcosts= number_format($rowout9["COSTS"], 1, ".", ","); 
        $ppmcases= number_format($rowouta["CASES"], 1, ".", ",");
        $ppmunits= number_format($rowoutb["UNITS"], 1, ".", ",");  
        $ppmcosts= number_format($rowoutc["COSTS"], 1, ".", ","); 
      
      ?>
      <tr>
        <table width="100%"  border="0" cellspacing="4" cellpadding="0" STYLE="BORDER: 1px solid #666666">  
          <tr>
            <td>
              <table width="100%"  border="0" cellspacing="4" cellpadding="0" STYLE="BORDER: 1px solid #666666">  
              <tr><td colspan="4"><B><font color="GREEN" ><?=$edate;?></font></B> 到貨總數: <B><font color="RED"><?=$rowin['CASES'];?></font></B> 組</td></tr>
              <tr><td>正常: <B><font color="BLUE" ><?=$rowin['CASE1'];?></font></B> 組</td><td>重做: <B><font color="BLUE"><?=$rowin['CASE2'];?></font></B> 組</td><td>修改: <B><font color="BLUE"><?=$rowin['CASE3'];?></font></B> 組</td><td>內返: <B><font color="BLUE"><?=$rowin['CASE4'];?></font></B> 組</td></tr>
              </table>
            </td>   
            <td>
              <table width="100%"  border="0" cellspacing="4" cellpadding="0" STYLE="BORDER: 1px solid #666666"> 
              <tr><td colspan="4"><B><font color="GREEN" ><?=$edate;?></font></B> 出貨總數: <B><font color="RED"><?=$rowout['CASES'];?></font></B> 組</td></tr>
              <tr><td>正常: <B><font color="BLUE"><?=$rowout['CASE1'];?></font></B> 組</td><td>重做: <B><font color="BLUE"><?=$rowout['CASE2'];?></font></B> 組</td><td>修改: <B><font color="BLUE"><?=$rowout['CASE3'];?></font></B> 組</td><td>內返: <B><font color="BLUE"><?=$rowout['CASE4'];?></font></B> 組</td></tr> 
              </table>
            </td>      
          </tr>    
        </table>      
      </tr>
      <tr>
          <table width="100%"  border="0" cellspacing="4" cellpadding="0" STYLE="BORDER: 1px solid #666666">  
            <tr>
              <td width="20%" valign="top">      
                <table width="100%"  border="0" cellspacing="4" cellpadding="0" STYLE="BORDER: 1px solid #666666">  
                  <tr><td><b><?=$ppmonth;?></b></td> <td style="text-align:right; color:red; font-weight: bold; "><?=$ppmcases;?></td> <td></td> </tr>                                         
                  <tr><td><b>出貨 </b></td> <td style="text-align:right; color:red; font-weight: bold; "><?=$ppmunits;?></td> <td></td> </tr> 
                  <tr><td><b>累計 </b></td> <td style="text-align:right; color:red; font-weight: bold; "><?=$ppmcosts;?></td> <td></td> </tr> 
                  <tr><td colspan="3">&nbsp;<td></tr>
                  <tr><td><b><?=$pmonth;?></b></td> <td style="text-align:right; color:red; font-weight: bold; "><?=$pmcases;?></td> <td></td> </tr>                                         
                  <tr><td><b>出貨 </b></td> <td style="text-align:right; color:red; font-weight: bold; "><?=$pmunits;?></td> <td></td> </tr> 
                  <tr><td><b>累計 </b></td> <td style="text-align:right; color:red; font-weight: bold; "><?=$pmcosts;?></td> <td></td> </tr> 
                  <tr><td colspan="3">&nbsp;<td></tr>
                  <tr><td><b><?=$month;?></b></td> <td style="text-align:right; color:red; font-weight: bold; "><?=$mcases;?></td> <td></td> </tr>                                         
                  <tr><td><b>出貨 </b></td> <td style="text-align:right; color:red; font-weight: bold; "><?=$munits;?></td> <td></td> </tr> 
                  <tr><td><b>累計 </b></td> <td style="text-align:right; color:red; font-weight: bold; "><?=$mcosts;?></td> <td></td> </tr> 
                  <tr><td colspan="3">&nbsp;<td></tr>
                  <tr><td><b><?=$year;?> </b></td> <td style="text-align:right; color:red; font-weight: bold; "><?=$ycases;?></td> <td></td> </tr> 
                  <tr><td><b>出貨 </b></td> <td style="text-align:right; color:red; font-weight: bold; "><?=$yunits;?></td> <td></td> </tr> 
                  <tr><td><b>累計 </b></td> <td style="text-align:right; color:red; font-weight: bold; "><?=$ycosts;?></td> <td></td> </tr> 
                  <tr><td colspan="3" style="text-align:center;"><a href="boss2.php"><h1>製處分析</h1></a></td></tr>   
                </table>
              </td>
              <td width="80%" valign="top">                                                                           
                    <div id="chartdiv" style="width: 100%; height: 450px;"></div>  
              </td>
            </tr>
          </table>   
      </tr> 
    </body>       
</html>