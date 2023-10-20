<?
          session_start(); 
          date_default_timezone_set('Asia/Shanghai')   ; 
          
          $db_host = "localhost:3302";
          $db_user = "oosuser";
          $db_pass = "veden*()";
          $db_name = "erpdb";   
          // make the connection
          $conn = mysql_connect($db_host,$db_user,$db_pass) or die('Oracle 11G server connect error!!');
          mysql_query("SET NAMES 'utf8'");
          mysql_select_db($db_name,$conn) or die('Database connect error!!');
          
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
          
          
          if (is_null($_POST['mid'])) {
              $mid = 'All';        
          } else {                                                                                
              $mid=$_POST['mid'];                                                             
          }  
          
          if ($mid=='All') {  
              $midfilter='';    
          } else {                   
              $midfilter=" where s13 in $mid ";  
              $inmidfilter=" and oebud04 in $mid ";   
          }    
          
          $month=date('Y-m', strtotime('-0 months', strtotime($edate))); 
          $year=date('Y', strtotime('-0 months', strtotime($edate))); 
          $mkdate='0,0,0,'.substr($edate,5,2) .','.substr($edate,8,2).','.substr($edate,0,4);                
          $pmonth=date('Y-m', strtotime('-1 months', mktime($mkdate)));      //要取出上個月的資料
          $ppmonth=date('Y-m', strtotime('-2 months', mktime($mkdate)));   //要取出上上個月的資料
          
          
          
          //本月各製處有效/delay/內返/重修 
          $file = "boss21.txt"; 
          $handle = fopen($file, 'w');
          
          $query2= " (select s13 , 0 sent, sum(s25*s26) tsent, 0 reject, 0 redo, 0 delay  from casesent where (date_format(s01,'%Y-%m')='$month') group by s13) ";  
          $query3= " (select rid s13, 0 sent, 0 tsent, sum(rqty1) reject, sum(rqty2) redo, 0 delay from casereject where (date_format(rdate,'%Y-%m')='$month') group by rid) ";   
          $query4= " (select d13 s13 , 0 sent, 0 tsent, 0 reject, 0 redo, sum(d25*d26) delay  from casedelay where (date_format(d21,'%Y-%m')='$month') and d16='' group by d13) ";    
          $query = "select sum(sent) sent, sum(tsent) tsent, sum(reject) reject, sum(redo) redo, sum(delay) delay from " . 
                   "(select s13 , sent, tsent, reject, redo, delay from " .
                   "( $query2 union all $query3 union $query4 ) a $midfilter ) b where (sent+tsent+reject+redo+delay)!=0  ";
          $result= mysql_query($query);
          $row = mysql_fetch_array($result);  
          $tsent  =$row['tsent'];
          $reject =$row['reject']; 
          $redo   =$row['redo']; 
          $delay  =$row['delay']; 
          $oksent = number_format(($row['tsent']-$row['delay']-$row['reject']-$row['redo']),1,'.','');
          $data = "返工,$reject\n"; 
          fwrite($handle, $data);  
          $data = "Redo,$redo\n"; 
          fwrite($handle, $data);  
          $data = "Delay,$delay\n";
          fwrite($handle, $data);   
          $data = "有效顆數,$oksent\n"; 
          fwrite($handle, $data);
          fclose($handle);  
          
          //本年各製處有效/delay/內返/重修 
          $file = "boss22.txt"; 
          $handle = fopen($file, 'w');
          
          $query2= " (select s13 , 0 sent, sum(s25*s26) tsent, 0 reject, 0 redo, 0 delay  from casesent where (date_format(s01,'%Y')='$year') group by s13) ";  
          $query3= " (select rid s13, 0 sent, 0 tsent, sum(rqty1) reject, sum(rqty2) redo, 0 delay from casereject where (date_format(rdate,'%Y')='$year') group by rid) ";   
          $query4= " (select d13 s13 , 0 sent, 0 tsent, 0 reject, 0 redo, sum(d25*d26) delay  from casedelay where (date_format(d21,'%Y')='$year') and d16='' group by d13) ";    
          $query = "select sum(sent) sent, sum(tsent) tsent, sum(reject) reject, sum(redo) redo, sum(delay) delay from " . 
                   "(select s13 , sent, tsent, reject, redo, delay from " .
                   "( $query2 union all $query3 union $query4 ) a $midfilter ) b where (sent+tsent+reject+redo+delay)!=0  ";
          $result= mysql_query($query);
          $row = mysql_fetch_array($result);  
          $tsent  =$row['tsent'];
          $reject =$row['reject']; 
          $redo   =$row['redo']; 
          $delay  =$row['delay']; 
          $oksent =number_format(($row['tsent']-$row['delay']-$row['reject']-$row['redo']),1,'.','');
          $data = "返工,$reject\n"; 
          fwrite($handle, $data);  
          $data = "Redo,$redo\n"; 
          fwrite($handle, $data);  
          $data = "Delay,$delay\n";
          fwrite($handle, $data);   
          $data = "有效顆數,$oksent\n"; 
          fwrite($handle, $data);
          fclose($handle);  
          
         
        ?>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>      
    <head>        
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        <title>製處進出貨狀況查詢</title>
        <link rel="stylesheet" href="style.css" type="text/css">
        <script src="../amcharts/amcharts.js" type="text/javascript"></script> 
        <script type="text/javascript" language="javascript" src="scripts.js"></script>                      
        <script type="text/javascript" language="javascript" src="My97DatePicker/WdatePicker.js"></script>    
        <script type="text/javascript" language="javascript" src="calendarDateInput.js"></script>     
        <script type="text/javascript">
            var chart1;
            var chart2; 
            var dataProvider1;       
            var dataProvider2;
            
            var chartData1 = [{      
            }];
            var chartData2 = [{        
            }];
                     
            window.onload = function() {
              createChart1();            
              loadCSV1("../boss21.txt");      
              createChart2();            
              loadCSV2("../boss22.txt");                               
            }
            
            function loadCSV1(file) {
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
              parseCSV1(request.responseText);
            }                                    
                                               
            
            // method which parses csv data

            function parseCSV1(data){       
                //replace UNIX new lines   
                data = data.replace (/\r\n/g, "\n");  
                //replace MAC new lines  
                data = data.replace (/\r/g, "\n");  
                //split into rows     
                var rows = data.split("\n");            
                // create array which will hold our data:   
                dataProvider1 = [];           
                // loop through all rows  
                for (var i = 0; i < rows.length; i++){   
                    // this line helps to skip empty rows  
                    if (rows[i]) {                            
                        // our columns are separated by comma  
                        var column = rows[i].split(",");  
                        // column is array now     
                        // first item is date     
                        var date = column[0];                        
                        var value1 = column[1];  
                        // create object which contains all these items:  
                        var dataObject = {date:date, value1:value1};    
                        // add object to dataProvider array     
                        dataProvider1.push(dataObject);  
                    }    
                }                      
                // set data provider to the chart  
                chart1.dataProvider = dataProvider1;                
                // this will force chart to rebuild using new data    
                chart1.validateData();       
            }      
                                                  
            // method which creates chart 
            function createChart1(){                
                      
                //chart variable is declared in the top   
                chart1 = new AmCharts.AmPieChart();
                chart1.dataProvider = chartData1;  
                chart1.addTitle("本月累計出貨", 16);                 
                chart1.titleField = "date";
                chart1.valueField = "value1";
                chart1.sequencedAnimation = true;
                chart1.startEffect = "elastic";
                chart1.innerRadius = "30%";
                chart1.startDuration = 2;
                chart1.labelRadius = 5;                
                chart1.depth3D = 10;
                chart1.angle = 15;             
                
                // LEGEND
                legend = new AmCharts.AmLegend();
                legend.align = "center";
                legend.markerType = "circle";
                chart1.addLegend(legend);
                                
                chart1.write('chartdiv1');
            }
            
            function loadCSV2(file) {
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
              parseCSV2(request.responseText);
            }                                    
                                               
            
            // method which parses csv data

            function parseCSV2(data){       
                //replace UNIX new lines   
                data = data.replace (/\r\n/g, "\n");  
                //replace MAC new lines  
                data = data.replace (/\r/g, "\n");  
                //split into rows     
                var rows = data.split("\n");            
                // create array which will hold our data:   
                dataProvider2 = [];           
                // loop through all rows  
                for (var i = 0; i < rows.length; i++){   
                    // this line helps to skip empty rows  
                    if (rows[i]) {                            
                        // our columns are separated by comma  
                        var column = rows[i].split(",");  
                        // column is array now     
                        // first item is date     
                        var date = column[0];                        
                        var value1 = column[1];  
                        // create object which contains all these items:  
                        var dataObject = {date:date, value1:value1};    
                        // add object to dataProvider array     
                        dataProvider2.push(dataObject);  
                    }    
                }                      
                // set data provider to the chart  
                chart2.dataProvider = dataProvider2;                
                // this will force chart to rebuild using new data    
                chart2.validateData();       
            }   
            
            function createChart2(){                
                      
                //chart variable is declared in the top   
                chart2 = new AmCharts.AmPieChart();
                chart2.dataProvider = chartData2;  
                chart2.addTitle("本年累計出貨", 16);                 
                chart2.titleField = "date";
                chart2.valueField = "value1";
                chart2.sequencedAnimation = true;
                chart2.startEffect = "elastic";
                chart2.innerRadius = "30%";
                chart2.startDuration = 2;
                chart2.labelRadius = 5;                
                chart2.depth3D = 10;
                chart2.angle = 15;             
                
                // LEGEND
                legend = new AmCharts.AmLegend();
                legend.align = "center";
                legend.markerType = "circle";
                chart2.addLegend(legend);
                                
                chart2.write('chartdiv2');
            }
                                          
        </script>
    </head>
    
    <body>
      <tr>
        <form name="form1" id="form1" method="post" action="<?=$PHP_SELF;?>">
             <table width="100%"  border="0" cellspacing="4" cellpadding="0" STYLE="BORDER: 1px solid #666666">
               <tr>
                 <td>日期:  <input name="edate" type="text" id="edate" size="12" maxlength="12" onfocus="new WdatePicker()" value=<?=$edate;?> > &nbsp; 
                     客戶:  <select name="mid" id="mid">  
                                <option value="All">全部製處</option>   
                                <option value="('690000')" <? if ($mid == "('690000')") echo " selected";?> > 研發部</option>
                                <option value="('6A1000')" <? if ($mid == "('6A1000')") echo " selected";?>>製一處</option>
                                <option value="('6A2000')" <? if ($mid == "('6A2000')") echo " selected";?>>製二處</option>
                                <option value="('6A3000')" <? if ($mid == "('6A3000')") echo " selected";?>>製三處</option>
                                <option value="('6A4100','6A4300')" <? if ($mid == "('6A4100','6A4300')") echo " selected";?>>製四金</option>
                                <option value="('6A4200','6A4300')" <? if ($mid == "('6A4200','6A4300')") echo " selected";?>>製四活</option>
                                <option value="('6A5000')" <? if ($mid == "('6A5000')") echo " selected";?>>製五處</option> 
                            </select> &nbsp; 
                     <input style="font-size: 20px;" type="submit" name="Submit2" value="查詢"></td> &nbsp;           
               </tr>
             </table>
        </form>
      </tr>
      <?
        //計算 $edate到貨幾顆        
        $sin= "select sum(cases) cases, sum(case1) case1, sum(case2) case2, sum(case3) case3, sum(case4) case4 from ".
               "(select cases, decode(ta_oea004, '1', cases, 0) case1, decode(ta_oea004, '2', cases, 0) case2, decode(ta_oea004, '3', cases, 0) case3, decode(ta_oea004, '4', cases, 0) case4  from " . 
                 "(select ta_oea004, sum(oeb12*(decode(imaud07,2,1,imaud07))) cases from oea_file, oeb_file, ima_file where oeaconf='Y' and oea01=oeb01 and oeb04=ima01 $inmidfilter and oea02=to_date('$edate','yy/mm/dd') group by ta_oea004 ))";   
        $erp_sqlin = oci_parse($erp_conn2,$sin );
        oci_execute($erp_sqlin);  
        $rowin = oci_fetch_array($erp_sqlin, OCI_ASSOC);
        
        //計算 $edate出貨幾顆 
        $sout= "select sum(cases) cases, sum(case1) case1, sum(case2) case2, sum(case3) case3, sum(case4) case4 from ".
               "(select cases, decode(ta_oea004, '1', cases, 0) case1, decode(ta_oea004, '2', cases, 0) case2, decode(ta_oea004, '3', cases, 0 ) case3, decode(ta_oea004, '4', cases, 0 ) case4  from " . 
                 "(select ta_oea004, sum(oeb12*(decode(imaud07,2,1,imaud07))) cases from oga_file, oea_file, oeb_file, ima_file where oga02=to_date('$edate','yy/mm/dd') and oga16=oea01 and oea01=oeb01 and oeb04=ima01 $inmidfilter group by ta_oea004 ))";   
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
        $sout3="select sum(tc_ofb14*tc_ofa24) costs from tc_ofb_file, tc_ofa_file where tc_ofb01=tc_ofa01 and to_char(tc_ofa02,'yyyy-mm')='$month' and tc_ofa02<=to_date('$edate','yy/mm/dd') ";
        $erp_sqlout3 = oci_parse($erp_conn2,$sout3);
        oci_execute($erp_sqlout3);  
        $rowout3 = oci_fetch_array($erp_sqlout3, OCI_ASSOC);
        
        
        //本年出貨組
        $sout4="select count(*) cases from oga_file, oea_file where to_char(oga02,'yyyy')='$year' and oga02<=to_date('$edate','yy/mm/dd') and oga16=oea01 $outoccfilter ";           
        $erp_sqlout4 = oci_parse($erp_conn1,$sout4);
        oci_execute($erp_sqlout4);  
        $rowout4 = oci_fetch_array($erp_sqlout4, OCI_ASSOC);
        
        //本年出貨顆
        $sout5="select sum(ogb12*imaud07) units from oga_file, ogb_file, oea_file, ima_file where to_char(oga02,'yyyy')='$year' and oga02<=to_date('$edate','yy/mm/dd') and oga16=oea01 $outoccfilter and oga01=ogb01 and ogb04=ima01 ";           
        $erp_sqlout5 = oci_parse($erp_conn1,$sout5);
        oci_execute($erp_sqlout5);  
        $rowout5 = oci_fetch_array($erp_sqlout5, OCI_ASSOC);
        
        //本年出貨金額
        $sout6="select sum(tc_ofb14*tc_ofa24) costs from tc_ofb_file, tc_ofa_file where tc_ofb01=tc_ofa01 and to_char(tc_ofa02,'yyyy')='$year' and tc_ofa02<=to_date('$edate','yy/mm/dd') ";
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
        $sout9="select sum(tc_ofb14*tc_ofa24) costs from tc_ofb_file, tc_ofa_file where tc_ofb01=tc_ofa01 and to_char(tc_ofa02,'yyyy-mm')='$pmonth' ";
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
        $soutc="select sum(tc_ofb14*tc_ofa24) costs from tc_ofb_file, tc_ofa_file where tc_ofb01=tc_ofa01 and to_char(tc_ofa02,'yyyy-mm')='$ppmonth' ";
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
              <tr><td colspan="4"><B><font color="GREEN" ><?=$edate;?></font></B> 到貨總數: <B><font color="RED"><?=$rowin['CASES'];?></font></B> 顆</td></tr>
              <tr><td>正常: <B><font color="BLUE" ><?=$rowin['CASE1'];?></font></B> 顆</td><td>重做: <B><font color="BLUE"><?=$rowin['CASE2'];?></font></B> 顆</td><td>修改: <B><font color="BLUE"><?=$rowin['CASE3'];?></font></B> 顆</td><td>內返: <B><font color="BLUE"><?=$rowin['CASE4'];?></font></B> 顆</td></tr>
              </table>
            </td>   
            <td>
              <table width="100%"  border="0" cellspacing="4" cellpadding="0" STYLE="BORDER: 1px solid #666666"> 
              <tr><td colspan="4"><B><font color="GREEN" ><?=$edate;?></font></B> 出貨總數: <B><font color="RED"><?=$rowout['CASES'];?></font></B> 組</td></tr>
              <tr><td>正常: <B><font color="BLUE"><?=$rowout['CASE1'];?></font></B> 顆</td><td>重做: <B><font color="BLUE"><?=$rowout['CASE2'];?></font></B> 顆</td><td>修改: <B><font color="BLUE"><?=$rowout['CASE3'];?></font></B> 顆</td><td>內返: <B><font color="BLUE"><?=$rowout['CASE4'];?></font></B> 顆</td></tr> 
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
                  <tr><td><b><?=$month;?></td> <td style="text-align:right; color:red; font-weight: bold; "><?=$mcases;?></td> <td></td> </tr>                                         
                  <tr><td><b>出貨 </b></td> <td style="text-align:right; color:red; font-weight: bold; "><?=$munits;?></td> <td></td> </tr> 
                  <tr><td><b>累計 </b></td> <td style="text-align:right; color:red; font-weight: bold; "><?=$mcosts;?></td> <td></td> </tr> 
                  <tr><td colspan="3">&nbsp;<td></tr>
                  <tr><td><b><?=$year;?></b></td> <td style="text-align:right; color:red; font-weight: bold; "><?=$ycases;?></td> <td></td> </tr> 
                  <tr><td><b>出貨 </b></td> <td style="text-align:right; color:red; font-weight: bold; "><?=$yunits;?></td> <td></td> </tr> 
                  <tr><td><b>累計 </b></td> <td style="text-align:right; color:red; font-weight: bold; "><?=$ycosts;?></td> <td></td> </tr>   
                  <tr><td colspan="3" style="text-align:center;"><a href="boss1.php"><h1>客戶分析</h1></a></td></tr>
                </table>
              </td>
              <td width="40%" valign="top">                                                                            
                    <div id="chartdiv1" style="width: 100%; height: 450px;"></div>        
              </td>
              <td width="40%" valign="top">                                                                           
                    <div id="chartdiv2" style="width: 100%; height: 450px;"></div>   
              </td>
            </tr>
          </table>   
      </tr> 
    </body>       
</html>