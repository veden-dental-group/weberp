<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>      
    <head>
        <?
          session_start(); 
          
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
          
          //取出31天內的出貨CASES數
          $bdate=date('Y-m-d', strtotime('-15 days', strtotime($edate)));
          
          if (is_null($_POST['occ01'])) {
              $occ01 = 'All';        
          } else {                                                                                
              $occ01=$_POST['occ01'];                                                             
          }  
          
          if ($occ01=='All') {  
              $inoccfilter='';
              $outoccfilter='';     
          } else {                   
              $inoccfilter=" and oea17='$occ01' ";
              $outoccfilter=" and oea04 in (select occ01 from occ_file where occ07='$occ01') ";   
          }    
          
          //取出30天 每天的出貨組數
          //計算 $edate出貨幾組   
          $File = "boss.txt"; 
          $Handle = fopen($File, 'w');
          $sout= "select to_char(oga02,'yyyy/mm/dd') oga02, substr(to_char(oga02,'DAY'),3,1) oga021,  count(*) cases from oga_file, oea_file where oga02 between to_date('$bdate','yy/mm/dd') and  to_date('$edate','yy/mm/dd') and oga16=oea01 $outoccfilter group by oga02 order by oga02 ";   
          $erp_sqlout = oci_parse($erp_conn1,$sout );
          oci_execute($erp_sqlout);  
          while ($rowout = oci_fetch_array($erp_sqlout, OCI_ASSOC)){   
              $oga02=$rowout['OGA02'];
              $oga021=$rowout['OGA021']; 
              $cases=$rowout['CASES'];
              $Data = "$oga02($oga021),$cases\n"; 
              fwrite($Handle, $Data);
          }
         
        ?>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        <title>進出貨狀況查詢</title>
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
              loadCSV("../boss.txt");                                    
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
                // chart must have a graph     
                
                // AXES
                // category
                var categoryAxis = chart.categoryAxis;
                categoryAxis.labelRotation = 45;
                categoryAxis.gridPosition = "start";     
                categoryAxis.title="15天客戶到貨趨勢圖";
                  
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
                  <tr><td><B>本月合計: <B><font color="RED"><?=$rowin['CASES'];?></font></B> 組</td></tr>                                         
                  <tr><td><B>本月合計: <B><font color="RED"><?=$rowin['CASES'];?></font></B> 顆</td></tr>
                  <tr><td><B>本月合計: <B><font color="RED"><?=$rowin['CASES'];?></font></B> 元</td></tr>
                  <tr><td><B>本年合計: <B><font color="RED"><?=$rowin['CASES'];?></font></B> 組</td></tr>
                  <tr><td><B>本年合計: <B><font color="RED"><?=$rowin['CASES'];?></font></B> 顆</td></tr>
                  <tr><td><B>本年合計: <B><font color="RED"><?=$rowin['CASES'];?></font></B> 元</td></tr>
                </table>
              </td>
              <td width="80%" valign="top"> 
                  <table width="100%"  border="0" cellspacing="4" cellpadding="0" STYLE="BORDER: 1px solid #666666">   
                    <div id="chartdiv" style="width: 100%; height: 300px;"></div>   
                  </table>             
              </td>
            </tr>
          </table>   
      </tr> 
    </body>       
</html>