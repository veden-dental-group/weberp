<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="content-type" content="text/html; charset=UTF-8">
  <title>Pie Chart</title>  
  <!-- prerequisites -->
  <link rel="stylesheet" href="amchartscss.css" type="text/css">
  <script src="amcharts/amcharts.js" type="text/javascript"></script>
  <script src="amcharts/serial.js" type="text/javascript"></script>
  <script src="amcharts/pie.js" type="text/javascript"></script>
  <script src="amcharts/exporting/amexport.js" type="text/javascript"></script>
  <script src="amcharts/exporting/rgbcolor.js" type="text/javascript"></script>
  <script src="amcharts/exporting/canvg.js" type="text/javascript"></script>
  <script src="amcharts/exporting/filesaver.js" type="text/javascript"></script>
  <script src="amcharts/jquery.min.js"></script>                             
  
  <!-- cutom functions -->
  <script>
      AmCharts.loadJSON = function(url) {
        // create the request
        if (window.XMLHttpRequest) {
          // IE7+, Firefox, Chrome, Opera, Safari
          var request = new XMLHttpRequest();
        } else {
          // code for IE6, IE5
          var request = new ActiveXObject('Microsoft.XMLHTTP');
        }

        // load it
        // the last "false" parameter ensures that our code will wait before the
        // data is loaded
        request.open('GET', url, false);
        request.send();

        // parse adn return the output
        return eval(request.responseText);
      };
      
      
    function sleep(ms){
      var dt = new Date();
      dt.setTime(dt.getTime() + ms);
      while (new Date().getTime() < dt.getTime());
    }
  </script>


  <script>


  // create chart
  AmCharts.ready(function() {  
    // load the data
    var chartData = AmCharts.loadJSON('erp_autoemail_clientdailycasereportv7-21_out.php');    
    var chart;  
    // PIE CHART
    chart = new AmCharts.AmPieChart();
    chart.dataProvider = chartData;
    chart.titleField = "customer";
    chart.valueField = "value";
    chart.outlineColor = "#FFFFFF";
    chart.outlineAlpha = 0.8;
    chart.marginTop = 0;  
    chart.marginRight=0;
    chart.minRadius=100;
    //chart.marginLeft=300;
    //char.marginBottom=300;
    chart.outlineThickness = 2;
   
    chart.balloonText = "[[title]]<br><span style='font-size:14px'><b>[[value]]</b> ([[percents]]%)</span>";
    // this makes the chart 3D
    chart.depth3D = 15;
    chart.angle = 30;     
    chart.colors=["#FF0F00", "#FF6600", "#FF9E01", "#FCD202", "#F8FF01", "#B0DE09", "#04D215", "#0D8ECF", "#0D52D1", "#2A0CD0", "#8A0CCF"]
    chart.labelText= "[[title]]: [[value]] ([[percents]]%)";
   //chart.labelRadiusField="radius"; 
    chart.labelRadius=5;      
   // chart.pieAlpha=2;
    chart.pieX="50%";
    chart.pieY="50%";
    chart.exportConfig = {
      menuItems:[]
    };
                
    chart.addListener('dataUpdated', function (event) {   
              chart.AmExport.output({format:'jpg', output: 'datastring'},
                 function(data) {
                    $.post("erp_autoemail_clientdailycasereportv7-22_out.php", {
                          imageData: data
                       });
                 });
           });      
         
               
            
              
    // WRITE
   // chart.exportConfig = {
   //   menuItems: [{
   //       icon: 'amcharts/images/export.png',
   //       format: 'jpg',
   //       onclick: function(a) {
   //           var output = a.output({
   //              format: 'jpg',
   //               output: 'datastring'
   //          }, function(data) {
   //               $.post("erp_createpiechart_save.php", {
   //                       imageData: data
   //                   });
   //           });
   //       }
   //   }] 
   // };   
                      
    chart.write("chartdiv");
    chart.validateData();   
                            
  });

  
</script>
</head>
<body>  
  <!-- chart container -->
  <div id="chartdiv" style="width: 450px; height: 280px;"></div>

  <!-- the chart code -->
</body>
</html>