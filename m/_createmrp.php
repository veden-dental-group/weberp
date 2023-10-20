<?        
  if ($ping !='Y') exit;               
  define('FPDF_FONTPATH','fpdf/font/');
  define("_SYSTEM_TTFONTS", "C:/Windows/Fonts/");

  require_once('fpdf/tfpdf.php');   

  if (class_exists('PDF')) {
  } else { 
    class PDF extends tFPDF   {  
      //Page header
      function Header()
      {    
          
          $this->AddFont('kaiu','','simhei.ttf',true); 
          $this->SetFont('kaiu','',10);
          $this->SetXY(20,10);  
          $this->Cell(0,10, '日期:'.$_SESSION['cdate'],0,0,'L'); 
          $this->SetXY(80,10);  
          $this->Cell(0,10,'製處:'. $_SESSION['classname'],0,0,'L'); 
          $this->SetXY(140,10);     
          $this->Cell(0,10,'工序:',0,0,'L');   
             
          $this->SetFont('kaiu','',10);    
          $this->SetXY(20,20);
          $header=array('序號','Case號','客戶','左上','右上','左下','右下','顆數','','');
          $size=array(10,30,30,20,20,20,20,10,75,25); 
          $align=array('C','C','C','C','C','C','C','C','C','C');
          for($i=0;$i<count($header);$i++) {
              $this->Cell($size[$i],5,$header[$i],'LBRT',0,$align[$i]);  
          }                
          $this->Ln();  
      }

      //Page footer
      function Footer()
      {
        //Go to 1.5 cm from bottom
        $this->SetY(-15);
        //Select Arial italic 8
        $this->SetFont('Arial','I',10);
        //Print centered page number
         $this->Cell(0,10,'Page '.$this->PageNo().'/{nb}',0,0,'C');     
      }
    }
}                         
    
//gen pdf
$pdf=new PDF('L','mm','A4');
$pdf->AliasNbPages();
$pdf->AddPage();    
$pdf->AddFont('kaiu','','simhei.ttf',true);   
$pdf->SetFont('kaiu','',10);

$datefilter   = $_SESSION['datefilter'];
$linefilter   = $_SESSION['linefilter'];
$clientfitler = $_SESSION['clientfilter'];
$redofilter   = $_SESSION['redofilter'];
$ms_host = "mrp"; //这里是ODBC的连接名称   
$ms_user = "sa"; //用户名
$ms_pass = ""; //密码  
$conn =odbc_connect($ms_host, $ms_user, $ms_pass);
$sql="select a.k_engcode code, d.c_sname name, c.k15_z01 z01, c.k15_z02 z02, " . 
   "c.k15_z03 z03, c.k15_z04 z04, c.k15_z05 z05, c.k15_z06 z06, c.k15_z07 z07, c.k15_z08 z08, c.k15_z09 z09, c.k15_z10 z10, " .
   "c.k15_z11 z11, c.k15_z12 z12, c.k15_z13 z13, c.k15_z14 z14, c.k15_z15 z15, c.k15_z16 z16, c.k15_z17 z17, c.k15_z18 z18, " . 
   "c.k15_z19 z19, c.k15_z20 z20, c.k15_z21 z21, c.k15_z22 z22, c.k15_z23 z23, c.k15_z24 z24, c.k15_z25 z25, c.k15_z26 z26, " . 
   "c.k15_z27 z27, c.k15_z28 z28, c.k15_z29 z29, c.k15_z30 z30, c.k15_z31 z31, c.k15_z32 z32 " . 
   "from sk01 a, sk012 b, sk015 c, sf02 d where a.k_slipno=b.k12_slipno and a.k_slipno=c.k15_slipno and a.k_cuscode=d.c_code " . $datefilter . $linefilter . $clientfilter . $redofilter . " order by code";

$rs=odbc_exec($conn,$sql);
$sn=0; 
$ttotal=0;
while (odbc_fetch_array($rs)) {    
      $sn++;
      
      $code   = trim(odbc_result($rs,"code"));  
      $name   = trim(odbc_result($rs,"name")); 
      $z01    = odbc_result($rs,"z01");  
      $z02    = odbc_result($rs,"z02");    
      $z03    = odbc_result($rs,"z03");
      $z04    = odbc_result($rs,"z04");   
      $z05    = odbc_result($rs,"z05");   
      $z06    = odbc_result($rs,"z06");   
      $z07    = odbc_result($rs,"z07");   
      $z08    = odbc_result($rs,"z08"); 
      $z09    = odbc_result($rs,"z09");   
      $z10    = odbc_result($rs,"z10");   
      $z11    = odbc_result($rs,"z11"); 
      $z12    = odbc_result($rs,"z12"); 
      $z13    = odbc_result($rs,"z13"); 
      $z14    = odbc_result($rs,"z14"); 
      $z15    = odbc_result($rs,"z15"); 
      $z16    = odbc_result($rs,"z16"); 
      $z17    = odbc_result($rs,"z17"); 
      $z18    = odbc_result($rs,"z18"); 
      $z19    = odbc_result($rs,"z19"); 
      $z20    = odbc_result($rs,"z20"); 
      $z21    = odbc_result($rs,"z21");   
      $z22    = odbc_result($rs,"z22");   
      $z23    = odbc_result($rs,"z23");   
      $z24    = odbc_result($rs,"z24");   
      $z25    = odbc_result($rs,"z25");   
      $z26    = odbc_result($rs,"z26");   
      $z27    = odbc_result($rs,"z27");   
      $z28    = odbc_result($rs,"z28");   
      $z29    = odbc_result($rs,"z29");   
      $z30    = odbc_result($rs,"z30");   
      $z31    = odbc_result($rs,"z31");   
      $z32    = odbc_result($rs,"z32");   
      
      $t1     = '';  
      $t2     = '';   
      $t3     = '';  
      $t4     = '';   
      $t      = 0;
      
      //左上
      if ($z01!='') {
          $t1 .= '8';
          $t++;
      }
      if ($z02!='') {
          $t1 .= '7';
          $t++;
      }
      if ($z03!='') {
          $t1 .= '6';
          $t++;
      }
      if ($z04!='') {
          $t1 .= '5';
          $t++;
      }
      if ($z05!='') {
          $t1 .= '4';
          $t++;
      }
      if ($z06!='') {
          $t1 .= '3';
          $t++;
      }
      if ($z07!='') {
          $t1 .= '2';
          $t++;
      }
      if ($z08!='') {
          $t1 .= '1';
          $t++;
      }
      
      if ($z09!='') {
          $t2 .= '1';
          $t++;
      }
      if ($z10!='') {
          $t2 .= '2';
          $t++;
      }
      if ($z11!='') {
          $t2 .= '3';
          $t++;
      }
      if ($z12!='') {
          $t2 .= '4';
          $t++;
      }
      if ($z13!='') {
          $t2 .= '5';
          $t++;
      }
      if ($z14!='') {
          $t2 .= '6';
          $t++;
      }
      if ($z15!='') {
          $t2 .= '7';
          $t++;
      }
      if ($z16!='') {
          $t2 .= '8';
          $t++;
      }
      
             
      if ($z17!='') {
          $t3 .= '8';
          $t++;
      }
      if ($z18!='') {
          $t3 .= '7';
          $t++;
      }
      if ($z19!='') {
          $t3 .= '6';
          $t++;
      }
      if ($z20!='') {
          $t3 .= '5';
          $t++;
      }
      if ($z21!='') {
          $t3 .= '4';
          $t++;
      }
      if ($z22!='') {
          $t3 .= '3';
          $t++;
      }
      if ($z23!='') {
          $t3 .= '2';
          $t++;
      }
      if ($z24!='') {
          $t3 .= '1';
          $t++;
      }
      
      //右下
      if ($z25!='') {
          $t4 .= '1';
          $t++;
      }
      if ($z26!='') {
          $t4 .= '2';
          $t++;
      }
      if ($z27!='') {
          $t4 .= '3';
          $t++;
      }
      if ($z28!='') {
          $t4 .= '4';
          $t++;
      }
      if ($z29!='') {
          $t4 .= '5';
          $t++;
      }    
      if ($z30!='') {
          $t4 .= '6';
          $t++;
      }
      if ($z31!='') {
          $t4 .= '7';
          $t++;
      }
      if ($z32!='') {
          $t4 .= '8';
          $t++;
      }                         
         
      $ttotal+=$t;
      
                                 
      $header=array();     
      $header[]= $sn;   
      $header[]= $code;
      $header[]= $name;          
      $header[]= $t1;
      $header[]= $t2;
      $header[]= $t3;     
      $header[]= $t4;     
      $header[]= $t;
      $header[]= '';     
      $header[]= '';          
      $pdf->AddFont('kaiu','','simhei.ttf',true); 
      $pdf->SetFont('kaiu','',10);   
      $pdf->SetX(20); 
      $size=array(10,30,30,20,20,20,20,10,75,25); 
      $align=array('C','C','C','C','C','C','C','R','C','C');
      for($i=0;$i<count($header);$i++) {
          $pdf->Cell($size[$i],5,$header[$i],'LBR',0,$align[$i]);  
      }  
         
      $pdf->Ln();   
}                
$header=array();     
$header[]= 'Total:';  
$header[]= $ttotal;
$header[]= '';     
$header[]= '';    
$pdf->AddFont('kaiu','','simhei.ttf',true);       
$pdf->SetFont('kaiu','',10);   
$pdf->SetX(20); 
$size=array(150,10,75,25); 
$align=array('C','R','C','C');
for($i=0;$i<count($header);$i++) {
    $pdf->Cell($size[$i],5,$header[$i],'LBR',0,$align[$i]);  
}  
         
$pdf->Ln();

$pdf->Output('CaseList.pdf', "D");
    
?>    

