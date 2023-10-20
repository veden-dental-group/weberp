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
          $this->SetXY(50,10);  
          $this->Cell(0,10,'製處:'. $_SESSION['classname'],0,0,'L'); 
          $this->SetXY(80,10);     
          $this->Cell(0,10,'工序:',0,0,'L');                            
          $this->SetXY(110,10);     
          $this->Cell(0,10,'組長簽名:',0,0,'L');                   
          $this->SetXY(160,10);     
          $this->Cell(0,10,'完成日期:',0,0,'L');  
             
          $this->SetFont('kaiu','',10);    
          $this->SetXY(10,20);
          $header=array('序號','客戶','Case號','簽核','產品名稱','顆數');
          $size=array(10,20,30,75,45,10); 
          $align=array('C','C','C','C','C','C');
          for($i=0;$i<count($header);$i++) {
              $this->Cell($size[$i],$_SESSION['point'],$header[$i],'LBRT',0,$align[$i]);  
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
$pdf=new PDF('P','mm','A4');
$pdf->AliasNbPages();
$pdf->AddPage();    
$pdf->AddFont('kaiu','','simhei.ttf',true);   
$pdf->SetFont('kaiu','',$_SESSION['font']);

$datefilter   = $_SESSION['datefilter'];
$linefilter   = $_SESSION['linefilter'];
$clientfitler = $_SESSION['clientfilter'];
$redofilter   = $_SESSION['redofilter'];
$ms_host = "mrp"; //这里是ODBC的连接名称   
$ms_user = "sa"; //用户名
$ms_pass = ""; //密码  
$conn =odbc_connect($ms_host, $ms_user, $ms_pass);
$sql="select a.k_slipno id, a.k_date cdate, a.k_engcode code, a.k_prdtcode pcode, a.k_qty qty, b.k12_class class, d.c_sname name, e.m1_prdtname pname " . 
     "from sk01 a, sk012 b, sk015 c, sf02 d, ps04 e where a.k_slipno=b.k12_slipno and a.k_slipno=c.k15_slipno and a.k_cuscode=d.c_code and a.k_prdtcode=e.m1_prdtcode " . $datefilter . $linefilter . $clientfilter . $redofilter . " order by (a.k_cuscode+a.k_engcode)";
   
$rs=odbc_exec($conn,$sql);
$sn=0; 
$ttotal=0;
$oldcode=''; 
while (odbc_fetch_array($rs)) {    
      $name   = iconv("big5","UTF-8",odbc_result($rs,"name"));                              
      $code   = trim(odbc_result($rs,"code"));  
      if ($code==$oldcode){
          $ecode='';
          $ename='';
          $esn='';
      } else {
          $ename=$name;
          $ecode=$code;
          $oldcode=$code;
          $sn++;
          $esn=$sn;
      }       
        
      $id     = odbc_result($rs,"id");
      $cdate  = substr(odbc_result($rs,"cdate"),0,10);   
      $class  = odbc_result($rs,"class");  
      $pcode  = odbc_result($rs,"pcode");
      $pname  = iconv("big5","UTF-8",odbc_result($rs,"pname"));
      $qty    = odbc_result($rs,"qty");     
      $ttotal += $qty;              
                                 
      $header=array();     
      $header[]= $esn;         
      $header[]= $ename;          
      $header[]= $ecode;   
      $header[]= '';
      $header[]= $pname;    
      $header[]= $qty; 
      $pdf->AddFont('kaiu','','simhei.ttf',true); 
      $pdf->SetFont('kaiu','',$_SESSION['font']);   
      $pdf->SetX(10); 
      $size=array(10,20,30,75,45,10); 
      $align=array('L','L','L','L','L','R');
      for($i=0;$i<count($header);$i++) {
          $pdf->Cell($size[$i],$_SESSION['point'],$header[$i],'LBR',0,$align[$i]);  
      }  
         
      $pdf->Ln();   
}                
$header=array();     
$header[]= 'Total:';  
$header[]= $ttotal;  
$pdf->AddFont('kaiu','','simhei.ttf',true);       
$pdf->SetFont('kaiu','',$_SESSION['font']);   
$pdf->SetX(10); 
$size=array(180,10); 
$align=array('C','R');
for($i=0;$i<count($header);$i++) {
    $pdf->Cell($size[$i],$_SESSION['point'],$header[$i],'LBR',0,$align[$i]);  
}  
         
$pdf->Ln();

$pdf->Output('CaseList.pdf', "D");
    
?>    

