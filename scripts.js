function ChangeColor(tableRow, highLight) {
    if (highLight) {
      tableRow.style.backgroundColor = '#ffffff';
      //tableRow.className='cursor';
    }else{
      tableRow.style.backgroundColor = '';
    }
 }
     
function validUserAdd() {
    errormsg=""; 
    if (document.getElementById("account").value =="") { 
        errormsg +=  "請輸入帳號 !!\n"; 
    } 
    if (document.getElementById("password1").value != document.getElementById("password2").value) { 
        errormsg +=  "密碼不相符 !!\n"; 
    }
    if ( errormsg != "" ) {
        alert(errormsg);           
        document.getElementById("password1").value='';
        document.getElementById("password2").value='';
        return false;
    }  
}    

function validChangeSfb03() {
    errormsg=""; 
    if (document.getElementById("sfa01").value =="") { 
        errormsg +=  "請輸入工單單號 !!\n"; 
    }  
    if (document.getElementById("sfa27").value =="") { 
        errormsg +=  "請輸入BOM料號 !!\n"; 
    }  
    if (document.getElementById("sfa03").value =="") { 
        errormsg +=  "請輸入新發料料號 !!\n"; 
    }  
    if ( errormsg != "" ) {
        alert(errormsg);            
        return false;
    }  
}    

function validPcAdd() {
    errormsg=""; 
    if (document.getElementById("departmentguid").value =="") { 
        errormsg +=  "請輸入部門 !!\n"; 
    } 
    if (document.getElementById("staffguid").value =="") { 
        errormsg +=  "請輸入使用者 !!\n"; 
    } 
    if (document.getElementById("pcname").value =="") { 
        errormsg +=  "請輸入主機名稱 !!\n"; 
    } 
    if (document.getElementById("login").value =="") { 
        errormsg +=  "請輸入登入帳號 !!\n"; 
    } 
    if (document.getElementById("email").value =="") { 
        errormsg +=  "請輸入Email !!\n"; 
    } 
    if (document.getElementById("model").value =="") { 
        errormsg +=  "請輸入主機型號 !!\n"; 
    } 
    if (document.getElementById("pcno").value =="") { 
        errormsg +=  "請輸入機器編號 !!\n"; 
    } 
    
    if ( errormsg != "" ) {
        alert(errormsg);
        return false;
    }  
}    


function validworklogAdd() {
    errormsg="";

    if (document.getElementById("department").value =="") { 
        errormsg +=  "請輸入部門 !!\n"; 
    } 

          if (document.getElementById("indate").value =="") { 
        errormsg +=  "日期 !!\n"; 
    } 
   

   
        if (document.getElementById("send").value =="") { 
        errormsg +=  "請輸入出货数!!\n"; 
    }  
        if (document.getElementById("arrive").value =="") { 
        errormsg +=  "請輸入到货数 !!\n"; 
    }  
        if (document.getElementById("delay").value =="") { 
        errormsg +=  "請輸delay !!\n"; 
    }   
        if (document.getElementById("back").value =="") { 
        errormsg +=  "請輸内返数 !!\n"; 
    }   
        if (document.getElementById("number").value =="") { 
        errormsg +=  "請輸入人员总计 !!\n"; 
    }     
        if (document.getElementById("induty").value =="") { 
        errormsg +=  "請輸入在职 !!\n"; 
    }    
        if (document.getElementById("offduty").value =="") { 
        errormsg +=  "請輸入离职 !!\n"; 
    }   
        if (document.getElementById("quite").value =="") { 
        errormsg +=  "請輸入辞职\离职 !!\n"; 
    }
            if (document.getElementById("starttime").value =="") { 
        errormsg +=  "上班时间 !!\n"; 
    }   
        if (document.getElementById("endtime").value =="") { 
        errormsg +=  "下班时间 !!\n"; 
    }
       
   
 if (document.getElementById("times").value =="") { 
        errormsg +=  "請輸入次数 !!\n"; 
    }    
 if (document.getElementById("reporter").value =="") { 
        errormsg +=  "請輸入报告人 !!\n"; 
    }  


    
    if ( errormsg != "" ) 
    {
        alert(errormsg);
        return false;
    }  
}    
      
      
function numberOnly(e,f) {        
    var key = window.event ? e.keyCode : e.which;        
    var keychar = String.fromCharCode(key);        
    var el = document.getElementById(f);  
    reg = /[0-9.]/;        
    var result = reg.test(keychar);        
    if(!result)         
        {             
            return false;        
        }        
    else        
        {            
            el.className = "";             
            return true;        
        }   
}
function numberOnlydash(e,f) {        
    var key = window.event ? e.keyCode : e.which;        
    var keychar = String.fromCharCode(key);        
    var el = document.getElementById(f);  
    reg = /[0-9-]/;        
    var result = reg.test(keychar);        
    if(!result)         
        {             
            return false;        
        }        
    else        
        {            
            el.className = "";             
            return true;        
        }   
}


/*页面里回车到下一控件的焦点*/
function enter2Tab(e){
    var k = window.event.keyCode;
    var a = document.activeElement; 
    if (a.tagName == "INPUT" && ( a.type == "text" || a.type == "password" || a.type == "checkbox" || a.type == "radio" ) || a.tagName == "SELECT")
        {
          if (k == 13) window.event.keyCode = 9;              
        }     
}
/*打开此功能请取消下行注释*/
document.onkeydown = enter2Tab;


