<?php
  $data['username'] =array('eq','frank');
  $data['id']       =array("between","$bdate,$edate")    ;
  $data['id']       =arrary("not between",array($bdate,$edate));
  //DB_LIKE_FIELDS    =>'username|password';
  
  
  $data['id']       =array(array('gt',3),array('lt',10),'aaaa','or');
  $list=$user->where($data)->select;
  
  
  $data['username'] ='frank';
  $data['id']       ='1';
  $data['password'] ='aaa';
  $data['_logic']   ='or';
  $data['_string']  ='sql command';
  
  $list=$user->where($data)->select();
  
  $data['username']     =array('eq','frank');
  $data['password']     =array('like', 'w%');
  $data['_logic']       ='or';
  $where['_complexe']   =$data;
  $where['id']          =array('lt',5);
  
  
  
  //$user->min('id'), max('id'), avg('total'), sum('total'), count() ;
  
  
  // $user->getN(), ->first(), ->last()
  
  $user =M('user','CommModel') ;
  
  //$user=M():
  $list=$user->query("select * from think_user where id='1'");
  
  $user=M('user');
  $list=$user->getByUsername('aaaa');
  
  
  //三大自動
  
  //自動驗証 :在自定義 Model中實現
  
  class UserModel extends Model() {
      
      //自動驗証
      protected $_validate=array(
        //array(驗証字段, 驗証規則, 錯誤規則, 驗証條件, 附加規則, 驗証時機);
        array('id','>0', 'ID >0', 'condition', 'another rule', 'db when'),
        array(),
        array(),
        
        //系統有封裝了常用驗証規則
        // require , email, url, currency, number 
      
        array('username','require','用戶名不能為空白');
        array('username','checkLen','Username too short or too long!!');
      
      );
      
      
      //自動完成
      //array(填充字段, 填充內容, 填充條件, 附加規則);
      
      $user=new UserModel();
      $user=D('user','UserModel');
      
      if ($user->create()){
            if ($user->add()){
                $this->success('Data saved');
            } else {
                $this->error('Data saved error!!');
            }             
          
      } else {
        $this->error($user->getErrir())    ;
          
      }
      
      
      //無限級分類
      //遞歸方式
      //ajax方式
      //親緣關係方式  -->效率高 用SQL語法
      cate:     id
                name
                pid
                path
      
      select id, name, pid, path, concat*(path,'-',id) as route from cate order by route
      
      class CateAction extends Action (
        function index(){
            $this->display();            
        } 
      )   
  }
  
  
  
  
  
  
  //自動填充
 // protected $_auto=array(
  
  
  //導入非核心的封裝類
  import('ORG.Util.Image');
  import('@:Org.Image';
  
  Image::buildImageVerify();
  
  //分頁要手動導入
  import('Org.Util.Page');
  
  $show=$page->show();  //show 用來秀page
  
  
  );
  
  
  //自動映射
  
 
  //多語言 放在lang下
  //一個語言 一個目錄
  //一個模塊 一個檔案 語言的檔案名稱必須和模塊名稱一樣, 若為共用的, 則建一個檔案 檔名為common.php
  //內容為:
  
  return array(
    'welcome'->'Bonjour';  
  );
  
  //多語言支援
  
  index.html 內容
  <a href="?l=zh-cn"> 簡體中文 </a><br>
  <a href="?l=zh-tw"> 系體中文 </a><br>  
  <a href="?l=en-us"> English </a><br>  
  
  在模板中 使用以下指令  
  {$Think.lang.welcome}
  {$Think.lang.xxx}
  
  或者在控制器(action file)裡 用 L() 來定義要變換的東西
  L('welcome','歡迎日文')
  
  
  XxxModel.class.php 的處理方法為:       
  array('uname', 'require', '{%welcome}' )  -->用大括號加% 後面接著變量名稱 即可多語言
    
  配制文件 : conf/config.php
  'LANG_SWITCH_ON'      =>true,
  'DEFAULT_LANG'        =>'zh-cn',
  'LANG_AUTO_DETECT'    =>true,
  
  
  //多模板支持
  conf/config.php
  'TMPL_SWITCH_ON'     =>true,
  'TMPL_DETECT_THEME'  =>true, 
  
  
  index.html 內容
  <a href="?t=green"> 簡體中文 </a><br>
  <a href="?t=red"> 系體中文 </a><br>  
  <a href="?t=default"> English </a><br
  
  建議更改 template的分割字元 由 { } -> <{   }>
  
  在config中加
  TMPL_L_DELIM='<{';
  TMPL_R_DELIM='}>'; 
  
  
  //在模板中 可以用 . 來接收陣列中的資料
  如在action中原來是 
     $user=M('user');
     $list=$user->where('id=1') ->select();
     $this->assign($title, $list);           //傳過去的是數組
    在模板中可以
    {$title.id}
    {$title.password}
    
    如果傳過去的是 obj
    $user=M('user');
    $user->where('id=1')->select();
    $this->assign($title, $user);   //傳遞過去的是物件
    則在模板中用 : 
    {$title:id}
    {$title:password}
    
    在模板中可以使用php的函數 中間用 | 隔開, 如
    {$title:username|strlower|ucfirst} 
    
    {變量|函1|函2|函3=參數1, 參數2, ###} --> 3個###表示將前面的變量帶進來
    
    
    在模板中也可以使用函數, 格式如下 用: 開頭 后後函數名 
    {:function()}  -->如 {:time()} 會在傳回時間值  
    若用 {~function()}  則只會執行函數 但不會傳回值                                       
    
    {:U('user/insert')}
    
    
    模板注釋
    {/*  */}
    
    {//}
    
    
    模板中的系統變量
    {$Think.get.xxx} 用來取得$_GET傳入的資料
    {$Think.server.xxx}
    {$Think.session.xxx} 
    {$Think.env.xxx}  
    {$Think.cookie.xxx}     
    或
    {$_GET.xxx} 用來取得$_GET傳入的資料
    {$_SERVER.xxx}
    {$_SESSION.xxx} 
    {$_ENV.xxx}  
    {$_COOKIE.xxx} 
    
     若要引用常數        
     {$Think.const.__SELF__}
     或
     {$Think.__SELF__}
     {$Think.MODULE_NAME}
     
     {$Think.now}
     {$Think.version}
     {$Think.template|basename}
     
     讀取config訊息
     {$Think.config.db_host}
     {$Think.config.db_user}   
     
     {$Think.lang.xxx}
     
     快速輸出資料    但以下的輸出 不支援後面再加函數
     {@var}  輸出 session 中的變量
     {#var}  輸入 cookie 中的變量
     {&var}  輸出 config 中的變量
     {%var}  輸出 lang 中的變量
     {.var}  輸出 get  中的變量
     {^var}  輸出 post 中的變量
     {*var}  輸出 常量
    
     可以為變量設定初始值 如下
     {$title|default='這傢伙很懶 什麼都沒留下'}
     
      如何設定header, footer
      1. 在Tpl\Public下產生  header.html及footer.html
      2. 在 index.html中      
          <include file='./Tpl/Public/header' />
          ...
          ...
          ...
          <include file='./Tpl/Public/footer' />   //完整檔案路徑
      
       或者          
          <include file='Public:header' />    //跨模板
          
       或者 
          <include file='skinname@Public:header' />     跨皮膚
      
       或者
          <include file='$header' />  使用變量 但需在action中先assign 值
          
       導入 JS, CSS
       
       <import type="js", "css" file="Js.Util.aaa" />
       或
       
       
       <volist  name="list"  id="vo" offset="5" length="2" >       name是 action 傳過來的變數名 id是在以下迴圈使用的
         {$vo.username}       
       </volist>
       
       偶數輸出
       <volist name="list" id="vo" mod="2">  
            <eq name="mod" value="1">
                {<$vo.username}
            </eq
       </volist>   
          
       只輸出 key  -->會只輸出  1  2  3
       <volist name="list" id="vo" key="k"> 
                {$k}  
       </volist>  
       
       每隔5個值 跳行
       <volist name="list" id="vo" mod="5"> 
            {$vo.username}
            <eq name="mod" value="4">  
                <br>
            </eq>  
       </volist>    
       
       
       <foreach name="list" item="vo">
          {$vo.username}       
       </foreach>
       
       <switch name="">
            <case value="1">aaa</case>
            <case value="2">bbb</case> 
            <default />ccc</case>         
       </switch>
       
       //判斷大於 等於 小於 小於等於...
       eq neq  gt egt  lt  elt  heq nheq
       
       <eq name="mod" value="2"> 
            xxxx 
       <else />
            yyyy
       </eq>
       
       <eq name="mod|strlen" value="1">   -->可以加上 php 函數
            xxxx 
       <else />
            yyyy
       </eq>
       
       <in name="a" value="1,2,3"> xxx </in>
       <notin name="a" value="1,2,3"> xxx </notin> 
       <range name="a" value="1,3,4,5,6" type="in"> xxx </range> 
       <range name="a" value="1,3,4,5,6" type="notin"> xxx </range> 
       
       <present name="a"> xxxx </present>  判斷是否有值 (NOT NULL)
       <notpresent name="a"> yyyy </notpresent>
       
       <empty name="a"> 
            xxx        
       <else />
            yyyyy
       </empty>
    
        <notempty name="a"> 
            xxx        
       <else />
            yyyyy
       </notempty>
       
       
       /判斷常量是否有定義
       <defined name="MODULE_NAME"> xxx </defined>
       <notdefined name="MODULE_NAME"> xxx </notdefined>  
       
       <defined name="MODULE_NAME"> 
            xxxx
       <else />
            yyyy
       </defined> 
       
       <if condition="$vo['id] eq 5" >  ->>這裡的vo 儘量不要用點       
            aaaa
       <else />
            bbbbb       
       </if>
       
       原樣輸出
       <literal>
          <if condition=''>
          </if>
       </literal>
       
        如何擴展 Template的標籤
        1. <tagLib name="cx,html" />
        2. 在 template 中的寫法 , 以select 為例
            <html:select name="id" />
        
          
  <html>
  <head>
  
  </head>
  
  <body>
    <form name='form1' id='form1' action='__URL__/add' method="POST">
    
    
  
  
  
    </form>
  </body>
  
  
  </html>
  
?>
