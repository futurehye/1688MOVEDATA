unit uTaoBaoInfoProUnit;

interface
 uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdHTTP,ShlObj,comobj,Excel2000,
  Vcl.OleCtrls, SHDocVw, Vcl.CheckLst, Vcl.Grids,vcl.dbgrids, Vcl.ComCtrls,
  Vcl.ExtCtrls, Vcl.ValEdit,Winapi.urlMon, Vcl.ExtDlgs,math,umain,upublicfun,msxml;

  function  GetTaoBaoProductsPageinfo(sProductUrl:string):string;
  function  ProTaobaoProductShowInfo(sprourl,sPicPath:string):string;
  function  ProtaobaoProductPageRviewPic(str1,savepath:string):string;
  function  GetProductsCode(str1:string):string;
  function  GettaobaoListpageUrl(sSourstring:string):boolean;
  procedure GetTaobaoListinfo(sPageUrl:string);
  procedure WriteTaobaoProductInfoToCsvFlie(sURL:STRING);
  function SetTaoBaoProductsInfo(strcsvmodeinfo:tstrings):string;
 // function GetTaobaoProductsinfo(sProductUrl,strcsvmodeinfo:string;iExcelRow:integer):string;

implementation

 //函数GetTaobaoProductsinfo根据介绍产品的页得到淘宝产品的主要信息，并保持到arrProdctusInfo数组中
//参数说明：sProductUrl为产品网址链接字符串
//获取信息原理：
//关于获取”产品描述“的方法：第一步：根据产品网址得到网页源码
//                        第二步：在源码中搜索“g_config.dynamicScript("”和字符串““）”可以得到描述也的网址
//                        第三步：根据得到的描述网址，再取出源代码，并将代码中的”var desc='“和”';“字符串去掉

function GettaobaoProductsPageinfo(sProductUrl:string):string;
 var
   strFilePath,stmp2,sUrl,sTitle,sPrice,sSavePicPath,sSavePriwePicPath,sProCode:string;
//   aCsvrecode: array[1..50] of string;
//   arrProjinxiaoInfo:array[1..11] of string;  //用于存放需要加入进销文档的信息（其中0位存放供货商）
   i:integer;
   sPageinfo,sCSVMODEINFO:Tstringlist;
//   req: IXMLHttpRequest;
 begin
 for i := 1 to 11 do
    arrProdctusInfo[i]:='';

    spageinfo:=Tstringlist.Create;
    sproducturl:= form1.IdHttpListPage.URL.URLEncode(sproducturl);
    sPageinfo.Text :=form1.IdHttpListPage.Get(sProducturl); //初始地址得到的产品页面源码
   //  showmessage(spageinfo.Text);

     arrProdctusInfo[1]:=searchstring(spageinfo.Text,'<title>','-淘宝网</title>');  //得到标题'
   //  arrProdctusInfo[2]:=searchstring(spageinfo.Text,'<strong id="J_StrPrice" ><em class="tb-rmb">&yen;</em><em class="tb-rmb-num">','</em></strong>'); //得到批发价格
     arrProdctusInfo[2]:='0';
     arrProdctusInfo[3]:=inttostr(ceil(strtofloat(arrProdctusInfo[2])*strtofloat('1.'+form1.vleBaseProduceInfo.Values['宝贝价格'])));  //计算得到新的价格

     stmp2:=searchstring(spageinfo.Text,'g_config.dynamicScript("','")');  //得到含有产品描述的链接地址
     if stmp2<>'' then
       begin

         strFilePath:=ModifyFileNameString(arrProdctusInfo[1]);
         if form1.lePicPath.text='' then
         sSavePicPath:=getcurrentdir+'\'+form1.cbSupplerName.Text+'\'+strFilePath+'\'  //在当前目录下建立一个以标题为名的文件夹用于存放产品介绍中的图片文件
         else
         sSavePicPath:=form1.lePicPath.Text+form1.cbSupplerName.Text+'\'+strFilePath;  //在当前目录下建立一个以标题为名的文件夹用于存放产品介绍中的图片文件
  //      showmessage(sSavepicpath);
         CreatMkDir(sSavePicPath);
    // if not DirectoryExists(sSavePicPath) then  MKDIR(sSavePicPath);
          //  showmessage('sSavepicpath');

     arrProdctusInfo[4]:=ProtaobaoProductShowInfo(stmp2,sSavePicPath);    //加入宝贝描述信息，这里已经对图片路径进行了处理
     end
     else
     arrProdctusInfo[4]:='';

     arrProdctusInfo[5]:=protaobaoProductPageRviewPic(sPageinfo.Text,'Addproducts');   //新图片
    //**************
    // sProCode:=GetProductsCode(' ');
    // if sprocode<>'' then   arrProdctusInfo[6]:=sprocode;   // 商家编码
     arrProdctusInfo[6]:='';
    //********************

     arrProdctusInfo[7]:='';//searchstring(spageinfo.Text,'品牌:&nbsp;','</li>');
     arrProdctusInfo[8]:='';//searchstring(spageinfo.Text,'型号:&nbsp;','</li>');
     arrProdctusInfo[9]:='';//'=HYPERLINK("'+sProductUrl+'","产品网址")';
     arrProdctusInfo[10]:='';//'=HYPERLINK("'+sSavePicPath+'","图片资料")';
   //   arrProdctusInfo[11]:=form1.cbProductsClass1.items[form1.cbProductsClass1.ItemIndex]+' '+form1.rgProductsclassSub.items[form1.rgProductsclassSub.ItemIndex]; //产品分类

     arrProdctusInfo[11]:='';


     spageinfo.Free;
     result:='';

 //  SCSVMODEINFO.Free;
 //  sltmp.Free;
  end;


   //得到taobao产品列表页中的分页链接地址
//返回数据说明：函数本身返回一个布尔变量，当有页（多于一页）返回1；当没有页返回0
//如果有分页，同时还通过数组aPageUrlInfo，返回每个分页的信息
function GettaobaoListpageUrl(sSourstring:string):boolean;
var
   stmpall,spage,snextinfo,spagebefor,spageurl:string;
   ipos1,ipage,ipages:integer;
begin

   if pos('<ul data-sp="paging-a">',sSourstring)>0  then
   begin

   stmpall:=searchstringsavemark(sSourstring,'<ul data-sp="paging-a">','</ul>');

  snextinfo:=searchstring(stmpall,'<a class="next" href="','" >下一页</a>');
 //  showmessage(snextinfo);
   spage:=searchstring(stmpall,'<em class="page-count">','</em>页');//得到总页数
   ipages:=strtoint(spage);
   setlength(aPageUrlInfo,ipages);
   spagebefor:=copy(snextinfo,1,pos('pageNum=',snextinfo))+'&pageNum=';
   for ipage := 1 to ipages do
   begin
    spageurl:=spagebefor+inttostr(ipage) +'#search-bar';
    aPageUrlInfo[ipage-1]:=spageurl;
   end;
  result:=true;
   end
   else
   begin
   result:=false;   //
   end;
end;


//函数GettaobaoListinfo得到淘宝产品列表页中的产品
//参数说明：sPageUrl为网址链接串
procedure GetTaobaoListinfo(sPageUrl:string);
  var
   stmpall,stmp1,stmp2,sUrl,sPageurlinfos,sTitle,sPrice:string;
   ipos1,iPages,i,iItems:integer;
   sPageinfo:Tstringlist;
 //  idhttpurl:Tidhttp;
begin
   iItems:=1;
   spageinfo:=Tstringlist.Create;

   spageurl:= form1.IdHttpListPage.URL.URLEncode(spageurl);
   sPageinfo.Text :=form1.IdHttpListPage.Get(spageurl); //初始地址得到的页面源码
   // showmessage(spageurl);

 // 初始化供应商信息表
   //showmessage(inttostr(form1.vleSupplerInfo.RowCount));
    if form1.vleSupplerInfo.RowCount>1 then
     begin
      for I := 1 to form1.vleSupplerInfo.RowCount-1 do
          form1.vleSupplerInfo.DeleteRow(1);
     end;
    form1.vleSupplerInfo.InsertRow('联系方式地址：',Recallstring(sPageInfo.text,'"','">公司档案</a>'),true);
 //  form1.Memo3.Text:=spageinfo.Text;

   if GetTaobaoListpageUrl(spageinfo.text)  then
     ipages:=high(apageurlinfo)
     else
     ipages:=0;
//     form1.sgShowTitle.RowCount:=(ipages+1)*30;

   for i := 0 to ipages do
    begin
     if i>0 then
       begin
       sPageurlinfos:=apageurlinfo[i];
   //    form1.memListUrl.Lines.Add(spageurlinfos);
  //     spageurlinfos:= form1.IdHttpListPage.URL.URLEncode(spageurlinfos);
       spageinfo.text:=form1.IdHttpListPage.Get(sPageurlinfos);
       end;
      stmpall:=searchstringsavemark(Spageinfo.text,'<div class="shop-hesper-bd grid">','<!--END OF  pagination-->');
      //找出产品列表中的大范围产品介绍列表
   //     stmpall:=searchstringsavemark(Spageinfo.text,'<ul class="offer-list-row">','</ul>');

       while pos('<dd class="detail">',stmpall)>0 do       //末尾标记还在字符串中就循环执行
       begin
       stmp1:=searchstringsavemark(stmpall,'<dd class="detail">','</dd>');
     //  showmessage(stmp1);
       sUrl:=searchstring(stmp1,'href="','" target="_blank">');
       sTitle:=searchstring(stmp1,'" target="_blank">','</a>');
       sPrice:=searchstring(stmp1,'class="c-price">','</span>');
       ipos1:=pos('</dl>',stmpall);
       stmpall:=copy(stmpall,ipos1+5,length(stmpall)-ipos1+5);  //这里的5来自于</dl>的字节数

       //form1.Memo2.Lines.Add(stitle);
//       form1.clbshowtitle.Items.Add(inttostr(iitems)+'.'+stitle);
       form1.sgShowTitle.RowCount:=iitems+1;
//       showmessage(inttostr(iitems));
       form1.sgShowTitle.Cells[0,iitems]:=inttostr(iitems);
       form1.sgShowtitle.Cells[1,iitems]:='yes';
       form1.sgShowtitle.Cells[2,iitems]:=stitle;
       form1.sgShowtitle.Cells[3,iitems]:=sprice;
       form1.sgShowtitle.Cells[4,iitems]:=surl;
       iItems:=iItems+1;
       end;

   end;
    spageinfo.Free;
end;

//通过产品页中嵌入的链接地址，找到描述产品详细信息的网页源码，将其中的斜杠（"\"）
//全部用空格替换，然后将其中的图片链接找到，并下载到当前路径中，同时将图片链接
//进行相应的更改，返回的字符串作为产品信息，加入CSV文件中
//将需要定制的内容也加入到产品信息介绍中
//函数ProtaobaoProductShowInfo(sprourl,sPicPath:string)处理产品页的图片信息
//参数spropageinfo为产品信息页源码在预处理后的处理码（主要是去除了'\')
//主要处理原理：将源码中涉及的图片下载到指定目录中，并替换源码中的图片路径
function  ProTaobaoProductShowInfo(sprourl,sPicPath:string):string;
 var
   sTmp,sDownPic,sDownPicFilename,spicfile,spropageinfo,sNewPageInfo:string;
   iPos1,inum,i:integer;
   arrPicFilepath:array of array of string;
   req: IXMLHttpRequest;
begin

   inum:=0;
   setlength(arrPicFilepath,100,2);  //设置一个二维数组

   req := CoXMLHTTP.Create;
        req.open('Get',sprourl, False, EmptyParam, EmptyParam);
        req.send(EmptyParam);
        spropageinfo:=req.responseText;  //得到描述页源代码

 //  spropageinfo:=form1.IdHttpListPage.Get(sprourl);
   spropageinfo:=StringReplace (spropageinfo, '\', '', [rfReplaceAll]);//将文本中的所有'\'用''替换
   spropageinfo:=StringReplace (spropageinfo, #13, '', [rfReplaceAll]);//将文本中的所有'\'用''替换
   spropageinfo:=StringReplace (spropageinfo, #10, '', [rfReplaceAll]);//将文本中的所有'\'用''替换
 //  showmessage('字符串长度：'+inttostr(length(spropageinfo)));
   sTmp:=sPropageinfo;

 //  spropageinfo:=searchstring(spropageinfo,'var offer_details={"content":"','"};');
 //  form1.Memo1.Text:=spropageinfo;
   while pos('src="',sTmp)>0 do
   begin
   inum:=inum+1;
//     showmessage(inttostr(pos(sTmp,'src="')));
     sDownPic:=searchstring(sTmp,'src="','"');
     ipos1:=pos('<img',sTmp);
//     snewtmp:=copy(stmp,1,ipos1);          //得到
     stmp:=copy(stmp,ipos1+4,length(stmp)-ipos1);
     spicfile:=StrRScan(pchar(sDownPic),'/');
   //  showmessage(sPicPath);
     if copy(sPicPath,length(sPicPath)-1,1)<>'\' then   sPicPath:=sPicPath+'\';

     sDownPicFilename:=sPicPath+copy(spicfile,2,length(spicfile)-1);
   //  showmessage(sdownpicfilename);
     UrlDownloadToFile(nil,pchar(sdownPic),pchar(sDownPicFilename), 0, nil);//将图片文件下载到指定目录
     arrPicFilepath[inum-1,0]:=sDownPic ; //将原源码中的图片路径放入数组中
     arrPicFilepath[inum-1,1]:='File:///'+sDownPicfilename;  //将存入本机图片的路径放入数组中
    // showmessage(stmp);
   end;
   //将原源码中的网络图片路径全部换为本地图片路径
//  showmessage('adddbd');
   for i := 0 to inum-1 do
   begin
   spropageinfo:=stringReplace(sPropageinfo,arrpicfilepath[i,0],arrpicfilepath[i,1],[rfReplaceAll]);
   end;

   spropageinfo:=copy(spropageinfo,10,length(spropageinfo)-10-2);
 //  showmessage(spropageinfo);
   spropageinfo:='<p><img src="http://img01.taobaocdn.com/imgextra/i1/1617533324/T2zlRUXepdXXXXXXXX_!!1617533324.gif"><img src="http://img03.taobaocdn.com/imgextra/i3/1617533324/T2wKHwXeRaXXXXXXXX_!!1617533324.jpg"></p>'+spropageinfo;
   spropageinfo:=spropageinfo+'<p><img align="absmiddle" src="http://img03.taobaocdn.com/imgextra/i3/1617533324/T2gRPcXb8bXXXXXXXX_!!1617533324.jpg" /></p>';
   result:=spropageinfo;
 //  form1.Memo1.text:=spropageinfo;

end;

//通过产品页中嵌入的链接地址，找到描述产品详细信息的网页源码，将其中的斜杠（"\"）
//全部用空格替换，然后将其中的图片链接找到，并下载到当前路径中，同时将图片链接
//进行相应的更改，返回的字符串作为产品信息，加入CSV文件中
//将需要定制的内容也加入到产品信息介绍中
{function ProAlibabaProductInfo(sprourl:string):string;
var
 strtmp,strtmp1:string;
begin
 strtmp:=form1.IdHttpListPage.Get(sprourl);}
 //strtmp:=searchstring(strtmp,'var offer_details={','"};');
{ strtmp1:=StringReplace (Strtmp, '\', '', [rfReplaceAll]);//将文本中的所有'\'用''替换
 propagepicinfo(strtmp1,'g:\abc\');
end;}

//函数ProTaobaoProductPageRviewPic主要用于处理淘宝产品页中的预览图片，
//一是将图片获取后存放到与CSV同名的目录（savepath）中，
//二是生成在淘宝助理中的新图片格式字符串即：文件名:1:n;其中N为预览显示中的具体位置
//参数str1为产品页信息的源代码，savepath为与csv文件同名的文件夹
function  ProtaobaoProductPageRviewPic(str1,savepath:string):string;
 var
  sTmp,sdownpic,sCsvNewPicString,s1,sSavePriwePicPath,sDownPicFilename,sPicSize:string;
  ipos1,iitem:integer;
 begin
  sTmp:=str1;
  iitem:=0;
  sSavePriwePicPath:=getcurrentdir+'\'+savepath;  //在当前目录下建立一个与csv同名的文件夹用于存放产品浏览图片文件
  if not DirectoryExists(sSavePriwePicPath) then  MKDIR(sSavePriwePicPath);

  //sPicSize:=searchstring(sTmp,'<div class="tb-booth tb-pic tb-s','">');   //得到预览大图的边长尺寸
  //sPicSize:='.jpg_'+sPicSize+'x'+sPicSize+'.jpg';  //形成预览图片标志
  sPicSize:='.jpg_400x400.jpg';  //形成预览图片标志
  stmp:=searchstring(stmp,'<ul id="J_UlThumb" class="tb-thumb tb-clearfix">','</ul>');

  while pos('src="',sTmp)>0 do
   begin

     sDownPic:=searchstring(sTmp,'src="','.jpg_');

     sDownPic:=sDownPic+sPicSize;   //待下载的预览图片名（文件原名+尺寸x尺寸.JPG)
 //    form1.Memo1.Lines.Add('pic1:'+sDownpic);
     ipos1:=pos('src="',sTmp);
//     snewtmp:=copy(stmp,1,ipos1);          //得到
     stmp:=copy(stmp,ipos1+5,length(stmp)-ipos1);

 //    spicfile:=StrRScan(pchar(sDownPic),'/');
     s1:= CreatRandomstr(25);
     sDownPicFilename:=savePath+'\'+s1+'.tbi';
     UrlDownloadToFile(nil,pchar(sdownPic),pchar(sDownPicFilename), 0, nil);//将图片文件下载到指定目录

     sCsvNewPicString:=sCsvNewPicString+s1+':1:'+inttostr(iitem)+':|;';
     iitem:=iitem+1;
     s1:='';
   end;

    result:=sCsvNewPicString;
 end;

 //得到指定淘宝助理模板文件的信息
 //sModeFile为指定的模板文件
 //通过执行该过程，将模版主要数据存入本过程中使用的公共二元数组arrProductsModeInfo中
 procedure  GetTaobaozhuliModeinfo(sModeFile:string);
  begin
  // setlength(arrPicFilepath,10,1);  //设置一个二维数组

  end;

// 本函数根据输入的参数，得到商品的编码
//返回值为商品的编码
//编码规则为：店铺（1)+商品主类（2）+商品子类（1）+供货商（2）+ 随机码（6）
function GetProductsCode(str1:string):string;

begin
  if  form1.cbSupplerName.text='' then  begin showmessage('请选择供应商');result:='';exit;end;
 //showmessage(inttostr( form1.cbSupplerName.Items.IndexOf(form1.cbSupplerName.text)));
  if  form1.cbProductsClass1.text='' then begin showmessage('请选择大类');result:='';exit; end;
 //   form1.cbProductsClass1.Items.IndexOf(form1.cbProductsClass1.Text)
  if  form1.rgProductsclassSub.ItemIndex=-1 then begin showmessage('请选择子类');result:='';exit; end;
//   result:='A1'+format('%.2d',[form1.cbProductsClass1.ItemIndex+1])+inttostr(form1.rgProductsclassSub.ItemIndex+1)+Format('%.2d',[form1.cbSupplerName.ItemIndex])+CreatRandomNumstr(6);
   result:='A1'+format('%.2d',[form1.cbProductsClass1.ItemIndex+1])+inttostr(form1.rgProductsclassSub.ItemIndex+1)+Format('%.2d',[ form1.cbSupplerName.Items.IndexOf(form1.cbSupplerName.text)])+CreatRandomNumstr(6);
  //  showmessage(result);
end;



//将淘宝中的产品信息成批量写入CSV文件
//参数Surl为某类文件的链接地址
procedure WriteTaobaoProductInfoToCsvFlie(sURL:STRING);
var
  CommaStr: TStringList;
  tstr1,tstrcsvmodeinfo:tstrings;
  strurl:string;
  I,k,iUsedRecode,iRecode: Integer;
  fJprice,fCprice:single;
 // arrExcelData:arrData;
  const
         str1='title,cid,seller_cids,stuff_status,location_state,location_city,item_type,price,auction_increment,num,valid_thru,freight_payer,';
         str2='post_fee,ems_fee,express_fee,has_invoice,';
         str3='has_warranty,approve_status,has_showcase,list_time,description,cateProps,postage_id,has_discount,';
         str4='modified,upload_fail_msg,picture_status,auction_point,picture,video,skuProps,inputPids,inputValues,outer_id,propAlias,auto_fill,num_id,local_cid,navigation_type,';
         str5='user_name,syncStatus,is_lighting_consigment,is_xinpin,foodparame,features,global_stock_type,sub_stock_type,sell_promise,item_size,item_weight';

begin

  iUsedRecode:=0;
  k:=0;
  tStrCsvModeinfo:=Tstringlist.Create;  //用于处理csv模版文件的数据
  tstr1:=tstringlist.Create;
  CsvLines := TStringList.Create;

  if form1.leModeFile.text<>'' then
  tStrCsvModeinfo.LoadFromFile(form1.leModeFile.text)
  else
  begin
  showmessage('请输入模板文件！');
  exit;
  end;       //载入CSV文件模板

 { if (form1.vleBaseProduceInfo.Values['宝贝价格']='') or (form1.vleBaseProduceInfo.Values['宝贝价格']='0')   then
    begin
      showmessage('请输入宝贝加价比例！');
      exit;
    end;}

  CsvLines.Add('version 1.00');   //第一行
  CsvLines.Add(str1+str2+str3+str4+str5);   //第二行
  CsvLines.Add('宝贝名称,宝贝类目,店铺类目,新旧程度,省,城市,出售方式,宝贝价格,加价幅度,宝贝数量,有效期,运费承担,平邮,EMS,快递,发票,保修,放入仓库,橱窗推荐,开始时间,宝贝描述,宝贝属性,邮费模版ID,会员打折,修改时间,上传状态,图片状态,返点比例,新图片,视频,销售属性组合,用户输入ID串,用户输入名-值对,商家编码,销售属性别名,代充类型,数字ID,本地ID,宝贝分类,账户名称,宝贝状态,闪电发货,新品,食品专项,尺码库,库存类型,库存计数,退换货承诺,物流体积,物流重量');//第三行

  //GetAlibabaListinfo(surl);
{  for  i:=1 to form1.sgShowTitle.RowCount do
    begin
    if form1.sgShowTitle.Cells[1,i]='yes'  then
     iUsedRecode :=iUsedRecode+1;                  //得到需要处理的有效记录
    end;
 }
    //setlength(arrExcelRecode,iUsedRecode,13);    //设置动态数组

    GettaobaoProductsPageinfo(sUrl);   //得到制定网页的淘宝页面信息，并存入arrProdctusInfo数组中


    Csvlines.Add(SetProductsinfo(tStrCsvModeinfo[3],k));


  CsvLines.SaveToFile(getcurrentdir+'\Addproducts.csv'); // 生成导入淘宝助理的CSV文件
 // WirteDataToExcel(form1.leProdcutsFile.text,form1.cbSupplerName.items[form1.cbSupplerName.ItemIndex]);
 // showmessage('请对重点数据进行修改！');
  CsvLines.Free;
  tStrCsvModeinfo.Free;
  tstr1.Free;

end;

 //函数SetTaoBaoProductsinfo根据得到的产品信息页的信息（arrProdctusInfo数组），将信息填入“淘宝助理”CSV文档字段
//参数说明：sProductUrl为产品网址链接字符串
function SetTaoBaoProductsInfo(strcsvmodeinfo:tstrings):string;
 var
   stmpall,stmp1,stmp2,stmp3,stmp4,stmp5,sUrl,sTitle,sPrice,sSavePicPath,sSavePriwePicPath,sProCode:string;
   aCsvrecode: array[1..50] of string;
   arrProjinxiaoInfo:array[1..11] of string;  //用于存放需要加入进销文档的信息（其中0位存放供货商）
   i,iItems:integer;
   sPageinfo,sCSVMODEINFO,slTmp:Tstringlist;

 begin
   sltmp:=tstringlist.Create;
   spageinfo:=Tstringlist.Create;
   SCSVMODEINFO:=TSTRINGLIST.Create;
   showmessage(strcsvmodeinfo.CommaText);
   SCSVMODEINFO.CommaText:=strcsvmodeinfo.CommaText;
//   sproducturl:= form1.IdHttpListPage.URL.URLEncode(sproducturl);
//   sPageinfo.Text :=form1.IdHttpListPage.Get(sProducturl); //初始地址得到的产品页面源码
   //  showmessage(spageinfo.Text);
   showmessage('DDTT');
     aCsvrecode[1]:=arrProdctusInfo[1];  //得到标题'
     aCsvrecode[2]:=sCsvModeInfo[1];//form1.vleBaseProduceInfo.Values['宝贝类目'];   //宝贝类目
     aCsvrecode[3]:=sCsvModeInfo[2];   //店铺类目
     if sCsvModeInfo[3]<>'' then aCsvrecode[4]:=sCsvModeInfo[3] else aCsvrecode[4]:='1';   //新旧程度
     if sCsvModeInfo[4]<>'' then
      begin
      aCsvrecode[5]:=sCsvModeInfo[4];   //加入省份
      aCsvrecode[6]:=sCsvModeInfo[5];  //加入城市
       end;
      showmessage('AAAA');
     if sCsvModeInfo[6]<>'' then aCsvrecode[7]:=sCsvModeInfo[6] else aCsvrecode[7]:='1';   //新旧程度
//     aCsvrecode[7]:=form1.vleBaseProduceInfo.Values['出售方式'];   //出售方式

  //   arrExcelRecode[iExcelRow,6]:=arrProdctusInfo[2]; //得到批发价格
     aCsvrecode[8]:=arrProdctusInfo[3];  //填充新的价格（售价）

   //  arrExcelRecode[iExcelRow,7]:=aCsvrecode[8];   //销售价格
     aCsvrecode[9]:='';   //加价幅度
     if sCsvModeInfo[9]<>'' then aCsvrecode[10]:=sCsvModeInfo[9] else aCsvrecode[10]:='100';   //产品数量
           showmessage('bbb');
     aCsvrecode[11]:=sCsvModeInfo[10];//form1.vleBaseProduceInfo.Values['有效期'];;   //有效期
     aCsvrecode[12]:=sCsvModeInfo[11]; //    form1.cbSelPostage.Text;   //运费承担
     aCsvrecode[13]:=sCsvModeInfo[12]; ///form1.lePostnom.Text;   //平邮
     aCsvrecode[14]:=sCsvModeInfo[13];  //form1.lePostems.Text;   //EMS
     aCsvrecode[15]:=sCsvModeInfo[14];  //form1.lePostems.Text;   //快递S
     aCsvrecode[16]:=sCsvModeInfo[15]; //    form1.cbSelPostage.Text;   //运费承担
     aCsvrecode[17]:=sCsvModeInfo[16]; ///form1.lePostnom.Text;   //平邮
     aCsvrecode[18]:=sCsvModeInfo[17];  //form1.lePostems.Text;   //EMS
     aCsvrecode[19]:=sCsvModeInfo[18];  //form1.lePostems.Text;   //快递S
     aCsvrecode[20]:=datetostr(now)+timetostr(now);//form1.vleBaseProduceInfo.Values['开始时间'];   //开始时间

     aCsvrecode[21]:=arrProdctusInfo[4];    //加入宝贝描述

//     showmessage('222');
     aCsvrecode[22]:=sCsvModeInfo[21];   //宝贝属性
     aCsvrecode[23]:=sCsvModeInfo[22];   //邮费模版ID
     aCsvrecode[24]:=sCsvModeInfo[23];   //会员打折
     aCsvrecode[25]:=sCsvModeInfo[24];   //修改时间
     aCsvrecode[26]:=sCsvModeInfo[25];   // 上传状态
     aCsvrecode[27]:=sCsvModeInfo[26];   // 图片状态
     aCsvrecode[28]:=sCsvModeInfo[27];   // 返点比例

     aCsvrecode[29]:=arrProdctusInfo[5];   //新图片

     aCsvrecode[30]:=sCsvModeInfo[29];   // 视频
     aCsvrecode[31]:=sCsvModeInfo[30];   // 销售属性组合
     aCsvrecode[32]:=sCsvModeInfo[31];   // 用户输入ID串
     aCsvrecode[33]:=sCsvModeInfo[32];   // 用户输入名-值对

     aCsvrecode[34]:=arrProdctusInfo[6];   // 商家编码


     for I := 35 to 43 do
      aCsvrecode[i]:=sCsvModeInfo[i-1];   {分别代表销售属性别名、代充类型\数字ID本地ID
	宝贝分类、账户名称、宝贝状态、闪电发货、新品}
     aCsvrecode[44]:='';   //	食品专项
     aCsvrecode[45]:='';   //	尺码库
     aCsvrecode[46]:='';   //	库存类型
     aCsvrecode[47]:='';   //	库存计数
     aCsvrecode[48]:='';  // else aCsvrecode[48]:='0';   //退换货承诺
     aCsvrecode[49]:='';   //	物流体积
     aCsvrecode[50]:='';   //	物流重量

     for i := 1 to 50 do  //取消所有的逗号标点
        stmpall:=stmpall+StringReplace(aCsvrecode[i],',','',[rfReplaceAll])+',';

      //将信息加入进销库记录中
  //    arrExcelRecode[iExcelRow,0]:=form1.cbSupplerName.items[form1.cbSupplerName.ItemIndex]; //供货商名字
 {     arrExcelRecode[iExcelRow,1]:=arrProdctusInfo[6];   //商品编号
      arrExcelRecode[iExcelRow,2]:=arrProdctusInfo[11]; //产品分类
      arrExcelRecode[iExcelRow,3]:=arrProdctusInfo[1];
      if (length(arrProdctusInfo[7])>30) or (length(arrProdctusInfo[7])<1) then  arrExcelRecode[iExcelRow,4]:='' else  arrExcelRecode[iExcelRow,4]:=arrProdctusInfo[7];
      stmp5:=arrProdctusInfo[8];
      if (length(stmp5)>30) or (length(stmp5)<1) then  arrExcelRecode[iExcelRow,5]:='' else  arrExcelRecode[iExcelRow,5]:=stmp5;

      arrExcelRecode[iExcelRow,8]:='有';
      arrExcelRecode[iExcelRow,9]:=aCsvrecode[15];
      arrExcelRecode[iExcelRow,10]:=arrProdctusInfo[9];
      arrExcelRecode[iExcelRow,11]:=arrProdctusInfo[10];
  }
  //    excelapp.activesheet.cells[12,5].value:='=HYPERLINK("http://www.126.com","图")';
  // excelapp.activesheet.cells[12,6].value:='=HYPERLINK("c:\","文件夹")';

   result:=stmpall;
   spageinfo.Free;
   SCSVMODEINFO.Free;
   sltmp.Free;
  end;

end.
