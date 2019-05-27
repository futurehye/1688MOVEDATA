unit uMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdHTTP,
  Vcl.OleCtrls, SHDocVw, Vcl.CheckLst, Vcl.Grids,vcl.dbgrids, Vcl.ComCtrls,
  Vcl.ExtCtrls, Vcl.ValEdit,Winapi.urlMon, Vcl.ExtDlgs,COMOBJ, Excel2000,math,inifiles,shellapi,
  IdIOHandler, IdIOHandlerSocket, IdIOHandlerStack, IdSSL, IdSSLOpenSSL,System.json,
  Data.DB;

type
  TForm1 = class(TForm)
    IdHttpListPage: TIdHTTP;
    pcShowInfo: TPageControl;
    TabSheet1: TTabSheet;
    GroupBox1: TGroupBox;
    wbShowProinfo: TWebBrowser;
    GroupBox2: TGroupBox;
    memListUrl: TMemo;
    BitBtn3: TBitBtn;
    GroupBox3: TGroupBox;
    sgShowTitle: TStringGrid;
    Panel1: TPanel;
    TabSheet2: TTabSheet;
    e: TGroupBox;
    GroupBox4: TGroupBox;
    BitBtn2: TBitBtn;
    TabSheet3: TTabSheet;
    TabSheet4: TTabSheet;
    BitBtn4: TBitBtn;
    odFileBox: TOpenDialog;
    leModeFile: TLabeledEdit;
    lePicPath: TLabeledEdit;
    leProdcutsFile: TLabeledEdit;
    btnModeFile: TBitBtn;
    btnPicPath: TBitBtn;
    btnProductsFlie: TBitBtn;
    GroupBox6: TGroupBox;
    edtSupplyName: TEdit;
    bntAddSupple: TBitBtn;
    cbSupplerName: TComboBox;
    GroupBox8: TGroupBox;
    cbProductsClass1: TComboBox;
    rgProductsclassSub: TRadioGroup;
    vleSupplerInfo: TValueListEditor;
    vleBaseProduceInfo: TValueListEditor;
    lbSelectTitle: TLabel;
    Memo1: TMemo;
    edtTaoBaoUrl: TEdit;
    Button2: TButton;
    Label1: TLabel;
    Button1: TButton;
    IdSSLIOHandlerSocketOpenSSL1: TIdSSLIOHandlerSocketOpenSSL;
    Button3: TButton;
    Memo2: TMemo;
    DBGrid1: TDBGrid;


    procedure BitBtn2Click(Sender: TObject);
    procedure sgShowTitleSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure sgShowTitleMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure FormCreate(Sender: TObject);
//    procedure sgShowTitleDrawColumnCell(Sender: TObject; const Rect: TRect;
//  DataCol: Integer; Column: TColumn; State: TGridDrawState);
//    procedure gridDrawCell(Sender: TObject; ACol, ARow: Integer;Rect: TRect; State: TGridDrawState);
    procedure sgShowTitleDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect; State:  TGridDrawState);
    procedure sgShowTitleClick(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
//    procedure cbSelPostageSelect(Sender: TObject);
    procedure vleBaseProduceInfo1SetEditText(Sender: TObject; ACol,
      ARow: Integer; const Value: string);
    procedure btnModeFileClick(Sender: TObject);
    procedure btnPicPathClick(Sender: TObject);
    procedure btnProductsFlieClick(Sender: TObject);
    procedure pcShowInfoChange(Sender: TObject);
    procedure cbProductsClass1Select(Sender: TObject);
    procedure cbProductsClass1MeasureItem(Control: TWinControl; Index: Integer;
      var Height: Integer);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure BitBtn5Click(Sender: TObject);

    procedure rgSupplerNameClick(Sender: TObject);
    procedure bntAddSuppleClick(Sender: TObject);
    procedure edtSupplyNameChange(Sender: TObject);

    procedure sgProduceInfoPreviewSelectCell(Sender: TObject; ACol,
      ARow: Integer; var CanSelect: Boolean);

    procedure BitBtn7Click(Sender: TObject);
    procedure BitBtn8Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);


 //   function GetProductsCode(str1:string):string;
  private
    { Private declarations }
  //  aPageUrlInfo:array of string;

  public
  //       function GetProductsCode(str1:string):string;
    { Public declarations }
  end;

var
  Form1: TForm1;
  CsvLines:TStringlist;    //用于导出到CSV
  CsvCName:TStringList;    //存放模板文档的 中文标识栏
  iCellMouseDown,iCol,iRow:integer;
  aPageUrlInfo:array of string;
  arrModeProdctusFile:array of array of string;
  arrProdctusInfo:array[0..11] of string;  //存放从产品页面得到的信息，包括标题、价格等
//  arrExcelRecode:array of array of string;
  fcheck,fnocheck:tbitmap;
implementation
 uses
   uPublicFun,uTaoBaoinfoprounit;

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
 var
   i:SmallInt;
   bmp:TBitmap;
 begin
   FCheck:= TBitmap.Create;
   FNoCheck:= TBitmap.Create;
   bmp:= TBitmap.create;
 try
bmp.handle := LoadBitmap( 0, PChar(OBM_CHECKBOXES ));
With FNoCheck Do Begin
width := bmp.width div 4;
height := bmp.height div 3;
canvas.copyrect( canvas.cliprect, bmp.canvas, canvas.cliprect );
End;
With FCheck Do Begin
width := bmp.width div 4;
height := bmp.height div 3;
canvas.copyrect(canvas.cliprect, bmp.canvas, rect( width, 0, 2*width, height ));
End;
finally
bmp.free
end;

sgShowtitle.ColWidths[0]:=30;
sgShowtitle.ColWidths[1]:=30;
sgShowtitle.ColWidths[2]:=470;
sgShowtitle.ColWidths[3]:=70;
sgShowtitle.ColWidths[4]:=1;          //用于保存标题的链接地址，实现隐藏
sgshowtitle.Cells[0,0]:='序号';
sgshowtitle.Cells[1,0]:='yes';
sgshowtitle.Cells[2,0]:='标题';
sgshowtitle.Cells[3,0]:='价格';

end;

procedure TForm1.FormShow(Sender: TObject);
begin
pcShowinfo.Width:=form1.Width;
pcshowinfo.Height:=form1.Height;

DeleteDirectory(getcurrentdir+'\Addproducts');  //清空目录
DeleteDirectory(getcurrentdir+'\PIC');  //清空目录
end;

procedure TForm1.pcShowInfoChange(Sender: TObject);
begin
//  if pcshowinfo.ActivePageIndex=1  then

// showmessage('dd');
end;

procedure TForm1.rgSupplerNameClick(Sender: TObject);
begin
  // rgSupplerName.Items.
end;





procedure TForm1.sgProduceInfoPreviewSelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
begin
   iCol:=Acol;
   irow:=arow;
end;


procedure TForm1.sgShowTitleClick(Sender: TObject);
var
 i,iSelectNum:integer;

 begin


 if sgshowtitle.col=1 then begin
  if sgshowtitle.Cells[sgshowtitle.col,sgshowtitle.row]='yes' then
    sgshowtitle.Cells[sgshowtitle.col,sgshowtitle.row]:='no'
  else
    sgshowtitle.Cells[sgshowtitle.col,sgshowtitle.row]:='yes';
   for I := 1 to sgshowtitle.Rowcount do
     begin
       if sgshowtitle.Cells[1,i]='yes' then
          iSelectNum:=iSelectNum+1;
     end;
   lbSelectTitle.Caption :='当前共选中:'+inttostr(iSelectNum)+'条。';
   iSelectNum:=0;
  end;

 if sgshowtitle.Col=2 then begin
 //  showmessage(sgshowtitle.Cells[4,sgshowtitle.Row]);
   wbShowProinfo.Navigate(sgshowtitle.Cells[4,sgshowtitle.Row]);
    end;

 //   clbselinfo.Items.
 if (sgshowtitle.Row=0) and (sgshowtitle.Col=1)  then
      for I := 1 to sgshowtitle.RowCount  do
       sgshowtitle.Cells[1,i]:=sgshowtitle.Cells[1,0];

end;

procedure TForm1.sgShowTitleDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect; State:  TGridDrawState);
begin

 if Acol=1 then begin
  if not (gdFixed in State) then
    with TStringGrid(Sender).Canvas do begin
      brush.Color:=clWindow;
      FillRect(Rect);
      if sgshowtitle.Cells[ACol,ARow]='yes' then
        Draw( (rect.right + rect.left - FCheck.width) div 2, (rect.bottom + rect.top - FCheck.height) div 2, FCheck )
      else
        Draw( (rect.right + rect.left - FCheck.width) div 2, (rect.bottom + rect.top - FCheck.height) div 2, FNoCheck );
    end;
  end;

end;

{

//searchstringSaveMark 这个函数是根据提供的前特征字符串和后特征字符串得到新的串
//这个函数将保留前后的特征字符串
function  SearchStringSavemark(s0,sBegin,sEnd:string):string ;
var
 iposbegin,iposend:integer;

begin
   iposbegin:=pos(sBegin,s0);
   s0:=copy(s0,iposbegin,length(s0)-iposbegin);
   iposEnd:=pos(sEnd,s0);
   result:=system.Copy(s0,1,iposend+length(send));


end;
//searchstring 这个函数是根据提供的前特征字符串和后特征字符串得到新的串
//这个函数将不保留前后的特征字符串
function  SearchString(s0,sBegin,sEnd:string):string ;
var
 iposbegin,iposend:longint;
 ts:Tstringlist;

begin
   iposbegin:=pos(sBegin,s0)+length(sbegin);
   s0:=copy(s0,iposbegin,length(s0)-iposbegin);
   iposEnd:=pos(sEnd,s0);
   result:=system.Copy(s0,1,iposend-1);
  
end;       }

//得到阿里巴巴批发产品列表页中的分页链接地址
//返回数据说明：函数本身返回一个布尔变量，当有页（多于一页）返回1；当没有页返回0
//如果有分页，同时还通过数组aPageUrlInfo，返回每个分页的信息
function GetAlibabaListpageUrl(sSourstring:string):boolean;
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


//函数GetAlibabaListinfo得到阿里巴巴批发产品列表页中的产品
//参数说明：sPageUrl为网址链接串
procedure GetAlibabaListinfo(sPageUrl:string);
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

   if GetAlibabaListpageUrl(spageinfo.text)  then
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
      stmpall:=searchstringsavemark(Spageinfo.text,'<ul class="offer-list-row">','</li>'+chr(13)+chr(10)+'			    </ul>'+chr(13)+chr(10)+'</div>');
   //     stmpall:=searchstringsavemark(Spageinfo.text,'<ul class="offer-list-row">','</ul>');

       while pos('</li>'+chr(13)+chr(10)+'			    </ul>',stmpall)>0 do
       begin
       stmp1:=searchstringsavemark(stmpall,'<li','</li>');
    //   showmessage(stmp1);
       sUrl:=searchstring(stmp1,'<a href="','" title="');
       sTitle:=searchstring(stmp1,'" title="','" target=');
       sPrice:=searchstring(stmp1,'<em>','</em>');
       ipos1:=pos('</li>',stmpall);
       stmpall:=copy(stmpall,ipos1+5,length(stmpall)-ipos1+5);

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
//函数ProAlibabaProductShowInfo(sprourl,sPicPath:string)处理产品页的图片信息
//参数spropageinfo为产品信息页源码在预处理后的处理码（主要是去除了'\')
//主要处理原理：将源码中涉及的图片下载到指定目录中，并替换源码中的图片路径
function  ProAlibabaProductShowInfo(sprourl,sPicPath:string):string;
 var
   sTmp,sDownPic,sDownPicFilename,spicfile,spropageinfo,sNewPageInfo,sDownPicMark:string;
   iPos1,inum,i,iPosPic:integer; //iPosPic用于判断字符串中是否含有问号的字符串
   arrPicFilepath:array of array of string;

begin
  // form1.Memo1.Lines.Add('链接：'+sprourl);
   //********sprourl 是产品详情介绍链接****************
   inum:=0;
   setlength(arrPicFilepath,100,2);  //设置一个二维数组
   spropageinfo:=form1.IdHttpListPage.Get(sprourl);
   spropageinfo:=StringReplace (spropageinfo, '\', '', [rfReplaceAll]);//将文本中的所有'\'用''替换
   sTmp:=sPropageinfo;   //得到产品详情内容（html代码形式）

    //showmessage(stmp);
   //form1.Memo1.Lines.Add('stmp:'+stmp);
//   sTmp:=searchstring(spropageinfo,'var offer_details={"content":"','"};');  //去除代码中的    var offer_details={"content":" 标志
 //  form1.Memo1.Text:=spropageinfo;
   //  form1.Memo1.Lines.Add(sTmp);
  //    showmessage(stmp);
   while pos('src="',sTmp)>0 do
   begin
     inum:=inum+1;

//     showmessage(inttostr(pos(sTmp,'src="')));
     sDownPicMark:=searchstring(sTmp,'src="','"');    //获得每一张图片的路径,判断为以“图片文件名后缀点”为标志
     // showmessage(sdownpicmark);
      iPosPic:=Pos('.jpg',sDownPicMark);   //

     if iPosPic>0 then   //存在有.jpg?safasf  样式的字符串
       sDownPic:=copy(sDownPicMark,1,iPosPic+3);


     ipos1:=pos('.jpg',sTmp);
//     snewtmp:=copy(stmp,1,ipos1);          //得到
     sTmp:=copy(sTmp,ipos1+5,length(sTmp)-ipos1);
    // spicfile:=StrRScan(pchar(sDownPic),'/');
     sPicFile:=Format('%.4d',[inum])+'.jpg';  //根据获得的图片的顺序，生成图片文件名


     if copy(sPicPath,length(sPicPath)-1,1)<>'\' then   sPicPath:=sPicPath+'\';

     sDownPicFilename:=sPicPath+copy(spicfile,2,length(spicfile)-1);
     sDownPicFilename:=stringReplace(sDownPicFilename,'\\','\',[rfReplaceAll]);
     UrlDownloadToFile(nil,pchar(sdownPic),pchar(sDownPicFilename), 0, nil);//将图片文件下载到指定目录
        //form1.Memo1.Lines.Add('aaaaaaspicfile:'+sDownPic);

     arrPicFilepath[inum-1,0]:=sDownPicMark ; //将原源码中的关于图片路径的字符串放入数组中
     arrPicFilepath[inum-1,1]:='FILE:///'+sDownPicfilename;  //将存入本机图片的路径放入数组中

   end;

   //form1.Memo1.Lines.Add('eeeedfff');
   //将原源码中的网络图片路径全部换为本地图片路径
   for i := 0 to inum-1 do
   begin
   spropageinfo:=stringReplace(sPropageinfo,arrPicFilepath[i,0],arrpicfilepath[i,1],[rfReplaceAll]);
   end;

   spropageinfo:=copy(spropageinfo,31,length(spropageinfo)-30-3);

   spropageinfo:='<p><img src="http://img01.taobaocdn.com/imgextra/i1/1617533324/T2zlRUXepdXXXXXXXX_!!1617533324.gif"><img src="http://img03.taobaocdn.com/imgextra/i3/1617533324/T2wKHwXeRaXXXXXXXX_!!1617533324.jpg"></p>'+spropageinfo;
   spropageinfo:=spropageinfo+'<p><img align="absmiddle" src="http://img03.taobaocdn.com/imgextra/i3/1617533324/T2gRPcXb8bXXXXXXXX_!!1617533324.jpg" /></p>';

   result:=spropageinfo;
  // form1.Memo1.text:=spropageinfo;
  //  showmessage(result);
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

//函数ProAlibabaProductPageRviewPic主要用于处理阿里巴巴具体产品页中预览图片，一是
//将图片获取后存放到与CSV同名的目录（savepath）中，二是生成在淘宝助理中的新图片格式
//字符串即：文件名:1:n;其中N为预览显示中的具体位置
//参数str1为产品页信息的源代码，savepath为与csv文件同名的文件夹
function  ProAlibabaProductPageRviewPic(str1,savepath:string):string;
 var
  stmp,sdownpic,sCsvNewPicString,s1,sSavePriwePicPath,sDownPicFilename:string;
  ipos1,iitem:integer;
 begin
  stmp:=str1;
  iitem:=0;
  sSavePriwePicPath:=getcurrentdir+'\'+savepath;  //在当前目录下建立一个与csv同名的文件夹用于存放产品浏览图片文件
  if not DirectoryExists(sSavePriwePicPath) then  MKDIR(sSavePriwePicPath);
    showmessage(sSavePriwePicPath);
  while pos('"preview":"',sTmp)>0 do
   begin

     sDownPic:=searchstring(sTmp,'"preview":"','"');
     ipos1:=pos('"preview":"',sTmp);
//     snewtmp:=copy(stmp,1,ipos1);          //得到
     stmp:=copy(stmp,ipos1+11,length(stmp)-ipos1);
 //    spicfile:=StrRScan(pchar(sDownPic),'/');
     s1:= CreatRandomstr(25);
     sDownPicFilename:=savePath+'\'+s1+'.tbi';
     UrlDownloadToFile(nil,pchar(sdownPic),pchar(sDownPicFilename), 0, nil);//将图片文件下载到指定目录

     sCsvNewPicString:=sCsvNewPicString+s1+':1:'+inttostr(iitem)+':|;';
     iitem:=iitem+1;
     s1:='';
   end;
//   showmessage(s);
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

//函数GetAlibabaProductsinfo根据介绍产品的页得到阿里巴巴批发产品的具体信息
//参数说明：sUrl为产品网址链接字符串

function GetAlibabaProductsinfo(sProductUrl,strcsvmodeinfo:string;iExcelRow:integer):string;
 var
   s1,stmpall,stmp1,stmp2,stmp3,stmp4,stmp5,sUrl,sTitle,sPrice,sSavePicPath,sSavePriwePicPath,sProCode:string;
   //aCsvrecode: array[1..50] of string;
   arrProjinxiaoInfo:array[1..11] of string;  //用于存放需要加入进销文档的信息（其中0位存放供货商）
   i,iItems:integer;
   sPageinfo,sCSVMODEINFO,slTmp,aCsvRecode:Tstringlist;
   strModeCsv:TStringlist;
 begin
   sltmp:=tstringlist.Create;
   spageinfo:=Tstringlist.Create;
   SCSVMODEINFO:=TSTRINGLIST.Create;
   strModeCsv:=TStringList.Create;
   aCsvRecode:=TStringList.Create;
//   sltmp.LoadFromFile(strcsvmodeinfo);
//   sltmp.Delimiter:=',';

//   strModeCsv.LoadFromFile(strcsvmodeinfo);
   sCsvModeInfo.DelimitedText:=',';
   sCsvModeInfo.CommaText:=strcsvmodeinfo;    //将模板行数据序列化
       showmessage(inttostr(sCsvModeInfo.Count));

         //获取产品详细介绍页面的HTML代码，为获取信息作准备
          sproducturl:= form1.IdHttpListPage.URL.URLEncode(sproducturl);
          sPageinfo.Text :=form1.IdHttpListPage.Get(sProducturl); //初始地址得到的产品页面源码
//          form1.Memo1.Lines.Add(spageinfo.Text);
     //下面的代码就是根据需要修改的数据列，获取并赋予数据值
//         showmessage(csvcname);
   for i := 0 to CsvCName.Count-1  do
     begin
       s1:=CsvCName[i];
      //  showmessage(S1);
       if pos(s1,'宝贝标题')>0 then
         sCsvModeinfo[i]:=searchstring(spageinfo.Text,'<h1 class="d-title">','</h1>');  //得到标题'

       if pos(s1,'省')>0  then   //加入经销商省市信息
         stmp1:=searchstring(spageinfo.Text,'<meta name="location" content="province=','">');
         sCsvModeInfo[i]:=copy(stmp1,1,pos(' ',stmp1)-1);   //加入省份
         sCsvModeInfo[i+1]:=copy(stmp1,pos(' ',stmp1),length(stmp1)-pos(' ',stmp1)+1);  //加入城市

       if pos(s1,'宝贝价格')>0 then
         sCsvModeInfo[i]:=inttostr(ceil(strtofloat(searchstring(spageinfo.Text,',"price":"','"}'))*strtofloat('1.'+form1.vleBaseProduceInfo.Values['宝贝价格'])));  //计算得到新的价格

       if pos(s1,'宝贝数量')>0 then
         sCsvModeInfo[i]:='100';

       if pos(s1,'开始时间')>0 then
       //  sCsvModeInfo[i]:=datetostr(now)+timetostr(now);//form1.vleBaseProduceInfo.Values['开始时间'];   //开始时间
            showmessage(CsvCName[I]);

       if pos(s1,'宝贝描述')>0 then
         begin
         showmessage('dfasffa1111');
         stmp2:=searchstring(spageinfo.Text,'data-tfs-url="','" data-enable=');
               showmessage('dfdf');
            if stmp2<>'' then
              begin
              // CsvCName[1]:=ModifyFileNameString(CsvCName[1]);
               sSavePicPath:=getcurrentdir+'\PIC';  //在当前目录下建立一个以标题为名的文件夹用于存放产品介绍中的图片文件
               CreatMkDir(sSavePicPath);
               showmessage(ssavepicpath);
               sCsvModeInfo[i]:=ProAlibabaProductShowInfo(stmp2,sSavePicPath);    //加入宝贝描述
               showmessage('dafasdfwerwer');
              end;
         end;

       if pos(s1,'宝贝属性')>0 then
         showmessage('shux111');

       if pos(s1,'新图片')>0 then    //宝贝主图显示
         begin
          sCsvModeInfo[i]:=proAlibabaProductPageRviewPic(sPageinfo.Text,'Addproducts');   //新图片
          showmessage(sCsvModeInfo[i]);
         end;

       if pos(s1,'无线详情')>0 then     showmessage('shux222');


     end;

       //  aCsvrecode[29]:=proAlibabaProductPageRviewPic(sPageinfo.Text,'Addproducts');   //新图片

     sltmp.LoadFromFile('d:\1.csv');
     sltmp[1]:=aCsvRecode.CommaText;
     sltmp.SaveToFile('d:\1.csv');
     //***for i := 1 to 50 do  //取消所有的逗号标点
     //***   stmpall:=stmpall+StringReplace(aCsvrecode[i],',','',[rfReplaceAll])+',';

      //将信息加入进销库记录中
     { arrExcelRecode[iExcelRow,0]:=form1.cbSupplerName.items[form1.cbSupplerName.ItemIndex]; //供货商名字
      arrExcelRecode[iExcelRow,1]:=sProcode;
      arrExcelRecode[iExcelRow,2]:=form1.cbProductsClass1.items[form1.cbProductsClass1.ItemIndex]+' '+form1.rgProductsclassSub.items[form1.rgProductsclassSub.ItemIndex]; //产品分类
      arrExcelRecode[iExcelRow,3]:=aCsvrecode[1];
      stmp4:=searchstring(spageinfo.Text,'<td class="de-feature">品牌：','</td>');
      if (length(stmp4)>30) or (length(stmp4)<1) then  arrExcelRecode[iExcelRow,4]:='' else  arrExcelRecode[iExcelRow,4]:=stmp4;
      stmp5:=searchstring(spageinfo.Text,'<td class="de-feature">型号：','</td>');
      if (length(stmp5)>30) or (length(stmp5)<1) then  arrExcelRecode[iExcelRow,5]:='' else  arrExcelRecode[iExcelRow,5]:=stmp5;

      arrExcelRecode[iExcelRow,8]:='有';
      arrExcelRecode[iExcelRow,9]:=aCsvrecode[15];
      arrExcelRecode[iExcelRow,10]:='=HYPERLINK("'+sProductUrl+'","产品网址")';
      arrExcelRecode[iExcelRow,11]:='=HYPERLINK("'+sSavePicPath+'","图片资料")';
      }
  //    excelapp.activesheet.cells[12,5].value:='=HYPERLINK("http://www.126.com","图")';
  // excelapp.activesheet.cells[12,6].value:='=HYPERLINK("c:\","文件夹")';
   result:=stmpall;
   spageinfo.Free;
   SCSVMODEINFO.Free;
   sltmp.Free;
     end;

//将阿里巴巴中的产品信息成批量写入CSV文件
//参数Surl为某类文件的链接地址
procedure WriteAlibabaProductInfoToCsvFlie(sURL:STRING);
var
  CommaStr: TStringList;
  tstr1,tstrcsvmodeinfo:TStringList;
  strurl,stmp:string;
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
  commaStr:=TStringList.Create;
  tStrCsvModeinfo.LoadFromFile(getcurrentdir+'\ModeDemo.csv');  //为调试方便设置
 tStrCsvModeinfo.Delimiter:=chr(9);//.DelimitedText:=' ';

  tStrCsvmODEINFO.StrictDelimiter:=True;
//  tstrcsvmodeinfo.LineBreak:='#9';

  showmessage(tStrCsvModeinfo[3]);
  CsvCName:=TStringList.Create;
  CsvCName.CommaText:=tStrCsvModeInfo[2];
  CommaStr.CommaText:=tStrCsvModeInfo[3];
  showmessage(inttostr(csvcname.Count));
  showmessage(inttostr(CommaStr.Count));
  form1.memo2.Lines.AddStrings(tstrcsvmodeinfo);
  showmessage(inttostr(form1.Memo2.Lines.Count));
//  showmessage(tStrCsvmodeinfo[3]);
//  CsvLines.Add('version 1.00');   //第一行
//  CsvLines.Add(str1+str2+str3+str4+str5);   //第二行
//  CsvLines.Add('宝贝名称,宝贝类目,店铺类目,新旧程度,省,城市,出售方式,宝贝价格,加价幅度,宝贝数量,有效期,运费承担,平邮,EMS,快递,发票,保修,放入仓库,橱窗推荐,开始时间,宝贝描述,宝贝属性,邮费模版ID,会员打折,修改时间,上传状态,图片状态,返点比例,新图片,视频,销售属性组合,用户输入ID串,用户输入名-值对,商家编码,销售属性别名,代充类型,数字ID,本地ID,宝贝分类,账户名称,宝贝状态,闪电发货,新品,食品专项,尺码库,库存类型,库存计数,退换货承诺,物流体积,物流重量');//第三行
  //  Csvlines.Add(GetAlibabaProductsinfo(strurl,tStrCsvModeinfo[3],k)); //将对应链接的信息写入CSV模板文件中
  //GetAlibabaListinfo(surl);
  //   showmessage(inttostr(form1.sgShowTitle.RowCount));

  for  i:=1 to form1.sgShowTitle.RowCount do
    begin
    if form1.sgShowTitle.Cells[1,i]='yes'  then
     iUsedRecode :=iUsedRecode+1;                  //得到需要处理的有效记录
    end;
   setlength(arrExcelRecode,iUsedRecode,13);    //设置动态数组

//    showmessage('inttostriusedrecode');

 //*** for  i:=1 to form1.sgShowTitle.RowCount do
 //**   begin

  //####      stmp:=GetAlibabaProductsinfo(surl,tStrCsvModeinfo[3],k);
        showmessage(stmp);
        CsvLines.DelimitedText:=',';
        CsvLines.Add(stmp);
        //****Csvlines.Add(GetAlibabaProductsinfo(surl,tStrCsvModeinfo[3],k));
//***    end;
     {
    //将已经形成的数据写入预览修改表格中，供修改确认
    //将公共数组arrExcelRecode的对应数据写入表格中
    iRecode:=high(arrexcelrecode);
    form1.sgProduceInfoPreview.RowCount:=iRecode+2;
  //  showmessage(inttostr(irecode));

    for I := Low(arrexcelrecode)+1 to High(arrexcelrecode)+1 do
     begin
       form1.sgProduceInfoPreview.Cells[1,i]:=arrExcelRecode[i-1,3]; //显示标题
       form1.sgProduceInfoPreview.Cells[2,i]:=arrExcelRecode[i-1,6]; //进货价格
       form1.sgProduceInfoPreview.Cells[3,i]:=arrExcelRecode[i-1,7]; //销售价格
       fJprice:=strtofloat(arrExcelRecode[i-1,6]);
       fCprice:=strtofloat(arrExcelRecode[i-1,7]);
       form1.sgProduceInfoPreview.Cells[4,i]:=floattostr(fCprice-fJprice);  //价差
       form1.sgProduceInfoPreview.Cells[5,i]:=floattostr((fCprice-fJprice)/fJprice*100)+'%';//);  //毛利率
       form1.sgProduceInfoPreview.Cells[6,i]:=arrExcelRecode[i-1,10];  //产品网址
     end;
       form1.sgProduceInfoPreview.Options:=form1.sgProduceInfoPreview.Options+[goEditing];

 // CsvLines.SaveToFile(getcurrentdir+'\Addproducts.csv'); // 生成导入淘宝助理的CSV文件
 // WirteDataToExcel(form1.leProdcutsFile.text,form1.cbSupplerName.items[form1.cbSupplerName.ItemIndex]);
  showmessage('请对重点数据进行修改！');
  }
 // CsvLines.Free;
  CsvLines.SaveToFile(getcurrentdir+'\Addproducts.csv'); // 生成导入淘宝助理的CSV文件

  tStrCsvModeinfo.Free;
  tstr1.Free;

end;


procedure TForm1.BitBtn1Click(Sender: TObject);
var
 mark:string;
 i:integer;
begin

 // showmessage(CreatRandomNumstr(4));
  if  cbSupplerName.ItemIndex=-1 then  begin showmessage('请选择供应商');exit;end;
  if  cbProductsClass1.ItemIndex=-1 then begin showmessage('请选择大类');exit; end;
  if  rgProductsclassSub.ItemIndex=-1 then begin showmessage('请选择子类');exit; end;
   mark:='1'+format('%.2d',[cbProductsClass1.ItemIndex+1])+inttostr(rgProductsclassSub.ItemIndex+1)+Format('%.2d',[cbSupplerName.ItemIndex])+CreatRandomNumstr(6);

//   showmessage(mark);
end;

procedure TForm1.BitBtn2Click(Sender: TObject);
var
 surl:string;
begin

WriteAlibabaProductInfoToCsvFlie('');
//GetAlibabaProductsinfo
//ProAlibabaProductShowInfo('http://img03.taobaocdn.com/tfscom/T1OqOoXBtaXXXXXXXX','h:\abc\');
//WriteAlibabaProductInfoToCsvFlie('http://shop1363803581282.cn.1688.com/page/offerlist_16125642.htm?showType=&sortType=showcase');
 CsvLines.SaveToFile(getcurrentdir+'\Addproducts.csv'); // 生成导入淘宝助理的CSV文件
 showmessage('数据写入完成');
 CsvLines.Free;
end;



procedure TForm1.BitBtn3Click(Sender: TObject);
var
  iNum:integer;
begin
if memlisturl.Text=''  then
  begin
    showmessage('没有网址信息，请输入!');
    exit;
  end;

//******************判断是产品详细信息还是某类产品的链接信息***************
//判断是产品详细信息链接的方式是，链接中出现："detail"的关键词,如果是单条直接执行生成CSV文件，否则将该类中的信息显示到列表中


for inum:=0 to memlisturl.lines.count-1 do
 begin
    if pos('detail',memlisturl.Lines.Strings[inum])>0 then
     begin
       showmessage('dfdf');
     end
     else
      GetAlibabaListinfo(memlisturl.Lines.Strings[inum]);
  end;

//********************************************************


end;

procedure TForm1.BitBtn4Click(Sender: TObject);
begin
 pcshowinfo.ActivePageIndex:=1;
end;

procedure TForm1.BitBtn5Click(Sender: TObject);
var
// arrM:arrData;
 i,j:integer;
 iniInfo: TIniFile;
strSouce,S2,S3:STRING;
begin

//  showmessage(inttostr(sgProduceInfoPreview.RowCount));
{ for i :=1 to  sgProduceInfoPreview.RowCount-1   do
    begin

  strSouce:=Csvlines.Strings[i+2];

  csvlines.Strings[i+2]:=StringReplace(Csvlines.Strings[i+2],arrExcelRecode[i-1,3],sgProduceInfoPreview.Cells[1,i], [rfReplaceAll]);
  csvlines.Strings[i+2]:=StringReplace(Csvlines.Strings[i+2],','+arrExcelRecode[i-1,7]+',',','+sgProduceInfoPreview.Cells[3,i]+',', []);

  arrExcelRecode[i-1,3]:=sgProduceInfoPreview.Cells[1,i]; //将标题修改后的数据写入EXECEL文件数组中
  arrExcelRecode[i-1,6]:=sgProduceInfoPreview.Cells[2,i]; //进货价格
  arrExcelRecode[i-1,7]:=sgProduceInfoPreview.Cells[3,i]; //销售价格

  end;
 }
 //showmessage(strsouce+'  '+);

 CsvLines.SaveToFile(getcurrentdir+'\Addproducts.csv'); // 生成导入淘宝助理的CSV文件

// WirteDataToExcel(form1.leProdcutsFile.text,form1.cbSupplerName.items[form1.cbSupplerName.ItemIndex]);
 CsvLines.Free;
//  tStrCsvModeinfo.Free;
 // tstr1.Free;

   showmessage('数据写入完成');


//sgProduceInfoPreview.Options:=sgProduceInfoPreview.Options+[goEditing];
{ setlength(arrM,5,12);
  for I := 0 to 4 do
    for j := 0 to 11 do
      arrm[i,j]:=inttostr(i)+inttostr(j);
WirteDataToExcel('h:\pro.xls','DELPHI1',arrM);}
//iniinfo:=Tinifile.Create(getcurrentdir+'\baseinfo.ini');
//showmessage(iniinfo.ReadString ('PATH','ProPicPath',''));
//iniinfo.Free;
// S1:='<a href="http://ambopower.cn.1688.com/page/contactinfo.htm">联系方式</a>';
// S2:='"';
// S3:='">公司档案';
//RecallString(S1,s2,s3);
//CreatMkDir('h:\dsf\46  43\h  dgh\');
end;


procedure TForm1.BitBtn7Click(Sender: TObject);
begin
GettaobaoListinfo('http://hdeda.taobao.com/category-645688275-304914751.htm?spm=a1z10.5.w4010-505923791.19.l5d8qy&search=y&catName=%B8%F6%C8%CB%B6%A8%CE%BB%C6%F7#bd');
end;

procedure TForm1.BitBtn8Click(Sender: TObject);
begin
  WriteTaobaoProductInfoToCsvFlie('http://item.taobao.com/item.htm?spm=a1z10.3.w1017-2405381374.13.idV1I4&id=36007854547&');
end;

procedure TForm1.btnModeFileClick(Sender: TObject);
var
 iniinfo:TINIFILE;
// strClassInfo:TStrings;
begin
  iniinfo:=Tinifile.Create(getcurrentdir+'\baseinfo.ini');
//  strClassinfo:=TStringList.Create ;

   odfilebox.Filter:='淘宝助理模板文件(*.csv)|*.csv';
  if odfilebox.Execute then
    lemodefile.Text:=odfilebox.FileName;

  iniinfo.WriteString('path','CsvModeFilePath',lemodefile.text);
  iniinfo.Free;
end;

procedure TForm1.bntAddSuppleClick(Sender: TObject);
var
ExcelApp: Variant;
rowlast,collast,i,isSheet,m,n:integer;
begin

  if leProdcutsFile.text<>'' then
     begin
      ExcelApp := CreateOleObject( 'Excel.Application' );
      ExcelApp.WorkBooks.Open( leProdcutsFile.text );
      ExcelApp.WorkSheets['供货商信息'].Activate;
      rowlast:=excelapp.Cells.SpecialCells(xlCellTypelastCell, EmptyParam).row;
    //  showmessage(inttostr(rowlast));
       for I := 3 to rowlast do
        begin
          if ExcelApp.cells[i,3].value=vleSupplerInfo.Values['公司简称：'] then
           begin
             showmessage('该供应商信息已经存在!');
             cbsupplername.Text:=vleSupplerInfo.Values['公司简称：'];
             ExcelApp.WorkBooks.Close;
             ExcelApp.Quit;
             exit;
           end;

        end;

       ExcelApp.ActiveSheet.Rows[rowlast+1].Insert;// 从最后一行添加一个行
       ExcelApp.Cells[rowlast+1,1].Value:= inttostr(strtoint(ExcelApp.Cells[rowlast,1].Value)+1);
       ExcelApp.Cells[rowlast+1,2].Value:= vleSupplerInfo.Values['公司名称：'];
       ExcelApp.Cells[rowlast+1,3].Value:= vleSupplerInfo.Values['公司简称：'];
       ExcelApp.Cells[rowlast+1,4].Value:= inttostr(strtoint(ExcelApp.Cells[rowlast,1].Value));
       ExcelApp.Cells[rowlast+1,5].Value:= vleSupplerInfo.Values['地址：'];
       ExcelApp.Cells[rowlast+1,6].Value:= vleSupplerInfo.Values['电话：'];
  //     ExcelApp.Cells[rowlast+1,7].Value:= vleSupplerInfo.Values['传真：'];
       ExcelApp.Cells[rowlast+1,7].Value:= vleSupplerInfo.Values['联系人：'];
       ExcelApp.Cells[rowlast+1,8].Value:= vleSupplerInfo.Values['移动电话：'];
       ExcelApp.Cells[rowlast+1,9].Value:='=HYPERLINK("'+vleSupplerInfo.Values['供应产品：']+'","供应产品")';
   //    ExcelApp.Cells[rowlast+1,11].Value:='=HYPERLINK("'+vleSupplerInfo.Values['公司信用：']+'","公司信用")';

      ExcelApp.Activeworkbook.save;
      ExcelApp.WorkBooks.Close;
      ExcelApp.Quit;
      cbSupplerName.items.add(vleSupplerInfo.Values['公司简称：']);
      cbsupplername.Text:=vleSupplerInfo.Values['公司简称：'];
     end;
end;

procedure TForm1.btnPicPathClick(Sender: TObject);
var
  NewDir: string;
  iniinfo:TINIFILE;
begin
   iniinfo:=Tinifile.Create(getcurrentdir+'\baseinfo.ini');

  if SelectFolderDialog(Handle, '选择描述图片存放文件夹', '', NewDir) then
    begin
     lePicPath.Text := NewDir;

    end;

  iniinfo.WriteString('path','ProPicPath',lepicpath.text);
  iniinfo.Free;

end;

procedure TForm1.btnProductsFlieClick(Sender: TObject);
var
  ExcelApp: Variant;
  rowlast,I:integer;
  iniinfo:TINIFILE;
// strClassInfo:TStrings;
begin
  iniinfo:=Tinifile.Create(getcurrentdir+'\baseinfo.ini');

   odfilebox.Filter:='Excel文件(*.xls)|*.xls';
  if odfilebox.Execute then
    leProdcutsfile.Text:=odfilebox.FileName;

    ExcelApp := CreateOleObject( 'Excel.Application' );
    ExcelApp.WorkBooks.Open( leProdcutsfile.text );
    ExcelApp.WorkSheets[ '供货商信息' ].Activate;
    rowlast:=excelapp.Cells.SpecialCells(xlCellTypelastCell, EmptyParam).row;  //获取EXECEL中最后一行
    cbSupplerName.Items.Clear;
    for I := 2 to ROWLAST do
       cbSupplerName.Items.Add(ExcelApp.Cells[I,3].Value);

    ExcelApp.WorkBooks.Close;
    ExcelApp.Quit;

    iniinfo.WriteString('path','ProManageFile',leProdcutsfile.text);
    iniinfo.Free;

end;

procedure TForm1.Button1Click(Sender: TObject);
var
 s1,sUrl:string;
begin
//GettaobaoProductsPageinfo('http://item.taobao.com/item.htm?spm=a1z10.3.w1017-2405381374.13.idV1I4&id=36007854547&');
//sUrl:='https://img.alicdn.com/tfscom/TB1wF6bOyLaK1RjSZFxXXamPFXa';
//ProAlibabaProductShowInfo(sUrl,'d:\img');
  sUrl:='https://detail.1688.com/offer/1270560502.html';
  WriteAlibabaProductInfoToCsvFlie(sUrl);
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  WriteTaobaoProductInfoToCsvFlie(edtTaobaoUrl.Text );

end;

procedure TForm1.Button3Click(Sender: TObject);
var
   tCsvMode,tName,tData:TStringList;
   JSONObject: TJSONObject; // JSON类
            i: Integer; // 循环变量
         temp: string; // 临时使用变量
    jsonArray: TJSONArray; // JSON数组变量
   tp: TJSONPair;

begin
   JSONObject:=TJSONObject.Create;

   tCsvMode:=TStringList.Create;
   tName:=TStringList.Create;
   tData:=TStringList.Create;
   tCsvMode.LoadFromFile('d:\1.csv');
   showmessage(tCsvMode[2]);
   showmessage(tCsvMode[3]);
   tName.DelimitedText:=',';
   tName.CommaText:=tCsvMode[2];
   tData.DelimitedText:=',';
   tData.CommaText:=tCsvMode[3];
   for i := 0 to tName.Count-1 do
     begin
     JSONObject.AddPair(tName[i],tData[i]);  // 把模板中的名字栏和 模板值JSON化
     end;
   showmessage(inttostr(JSONObject.Count));
 //  tp:=jsonobject.Get('宝贝名称');

   tp:=TJSONPair.Create('宝贝名称','asfasfasfsafasf');

//   JSONObject.SetPairs(tp);
//   JSONOBJECT.Values('宝贝名称').Value:='asfdafda';
   showmessage(JSONObject.GetValue('宝贝名称').ToString);

   memo2.Lines.Add(JSONObject.ToString);
   JSONObject.Free;
   tData.Free;
   tName.Free;
   tCsvMode.Free;
end;

procedure TForm1.cbProductsClass1MeasureItem(Control: TWinControl;
  Index: Integer; var Height: Integer);
begin
 //  showmessage(cbproductsclass1.Items.Strings[index]);
end;

procedure TForm1.cbProductsClass1Select(Sender: TObject);
var
iniinfo:Tinifile;
strClassinfo:Tstrings;
strTmp:string;
begin

  iniinfo:=Tinifile.Create(getcurrentdir+'\baseinfo.ini');
  strClassinfo:=tstringlist.Create;
  //iniinfo.ReadSectionValues('PROCLASS',strClassinfo);
  //showmessage(cbproductsclass1.Items[cbproductsclass1.ItemIndex]);
   strtmp:=iniinfo.ReadString('PROCLASS',cbproductsclass1.Items[cbproductsclass1.ItemIndex],'' );
   strclassinfo:=SplitString(strtmp,',');

       rgproductsclasssub.Items.Clear;
       rgproductsclasssub.Items:=strclassinfo;

    iniinfo.Free;
    strclassinfo.Free;
end;


procedure TForm1.edtSupplyNameChange(Sender: TObject);
begin
   vleSupplerInfo.Values['公司简称：']:=edtsupplyname.Text;
end;

procedure TForm1.sgShowTitleMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
 iCellMouseDown:=1;
end;

procedure TForm1.sgShowTitleSelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
  var
   surl:string;
begin
if (Acol=1) and (iCellMouseDown=1) then
begin
surl:=form1.sgShowtitle.Cells[Acol+2,Arow];
//showmessage(surl);
iCellMouseDown:=0;
end;
end;



procedure TForm1.vleBaseProduceInfo1SetEditText(Sender: TObject; ACol,
  ARow: Integer; const Value: string);
begin
// showmessage(value);
end;

end.
