unit uPublicFun;
 //本单元主要存放软件需要的公共函数和过程

 interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdHTTP,ShlObj,comobj,Excel2000,
  Vcl.OleCtrls, SHDocVw, Vcl.CheckLst, Vcl.Grids,vcl.dbgrids, Vcl.ComCtrls,
  Vcl.ExtCtrls, Vcl.ValEdit,Winapi.urlMon, Vcl.ExtDlgs,jpeg;

 // type
 //   arrData = array of array of string;
    { Private declarations }
  //  aPageUrlInfo:array of string;

     function  SearchStringSavemark(s0,sBegin,sEnd:string):string;
     function  SearchString(s0,sBegin,sEnd:string):string ;
     function CreatRandomstr(strLong:integer): string;
     function SelectFolderDialog(const Handle: integer; const Caption: string;
        const InitFolder: WideString; var SelectedFolder: string): boolean;
     function CreatRandomNumstr(strLong:integer): string;
     procedure WirteDataToExcel(sexcelfile,ssheetname:string);
     function SplitString(const source, ch: string): TStringList;
     function ModifyFileNameString(strSouceFilename:string):string; //
     function  RecallString(strSource,strBeforeMark,strBackMark:STRING):string;
     procedure CreatMkDir(strPath:string);
     function DeleteDirectory(NowPath: string): Boolean; // 删除整个目录(不删文件夹）

     function SetProductsInfo(strcsvmodeinfo:string;iExcelRow:integer):string;
     function CompressImageFile(FileName: string;  Width, Height: integer; PressQuality:Integer= 90): Boolean;
     var
     arrExcelRecode:array of array of string;
     arrProdctusInfo:array[1..11] of string;  //存放从产品页面得到的信息，包括标题、价格等

implementation



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

end;

//生产有大小写字符的随机字符串
//参数 strLong:integer 为生成的串长度
function CreatRandomstr(strLong:integer): string;

{max length of generated password}
// const
 //   intMAX_PW_LEN = 50;
 var
    i: Byte;
    s: string;
 begin
    {if you want to use the 'A..Z' characters}
       s := 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
   {if you want to use the 'a..z' characters}
     s:=s + 'abcdefghijklmnopqrstuvwxyz';

     s := s + '0123456789';
    for i := 0 to strLong-1 do
      Result := Result + s[Random(Length(s)-1)+1];
 end;

 //用数字创建一个指定长度的随机字符串
 //参数strLong用于确定字符串的长度
 function CreatRandomNumstr(strLong:integer): string;

{max length of generated password}
// const
 //   intMAX_PW_LEN = 50;
 var
    i: Byte;
    s: string;
 begin
    {if you want to use the 'A..Z' characters}
    //   s := 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
   {if you want to use the 'a..z' characters}
    // s:=s + 'abcdefghijklmnopqrstuvwxyz';

     s :='0123456789';
    for i := 0 to strLong-1 do
      Result := Result + s[Random(Length(s)-1)+1];
 end;

//函数SplitString用于从源字符串中通过分隔字符获取分段的字符，并作为字符列表返回
//例如：SplitString('abc,bzcde,efg',','),则返回abc  bzcde  efg三个串。
function SplitString(const source, ch: string): TStringList;
var
  temp, t2: string;
  i: integer;
begin
  result := TStringList.Create;
  temp := source;
  i := pos(ch, source);
  while i <> 0 do
  begin
    t2 := copy(temp, 0, i - 1);
    if (t2 <> '') then
      result.Add(t2);
    delete(temp, 1, i - 1 + Length(ch));
    i := pos(ch, temp);
  end;
  result.Add(temp);
end;
//检查文件名中是否有非法字符，如果有则将字符替换为'&'，返回修改后的文件名
function ModifyFileNameString(strSouceFilename:string):string;
var
 i:integer;
const
  subNovaluestr='<>/\|:"*?';
begin
 for i := 1 to length(subNoValuestr) do
 begin
    if pos(subNovaluestr[i],strSouceFilename)>0 then
     strSouceFilename:=StringReplace(strSouceFilename,subnovaluestr[i],'&',[rfReplaceAll]);
 end;
result:=strSouceFilename;
end;

 //用于从对话框中获取文件夹路径名
function SelectFolderDialog(const Handle: integer; const Caption: string;
  const InitFolder: WideString; var SelectedFolder: string): boolean;
var
  BInfo: _browseinfo;
  Buffer: array [0 .. MAX_PATH] of Char;
  ID: IShellFolder;
  Eaten, Attribute: Cardinal;
  ItemID: PItemidlist;
begin
  Result := False;
  BInfo.HwndOwner := Handle;
  BInfo.lpfn := nil;
  BInfo.lpszTitle := Pchar(Caption);
  BInfo.ulFlags := BIF_RETURNONLYFSDIRS + BIF_NEWDIALOGSTYLE;
  SHGetDesktopFolder(ID);
  ID.ParseDisplayName(0, nil, PWideChar(InitFolder), Eaten, ItemID, Attribute);
  BInfo.pidlRoot := ItemID;
  GetMem(BInfo.pszDisplayName, MAX_PATH);
  try
    if SHGetPathFromIDList(SHBrowseForFolder(BInfo), Buffer) then
    begin
      SelectedFolder := Buffer;
      if Length(SelectedFolder) <> 3 then
        SelectedFolder := SelectedFolder + '\';
      Result := True;
    end
    else
    begin
      SelectedFolder := '';
      Result := False;
    end;
  finally
    FreeMem(BInfo.pszDisplayName);
  end;
end;


//从字符串中回溯获得标识字符串之间的字符，其中STRBACKMARK为主串，在strSOUCE中为唯一串
//参数：strSource为源串，sBefoureMark为前面标识串(一般为一个字符），STRBACKMARK为后标识串。
function  RecallString(strSource,strBeforeMark,strBackMark:string):string;
var
  i,istrlen:integer;
  str1,strtmp:string;
begin
  i:=0;
   if pos(strbackmark,strSource)=0 then
     begin
       result:='';
       exit;
     end;

    strtmp:=copy(strSource,1,pos(strbackmark,strsource)-1);
    istrlen:=length(strtmp);
    while copy(strtmp,istrlen-i,1)<>strbeforemark do
     begin
       str1:=copy(strtmp,istrlen-i,1)+str1;
       i:=i+1;
     end;
  result:=str1;
end;


procedure WirteDataToExcel(sexcelfile,ssheetname:string);
var
ExcelApp: Variant;
//sheet:variant;
 rowlast,collast,i,isSheet,m,n:integer;
 tmpsheetname:string;
//Temp_Worksheet: _WorkSheet;

begin
//showmessage('adddd');
issheet:=0;
ExcelApp := CreateOleObject( 'Excel.Application' );
ExcelApp.WorkBooks.Open( sexcelfile );
 for i := 1 to excelapp.WorkSheets.Count  do
   begin
    tmpSheetName := ExcelApp.WorkSheets[i].Name;;
       if tmpSheetName = sSheetName then
          begin
            ExcelApp.WorkSheets[sSheetname].Activate; //设置一个活动的Sheet
            issheet:=1;
            break;
          end;
   end;
  if issheet=0 then      //新增一个SHEET
     begin
       ExcelApp.WorkSheets.Add;
       Excelapp.workbooks[1].sheets['sheet1'].name:=sSheetName;
      //  sheet:=Excelapp.workbooks[1].sheets['test'];
       ExcelApp.WorkSheets[sSheetname].Activate; //设置一个活动的Sheet
       //添加表头
       ExcelApp.ActiveSheet.cells[2,1]:='序号';
       ExcelApp.ActiveSheet.cells[2,2]:='商品编码';
       ExcelApp.ActiveSheet.cells[2,3]:='产品分类';
       ExcelApp.ActiveSheet.cells[2,4]:='宝贝标题';
       ExcelApp.ActiveSheet.cells[2,5]:='品牌';
       ExcelApp.ActiveSheet.cells[2,6]:='型号';
       ExcelApp.ActiveSheet.cells[2,7]:='进货价格';
       ExcelApp.ActiveSheet.cells[2,8]:='销售价格';
       ExcelApp.ActiveSheet.cells[2,9]:='是否有货';
       ExcelApp.ActiveSheet.cells[2,10]:='进货快递费';
       ExcelApp.ActiveSheet.cells[2,11]:='供货网址';
       ExcelApp.ActiveSheet.cells[2,12]:='产品图片';
       ExcelApp.ActiveSheet.cells[2,13]:='宝贝网址';

//    excelapp.activesheet.cells[12,5].value:='=HYPERLINK("http://www.126.com","图")';
//    excelapp.activesheet.cells[12,6].value:='=HYPERLINK("c:\","文件夹")';

      end;
         //ExcelApp.WorkSheets[SheetName].Delete;   //删除
      //将产品信息写入产品信息EXCEL文档中
         rowlast :=excelapp.Cells.SpecialCells(xlCellTypelastCell, EmptyParam).row;  //获取EXECEL中最后一行
         for m:= low(arrExcelRecode) to high(arrExcelRecode) do
          for n := low(arrExcelRecode[m]) to high(arrExcelRecode[m]) do
             begin
              ExcelApp.ActiveSheet.cells[rowlast+m+1,n+1]:=arrExcelRecode[m,n];
             end;

     //将供应商信息写入产品信息excel文档中

 ExcelApp.Activeworkbook.save;
 ExcelApp.WorkBooks.Close;
 ExcelApp.Quit;

end;

//根据字符串确定建立多重文件夹系统
procedure CreatMkDir(strPath:string);
var
 str1:string;
 i:integer;
 strPathNames:Tstringlist;
begin
 strpathnames:=tstringlist.Create;
   strpathnames:=SplitString(strpath,'\');
   for i := 0 to strpathnames.Count-1  do
      begin
       if i=0  then str1:=strpathnames[i] else
//         begin
          str1:=str1+'\'+strpathnames.Strings[i];
       if not DirectoryExists(str1) then  MKDIR(str1);
      end;
//          if not DirectoryExists(sSavePicPath) then  MKDIR(sSavePicPath);


 strpathnames.Free;

 end;

 //函数SetProductsinfo根据得到的产品信息页的信息（arrProdctusInfo数组），将信息填入“淘宝助理”CSV文档字段
//参数说明：sProductUrl为产品网址链接字符串
function SetProductsInfo(strcsvmodeinfo:string;iExcelRow:integer):string;
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
   SCSVMODEINFO.CommaText:=strcsvmodeinfo;
//   sproducturl:= form1.IdHttpListPage.URL.URLEncode(sproducturl);
//   sPageinfo.Text :=form1.IdHttpListPage.Get(sProducturl); //初始地址得到的产品页面源码
   //  showmessage(spageinfo.Text);

     aCsvrecode[1]:=arrProdctusInfo[1];  //得到标题'
     aCsvrecode[2]:=sCsvModeInfo[1];//form1.vleBaseProduceInfo.Values['宝贝类目'];   //宝贝类目
     aCsvrecode[3]:=sCsvModeInfo[2];   //店铺类目
     if sCsvModeInfo[3]<>'' then aCsvrecode[4]:=sCsvModeInfo[3] else aCsvrecode[4]:='1';   //新旧程度
     if sCsvModeInfo[4]<>'' then
      begin
      aCsvrecode[5]:=sCsvModeInfo[4];   //加入省份
      aCsvrecode[6]:=sCsvModeInfo[5];  //加入城市
       end;

     if sCsvModeInfo[6]<>'' then aCsvrecode[7]:=sCsvModeInfo[6] else aCsvrecode[7]:='1';   //新旧程度
//     aCsvrecode[7]:=form1.vleBaseProduceInfo.Values['出售方式'];   //出售方式

   //  arrExcelRecode[iExcelRow,6]:=arrProdctusInfo[2]; //得到批发价格
     aCsvrecode[8]:=arrProdctusInfo[3];  //填充新的价格（售价）

   //  arrExcelRecode[iExcelRow,7]:=aCsvrecode[8];   //销售价格
     aCsvrecode[9]:='';   //加价幅度
     if sCsvModeInfo[9]<>'' then aCsvrecode[10]:=sCsvModeInfo[9] else aCsvrecode[10]:='100';   //产品数量

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

  function DeleteDirectory(NowPath: string): Boolean; // 删除整个目录(不删文件夹）
var
  search: TSearchRec;
  ret: integer;
  key: string;
begin
  if NowPath[Length(NowPath)] <> '\' then
    NowPath := NowPath + '\';
  key := NowPath + '*.*';
  ret := findFirst(key, faanyfile, search);
  while ret = 0 do
  begin
    if ((search.Attr and fadirectory) = fadirectory) then
    begin
      if (search.Name <> '.') and (search.name <> '..') then
        DeleteDirectory(NowPath + search.name);
    end
    else
    begin
      if ((search.Attr and fadirectory) <> fadirectory) then
      begin
        deletefile(NowPath + search.name);
      end;
    end;
    ret := FindNext(search);
  end;
  findClose(search);
  //removedir(NowPath); 如果需要删除文件夹则添加
  result := True;
end;


/// <summary>
/// 压缩图片(BMP、JPG、PNG)
/// </summary>
/// <param name="FileName">文件路径</param>
/// <param name="Width">需要压缩后的宽度</param>
/// <param name="Height">需要压缩后的高度</param>
/// <param name="PressQuality">压缩质量</param>
/// <returns>是否压缩成功</returns>

function CompressImageFile(FileName: string;  Width, Height: integer; PressQuality:Integer= 90): Boolean;
   function GetNewSize(OldWidth, OldHeight: integer; NewWidth, NewHeight: integer; var RetWidth, RetHeight: integer):Boolean;
   var
       H:Boolean;
   begin
       Result := False;
       if (NewHeight < OldHeight) or (NewWidth < OldWidth) then
       begin
          H := NewHeight < OldHeight;

          if H then
          begin //按比例缩小,按高度来算高度的
             RetHeight := NewHeight;
             RetWidth := Round(OldWidth *  (NewHeight/OldHeight));
          end
          else
          begin //按比例缩小,按宽度来算宽度的
             RetWidth := NewWidth;
             RetHeight := Round(OldHeight * (NewWidth/OldWidth));
          end;
          Result:=True;
       end;
   end;
var
   bmp: TBitmap;
   jpg: TJpegImage;
//   png: TPNGGraphic;
   i: Integer;
   sTemp: string;
begin

   Result := False;
   try
      bmp := TBitmap.Create;
      jpg := TJPEGImage.Create;
    //  png := TPNGGraphic.Create;


      begin
         jpg.LoadFromFile(filename);
        // if GetNewSize(jpg.Width,jpg.height,Width,Height,Width,Height) then
        // begin
            bmp.height := Height;
            bmp.Width := Width;
            bmp.Canvas.StretchDraw(bmp.Canvas.ClipRect, jpg);
            jpg.Assign(bmp);
            jpg.CompressionQuality := PressQuality;
            jpg.Compress;
            sTemp := 'c:\aaa.jpg';//filename + '.lq';
            jpg.SaveToFile(sTemp);
            //DeleteFile(filename);
            //CopyFile(PChar(sTemp), PChar(filename), True);
            //DeleteFile(sTemp);
            Result := True;
        // end;

       end;
   finally
      FreeAndNil(bmp);
      FreeAndNil(jpg);
   //   FreeAndNil(png);
   end;
end;
end.
