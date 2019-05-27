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
  CsvLines:TStringlist;    //���ڵ�����CSV
  CsvCName:TStringList;    //���ģ���ĵ��� ���ı�ʶ��
  iCellMouseDown,iCol,iRow:integer;
  aPageUrlInfo:array of string;
  arrModeProdctusFile:array of array of string;
  arrProdctusInfo:array[0..11] of string;  //��ŴӲ�Ʒҳ��õ�����Ϣ���������⡢�۸��
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
sgShowtitle.ColWidths[4]:=1;          //���ڱ����������ӵ�ַ��ʵ������
sgshowtitle.Cells[0,0]:='���';
sgshowtitle.Cells[1,0]:='yes';
sgshowtitle.Cells[2,0]:='����';
sgshowtitle.Cells[3,0]:='�۸�';

end;

procedure TForm1.FormShow(Sender: TObject);
begin
pcShowinfo.Width:=form1.Width;
pcshowinfo.Height:=form1.Height;

DeleteDirectory(getcurrentdir+'\Addproducts');  //���Ŀ¼
DeleteDirectory(getcurrentdir+'\PIC');  //���Ŀ¼
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
   lbSelectTitle.Caption :='��ǰ��ѡ��:'+inttostr(iSelectNum)+'����';
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

//searchstringSaveMark ��������Ǹ����ṩ��ǰ�����ַ����ͺ������ַ����õ��µĴ�
//�������������ǰ��������ַ���
function  SearchStringSavemark(s0,sBegin,sEnd:string):string ;
var
 iposbegin,iposend:integer;

begin
   iposbegin:=pos(sBegin,s0);
   s0:=copy(s0,iposbegin,length(s0)-iposbegin);
   iposEnd:=pos(sEnd,s0);
   result:=system.Copy(s0,1,iposend+length(send));


end;
//searchstring ��������Ǹ����ṩ��ǰ�����ַ����ͺ������ַ����õ��µĴ�
//���������������ǰ��������ַ���
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

//�õ�����Ͱ�������Ʒ�б�ҳ�еķ�ҳ���ӵ�ַ
//��������˵��������������һ����������������ҳ������һҳ������1����û��ҳ����0
//����з�ҳ��ͬʱ��ͨ������aPageUrlInfo������ÿ����ҳ����Ϣ
function GetAlibabaListpageUrl(sSourstring:string):boolean;
var                                                  
   stmpall,spage,snextinfo,spagebefor,spageurl:string;
   ipos1,ipage,ipages:integer;
begin

   if pos('<ul data-sp="paging-a">',sSourstring)>0  then
   begin

   stmpall:=searchstringsavemark(sSourstring,'<ul data-sp="paging-a">','</ul>');

  snextinfo:=searchstring(stmpall,'<a class="next" href="','" >��һҳ</a>');
 //  showmessage(snextinfo);
   spage:=searchstring(stmpall,'<em class="page-count">','</em>ҳ');//�õ���ҳ��
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


//����GetAlibabaListinfo�õ�����Ͱ�������Ʒ�б�ҳ�еĲ�Ʒ
//����˵����sPageUrlΪ��ַ���Ӵ�
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
   sPageinfo.Text :=form1.IdHttpListPage.Get(spageurl); //��ʼ��ַ�õ���ҳ��Դ��
   // showmessage(spageurl);

 // ��ʼ����Ӧ����Ϣ��
   //showmessage(inttostr(form1.vleSupplerInfo.RowCount));
    if form1.vleSupplerInfo.RowCount>1 then
     begin
      for I := 1 to form1.vleSupplerInfo.RowCount-1 do
          form1.vleSupplerInfo.DeleteRow(1);
     end;
    form1.vleSupplerInfo.InsertRow('��ϵ��ʽ��ַ��',Recallstring(sPageInfo.text,'"','">��˾����</a>'),true);
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

//ͨ����Ʒҳ��Ƕ������ӵ�ַ���ҵ�������Ʒ��ϸ��Ϣ����ҳԴ�룬�����е�б�ܣ�"\"��
//ȫ���ÿո��滻��Ȼ�����е�ͼƬ�����ҵ��������ص���ǰ·���У�ͬʱ��ͼƬ����
//������Ӧ�ĸ��ģ����ص��ַ�����Ϊ��Ʒ��Ϣ������CSV�ļ���
//����Ҫ���Ƶ�����Ҳ���뵽��Ʒ��Ϣ������
//����ProAlibabaProductShowInfo(sprourl,sPicPath:string)�����Ʒҳ��ͼƬ��Ϣ
//����spropageinfoΪ��Ʒ��ϢҳԴ����Ԥ�����Ĵ����루��Ҫ��ȥ����'\')
//��Ҫ����ԭ����Դ�����漰��ͼƬ���ص�ָ��Ŀ¼�У����滻Դ���е�ͼƬ·��
function  ProAlibabaProductShowInfo(sprourl,sPicPath:string):string;
 var
   sTmp,sDownPic,sDownPicFilename,spicfile,spropageinfo,sNewPageInfo,sDownPicMark:string;
   iPos1,inum,i,iPosPic:integer; //iPosPic�����ж��ַ������Ƿ����ʺŵ��ַ���
   arrPicFilepath:array of array of string;

begin
  // form1.Memo1.Lines.Add('���ӣ�'+sprourl);
   //********sprourl �ǲ�Ʒ�����������****************
   inum:=0;
   setlength(arrPicFilepath,100,2);  //����һ����ά����
   spropageinfo:=form1.IdHttpListPage.Get(sprourl);
   spropageinfo:=StringReplace (spropageinfo, '\', '', [rfReplaceAll]);//���ı��е�����'\'��''�滻
   sTmp:=sPropageinfo;   //�õ���Ʒ�������ݣ�html������ʽ��

    //showmessage(stmp);
   //form1.Memo1.Lines.Add('stmp:'+stmp);
//   sTmp:=searchstring(spropageinfo,'var offer_details={"content":"','"};');  //ȥ�������е�    var offer_details={"content":" ��־
 //  form1.Memo1.Text:=spropageinfo;
   //  form1.Memo1.Lines.Add(sTmp);
  //    showmessage(stmp);
   while pos('src="',sTmp)>0 do
   begin
     inum:=inum+1;

//     showmessage(inttostr(pos(sTmp,'src="')));
     sDownPicMark:=searchstring(sTmp,'src="','"');    //���ÿһ��ͼƬ��·��,�ж�Ϊ�ԡ�ͼƬ�ļ�����׺�㡱Ϊ��־
     // showmessage(sdownpicmark);
      iPosPic:=Pos('.jpg',sDownPicMark);   //

     if iPosPic>0 then   //������.jpg?safasf  ��ʽ���ַ���
       sDownPic:=copy(sDownPicMark,1,iPosPic+3);


     ipos1:=pos('.jpg',sTmp);
//     snewtmp:=copy(stmp,1,ipos1);          //�õ�
     sTmp:=copy(sTmp,ipos1+5,length(sTmp)-ipos1);
    // spicfile:=StrRScan(pchar(sDownPic),'/');
     sPicFile:=Format('%.4d',[inum])+'.jpg';  //���ݻ�õ�ͼƬ��˳������ͼƬ�ļ���


     if copy(sPicPath,length(sPicPath)-1,1)<>'\' then   sPicPath:=sPicPath+'\';

     sDownPicFilename:=sPicPath+copy(spicfile,2,length(spicfile)-1);
     sDownPicFilename:=stringReplace(sDownPicFilename,'\\','\',[rfReplaceAll]);
     UrlDownloadToFile(nil,pchar(sdownPic),pchar(sDownPicFilename), 0, nil);//��ͼƬ�ļ����ص�ָ��Ŀ¼
        //form1.Memo1.Lines.Add('aaaaaaspicfile:'+sDownPic);

     arrPicFilepath[inum-1,0]:=sDownPicMark ; //��ԭԴ���еĹ���ͼƬ·�����ַ�������������
     arrPicFilepath[inum-1,1]:='FILE:///'+sDownPicfilename;  //�����뱾��ͼƬ��·������������

   end;

   //form1.Memo1.Lines.Add('eeeedfff');
   //��ԭԴ���е�����ͼƬ·��ȫ����Ϊ����ͼƬ·��
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

//ͨ����Ʒҳ��Ƕ������ӵ�ַ���ҵ�������Ʒ��ϸ��Ϣ����ҳԴ�룬�����е�б�ܣ�"\"��
//ȫ���ÿո��滻��Ȼ�����е�ͼƬ�����ҵ��������ص���ǰ·���У�ͬʱ��ͼƬ����
//������Ӧ�ĸ��ģ����ص��ַ�����Ϊ��Ʒ��Ϣ������CSV�ļ���
//����Ҫ���Ƶ�����Ҳ���뵽��Ʒ��Ϣ������
{function ProAlibabaProductInfo(sprourl:string):string;
var
 strtmp,strtmp1:string;
begin
 strtmp:=form1.IdHttpListPage.Get(sprourl);}
 //strtmp:=searchstring(strtmp,'var offer_details={','"};');
{ strtmp1:=StringReplace (Strtmp, '\', '', [rfReplaceAll]);//���ı��е�����'\'��''�滻
 propagepicinfo(strtmp1,'g:\abc\');
end;}

//����ProAlibabaProductPageRviewPic��Ҫ���ڴ�����Ͱ;����Ʒҳ��Ԥ��ͼƬ��һ��
//��ͼƬ��ȡ���ŵ���CSVͬ����Ŀ¼��savepath���У������������Ա������е���ͼƬ��ʽ
//�ַ��������ļ���:1:n;����NΪԤ����ʾ�еľ���λ��
//����str1Ϊ��Ʒҳ��Ϣ��Դ���룬savepathΪ��csv�ļ�ͬ�����ļ���
function  ProAlibabaProductPageRviewPic(str1,savepath:string):string;
 var
  stmp,sdownpic,sCsvNewPicString,s1,sSavePriwePicPath,sDownPicFilename:string;
  ipos1,iitem:integer;
 begin
  stmp:=str1;
  iitem:=0;
  sSavePriwePicPath:=getcurrentdir+'\'+savepath;  //�ڵ�ǰĿ¼�½���һ����csvͬ�����ļ������ڴ�Ų�Ʒ���ͼƬ�ļ�
  if not DirectoryExists(sSavePriwePicPath) then  MKDIR(sSavePriwePicPath);
    showmessage(sSavePriwePicPath);
  while pos('"preview":"',sTmp)>0 do
   begin

     sDownPic:=searchstring(sTmp,'"preview":"','"');
     ipos1:=pos('"preview":"',sTmp);
//     snewtmp:=copy(stmp,1,ipos1);          //�õ�
     stmp:=copy(stmp,ipos1+11,length(stmp)-ipos1);
 //    spicfile:=StrRScan(pchar(sDownPic),'/');
     s1:= CreatRandomstr(25);
     sDownPicFilename:=savePath+'\'+s1+'.tbi';
     UrlDownloadToFile(nil,pchar(sdownPic),pchar(sDownPicFilename), 0, nil);//��ͼƬ�ļ����ص�ָ��Ŀ¼

     sCsvNewPicString:=sCsvNewPicString+s1+':1:'+inttostr(iitem)+':|;';
     iitem:=iitem+1;
     s1:='';
   end;
//   showmessage(s);
    result:=sCsvNewPicString;

 end;

 //�õ�ָ���Ա�����ģ���ļ�����Ϣ
 //sModeFileΪָ����ģ���ļ�
 //ͨ��ִ�иù��̣���ģ����Ҫ���ݴ��뱾������ʹ�õĹ�����Ԫ����arrProductsModeInfo��
 procedure  GetTaobaozhuliModeinfo(sModeFile:string);
  begin
  // setlength(arrPicFilepath,10,1);  //����һ����ά����

  end;

// ��������������Ĳ������õ���Ʒ�ı���
//����ֵΪ��Ʒ�ı���
//�������Ϊ�����̣�1)+��Ʒ���ࣨ2��+��Ʒ���ࣨ1��+�����̣�2��+ ����루6��
function GetProductsCode(str1:string):string;

begin
  if  form1.cbSupplerName.text='' then  begin showmessage('��ѡ��Ӧ��');result:='';exit;end;
 //showmessage(inttostr( form1.cbSupplerName.Items.IndexOf(form1.cbSupplerName.text)));
  if  form1.cbProductsClass1.text='' then begin showmessage('��ѡ�����');result:='';exit; end;
 //   form1.cbProductsClass1.Items.IndexOf(form1.cbProductsClass1.Text)
  if  form1.rgProductsclassSub.ItemIndex=-1 then begin showmessage('��ѡ������');result:='';exit; end;
//   result:='A1'+format('%.2d',[form1.cbProductsClass1.ItemIndex+1])+inttostr(form1.rgProductsclassSub.ItemIndex+1)+Format('%.2d',[form1.cbSupplerName.ItemIndex])+CreatRandomNumstr(6);
   result:='A1'+format('%.2d',[form1.cbProductsClass1.ItemIndex+1])+inttostr(form1.rgProductsclassSub.ItemIndex+1)+Format('%.2d',[ form1.cbSupplerName.Items.IndexOf(form1.cbSupplerName.text)])+CreatRandomNumstr(6);
  //  showmessage(result);
end;

//����GetAlibabaProductsinfo���ݽ��ܲ�Ʒ��ҳ�õ�����Ͱ�������Ʒ�ľ�����Ϣ
//����˵����sUrlΪ��Ʒ��ַ�����ַ���

function GetAlibabaProductsinfo(sProductUrl,strcsvmodeinfo:string;iExcelRow:integer):string;
 var
   s1,stmpall,stmp1,stmp2,stmp3,stmp4,stmp5,sUrl,sTitle,sPrice,sSavePicPath,sSavePriwePicPath,sProCode:string;
   //aCsvrecode: array[1..50] of string;
   arrProjinxiaoInfo:array[1..11] of string;  //���ڴ����Ҫ��������ĵ�����Ϣ������0λ��Ź����̣�
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
   sCsvModeInfo.CommaText:=strcsvmodeinfo;    //��ģ�����������л�
       showmessage(inttostr(sCsvModeInfo.Count));

         //��ȡ��Ʒ��ϸ����ҳ���HTML���룬Ϊ��ȡ��Ϣ��׼��
          sproducturl:= form1.IdHttpListPage.URL.URLEncode(sproducturl);
          sPageinfo.Text :=form1.IdHttpListPage.Get(sProducturl); //��ʼ��ַ�õ��Ĳ�Ʒҳ��Դ��
//          form1.Memo1.Lines.Add(spageinfo.Text);
     //����Ĵ�����Ǹ�����Ҫ�޸ĵ������У���ȡ����������ֵ
//         showmessage(csvcname);
   for i := 0 to CsvCName.Count-1  do
     begin
       s1:=CsvCName[i];
      //  showmessage(S1);
       if pos(s1,'��������')>0 then
         sCsvModeinfo[i]:=searchstring(spageinfo.Text,'<h1 class="d-title">','</h1>');  //�õ�����'

       if pos(s1,'ʡ')>0  then   //���뾭����ʡ����Ϣ
         stmp1:=searchstring(spageinfo.Text,'<meta name="location" content="province=','">');
         sCsvModeInfo[i]:=copy(stmp1,1,pos(' ',stmp1)-1);   //����ʡ��
         sCsvModeInfo[i+1]:=copy(stmp1,pos(' ',stmp1),length(stmp1)-pos(' ',stmp1)+1);  //�������

       if pos(s1,'�����۸�')>0 then
         sCsvModeInfo[i]:=inttostr(ceil(strtofloat(searchstring(spageinfo.Text,',"price":"','"}'))*strtofloat('1.'+form1.vleBaseProduceInfo.Values['�����۸�'])));  //����õ��µļ۸�

       if pos(s1,'��������')>0 then
         sCsvModeInfo[i]:='100';

       if pos(s1,'��ʼʱ��')>0 then
       //  sCsvModeInfo[i]:=datetostr(now)+timetostr(now);//form1.vleBaseProduceInfo.Values['��ʼʱ��'];   //��ʼʱ��
            showmessage(CsvCName[I]);

       if pos(s1,'��������')>0 then
         begin
         showmessage('dfasffa1111');
         stmp2:=searchstring(spageinfo.Text,'data-tfs-url="','" data-enable=');
               showmessage('dfdf');
            if stmp2<>'' then
              begin
              // CsvCName[1]:=ModifyFileNameString(CsvCName[1]);
               sSavePicPath:=getcurrentdir+'\PIC';  //�ڵ�ǰĿ¼�½���һ���Ա���Ϊ�����ļ������ڴ�Ų�Ʒ�����е�ͼƬ�ļ�
               CreatMkDir(sSavePicPath);
               showmessage(ssavepicpath);
               sCsvModeInfo[i]:=ProAlibabaProductShowInfo(stmp2,sSavePicPath);    //���뱦������
               showmessage('dafasdfwerwer');
              end;
         end;

       if pos(s1,'��������')>0 then
         showmessage('shux111');

       if pos(s1,'��ͼƬ')>0 then    //������ͼ��ʾ
         begin
          sCsvModeInfo[i]:=proAlibabaProductPageRviewPic(sPageinfo.Text,'Addproducts');   //��ͼƬ
          showmessage(sCsvModeInfo[i]);
         end;

       if pos(s1,'��������')>0 then     showmessage('shux222');


     end;

       //  aCsvrecode[29]:=proAlibabaProductPageRviewPic(sPageinfo.Text,'Addproducts');   //��ͼƬ

     sltmp.LoadFromFile('d:\1.csv');
     sltmp[1]:=aCsvRecode.CommaText;
     sltmp.SaveToFile('d:\1.csv');
     //***for i := 1 to 50 do  //ȡ�����еĶ��ű��
     //***   stmpall:=stmpall+StringReplace(aCsvrecode[i],',','',[rfReplaceAll])+',';

      //����Ϣ����������¼��
     { arrExcelRecode[iExcelRow,0]:=form1.cbSupplerName.items[form1.cbSupplerName.ItemIndex]; //����������
      arrExcelRecode[iExcelRow,1]:=sProcode;
      arrExcelRecode[iExcelRow,2]:=form1.cbProductsClass1.items[form1.cbProductsClass1.ItemIndex]+' '+form1.rgProductsclassSub.items[form1.rgProductsclassSub.ItemIndex]; //��Ʒ����
      arrExcelRecode[iExcelRow,3]:=aCsvrecode[1];
      stmp4:=searchstring(spageinfo.Text,'<td class="de-feature">Ʒ�ƣ�','</td>');
      if (length(stmp4)>30) or (length(stmp4)<1) then  arrExcelRecode[iExcelRow,4]:='' else  arrExcelRecode[iExcelRow,4]:=stmp4;
      stmp5:=searchstring(spageinfo.Text,'<td class="de-feature">�ͺţ�','</td>');
      if (length(stmp5)>30) or (length(stmp5)<1) then  arrExcelRecode[iExcelRow,5]:='' else  arrExcelRecode[iExcelRow,5]:=stmp5;

      arrExcelRecode[iExcelRow,8]:='��';
      arrExcelRecode[iExcelRow,9]:=aCsvrecode[15];
      arrExcelRecode[iExcelRow,10]:='=HYPERLINK("'+sProductUrl+'","��Ʒ��ַ")';
      arrExcelRecode[iExcelRow,11]:='=HYPERLINK("'+sSavePicPath+'","ͼƬ����")';
      }
  //    excelapp.activesheet.cells[12,5].value:='=HYPERLINK("http://www.126.com","ͼ")';
  // excelapp.activesheet.cells[12,6].value:='=HYPERLINK("c:\","�ļ���")';
   result:=stmpall;
   spageinfo.Free;
   SCSVMODEINFO.Free;
   sltmp.Free;
     end;

//������Ͱ��еĲ�Ʒ��Ϣ������д��CSV�ļ�
//����SurlΪĳ���ļ������ӵ�ַ
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
  tStrCsvModeinfo:=Tstringlist.Create;  //���ڴ���csvģ���ļ�������
  tstr1:=tstringlist.Create;
  CsvLines := TStringList.Create;
  commaStr:=TStringList.Create;
  tStrCsvModeinfo.LoadFromFile(getcurrentdir+'\ModeDemo.csv');  //Ϊ���Է�������
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
//  CsvLines.Add('version 1.00');   //��һ��
//  CsvLines.Add(str1+str2+str3+str4+str5);   //�ڶ���
//  CsvLines.Add('��������,������Ŀ,������Ŀ,�¾ɳ̶�,ʡ,����,���۷�ʽ,�����۸�,�Ӽ۷���,��������,��Ч��,�˷ѳе�,ƽ��,EMS,���,��Ʊ,����,����ֿ�,�����Ƽ�,��ʼʱ��,��������,��������,�ʷ�ģ��ID,��Ա����,�޸�ʱ��,�ϴ�״̬,ͼƬ״̬,�������,��ͼƬ,��Ƶ,�����������,�û�����ID��,�û�������-ֵ��,�̼ұ���,�������Ա���,��������,����ID,����ID,��������,�˻�����,����״̬,���緢��,��Ʒ,ʳƷר��,�����,�������,������,�˻�����ŵ,�������,��������');//������
  //  Csvlines.Add(GetAlibabaProductsinfo(strurl,tStrCsvModeinfo[3],k)); //����Ӧ���ӵ���Ϣд��CSVģ���ļ���
  //GetAlibabaListinfo(surl);
  //   showmessage(inttostr(form1.sgShowTitle.RowCount));

  for  i:=1 to form1.sgShowTitle.RowCount do
    begin
    if form1.sgShowTitle.Cells[1,i]='yes'  then
     iUsedRecode :=iUsedRecode+1;                  //�õ���Ҫ�������Ч��¼
    end;
   setlength(arrExcelRecode,iUsedRecode,13);    //���ö�̬����

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
    //���Ѿ��γɵ�����д��Ԥ���޸ı���У����޸�ȷ��
    //����������arrExcelRecode�Ķ�Ӧ����д������
    iRecode:=high(arrexcelrecode);
    form1.sgProduceInfoPreview.RowCount:=iRecode+2;
  //  showmessage(inttostr(irecode));

    for I := Low(arrexcelrecode)+1 to High(arrexcelrecode)+1 do
     begin
       form1.sgProduceInfoPreview.Cells[1,i]:=arrExcelRecode[i-1,3]; //��ʾ����
       form1.sgProduceInfoPreview.Cells[2,i]:=arrExcelRecode[i-1,6]; //�����۸�
       form1.sgProduceInfoPreview.Cells[3,i]:=arrExcelRecode[i-1,7]; //���ۼ۸�
       fJprice:=strtofloat(arrExcelRecode[i-1,6]);
       fCprice:=strtofloat(arrExcelRecode[i-1,7]);
       form1.sgProduceInfoPreview.Cells[4,i]:=floattostr(fCprice-fJprice);  //�۲�
       form1.sgProduceInfoPreview.Cells[5,i]:=floattostr((fCprice-fJprice)/fJprice*100)+'%';//);  //ë����
       form1.sgProduceInfoPreview.Cells[6,i]:=arrExcelRecode[i-1,10];  //��Ʒ��ַ
     end;
       form1.sgProduceInfoPreview.Options:=form1.sgProduceInfoPreview.Options+[goEditing];

 // CsvLines.SaveToFile(getcurrentdir+'\Addproducts.csv'); // ���ɵ����Ա������CSV�ļ�
 // WirteDataToExcel(form1.leProdcutsFile.text,form1.cbSupplerName.items[form1.cbSupplerName.ItemIndex]);
  showmessage('����ص����ݽ����޸ģ�');
  }
 // CsvLines.Free;
  CsvLines.SaveToFile(getcurrentdir+'\Addproducts.csv'); // ���ɵ����Ա������CSV�ļ�

  tStrCsvModeinfo.Free;
  tstr1.Free;

end;


procedure TForm1.BitBtn1Click(Sender: TObject);
var
 mark:string;
 i:integer;
begin

 // showmessage(CreatRandomNumstr(4));
  if  cbSupplerName.ItemIndex=-1 then  begin showmessage('��ѡ��Ӧ��');exit;end;
  if  cbProductsClass1.ItemIndex=-1 then begin showmessage('��ѡ�����');exit; end;
  if  rgProductsclassSub.ItemIndex=-1 then begin showmessage('��ѡ������');exit; end;
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
 CsvLines.SaveToFile(getcurrentdir+'\Addproducts.csv'); // ���ɵ����Ա������CSV�ļ�
 showmessage('����д�����');
 CsvLines.Free;
end;



procedure TForm1.BitBtn3Click(Sender: TObject);
var
  iNum:integer;
begin
if memlisturl.Text=''  then
  begin
    showmessage('û����ַ��Ϣ��������!');
    exit;
  end;

//******************�ж��ǲ�Ʒ��ϸ��Ϣ����ĳ���Ʒ��������Ϣ***************
//�ж��ǲ�Ʒ��ϸ��Ϣ���ӵķ�ʽ�ǣ������г��֣�"detail"�Ĺؼ���,����ǵ���ֱ��ִ������CSV�ļ������򽫸����е���Ϣ��ʾ���б���


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

  arrExcelRecode[i-1,3]:=sgProduceInfoPreview.Cells[1,i]; //�������޸ĺ������д��EXECEL�ļ�������
  arrExcelRecode[i-1,6]:=sgProduceInfoPreview.Cells[2,i]; //�����۸�
  arrExcelRecode[i-1,7]:=sgProduceInfoPreview.Cells[3,i]; //���ۼ۸�

  end;
 }
 //showmessage(strsouce+'  '+);

 CsvLines.SaveToFile(getcurrentdir+'\Addproducts.csv'); // ���ɵ����Ա������CSV�ļ�

// WirteDataToExcel(form1.leProdcutsFile.text,form1.cbSupplerName.items[form1.cbSupplerName.ItemIndex]);
 CsvLines.Free;
//  tStrCsvModeinfo.Free;
 // tstr1.Free;

   showmessage('����д�����');


//sgProduceInfoPreview.Options:=sgProduceInfoPreview.Options+[goEditing];
{ setlength(arrM,5,12);
  for I := 0 to 4 do
    for j := 0 to 11 do
      arrm[i,j]:=inttostr(i)+inttostr(j);
WirteDataToExcel('h:\pro.xls','DELPHI1',arrM);}
//iniinfo:=Tinifile.Create(getcurrentdir+'\baseinfo.ini');
//showmessage(iniinfo.ReadString ('PATH','ProPicPath',''));
//iniinfo.Free;
// S1:='<a href="http://ambopower.cn.1688.com/page/contactinfo.htm">��ϵ��ʽ</a>';
// S2:='"';
// S3:='">��˾����';
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

   odfilebox.Filter:='�Ա�����ģ���ļ�(*.csv)|*.csv';
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
      ExcelApp.WorkSheets['��������Ϣ'].Activate;
      rowlast:=excelapp.Cells.SpecialCells(xlCellTypelastCell, EmptyParam).row;
    //  showmessage(inttostr(rowlast));
       for I := 3 to rowlast do
        begin
          if ExcelApp.cells[i,3].value=vleSupplerInfo.Values['��˾��ƣ�'] then
           begin
             showmessage('�ù�Ӧ����Ϣ�Ѿ�����!');
             cbsupplername.Text:=vleSupplerInfo.Values['��˾��ƣ�'];
             ExcelApp.WorkBooks.Close;
             ExcelApp.Quit;
             exit;
           end;

        end;

       ExcelApp.ActiveSheet.Rows[rowlast+1].Insert;// �����һ�����һ����
       ExcelApp.Cells[rowlast+1,1].Value:= inttostr(strtoint(ExcelApp.Cells[rowlast,1].Value)+1);
       ExcelApp.Cells[rowlast+1,2].Value:= vleSupplerInfo.Values['��˾���ƣ�'];
       ExcelApp.Cells[rowlast+1,3].Value:= vleSupplerInfo.Values['��˾��ƣ�'];
       ExcelApp.Cells[rowlast+1,4].Value:= inttostr(strtoint(ExcelApp.Cells[rowlast,1].Value));
       ExcelApp.Cells[rowlast+1,5].Value:= vleSupplerInfo.Values['��ַ��'];
       ExcelApp.Cells[rowlast+1,6].Value:= vleSupplerInfo.Values['�绰��'];
  //     ExcelApp.Cells[rowlast+1,7].Value:= vleSupplerInfo.Values['���棺'];
       ExcelApp.Cells[rowlast+1,7].Value:= vleSupplerInfo.Values['��ϵ�ˣ�'];
       ExcelApp.Cells[rowlast+1,8].Value:= vleSupplerInfo.Values['�ƶ��绰��'];
       ExcelApp.Cells[rowlast+1,9].Value:='=HYPERLINK("'+vleSupplerInfo.Values['��Ӧ��Ʒ��']+'","��Ӧ��Ʒ")';
   //    ExcelApp.Cells[rowlast+1,11].Value:='=HYPERLINK("'+vleSupplerInfo.Values['��˾���ã�']+'","��˾����")';

      ExcelApp.Activeworkbook.save;
      ExcelApp.WorkBooks.Close;
      ExcelApp.Quit;
      cbSupplerName.items.add(vleSupplerInfo.Values['��˾��ƣ�']);
      cbsupplername.Text:=vleSupplerInfo.Values['��˾��ƣ�'];
     end;
end;

procedure TForm1.btnPicPathClick(Sender: TObject);
var
  NewDir: string;
  iniinfo:TINIFILE;
begin
   iniinfo:=Tinifile.Create(getcurrentdir+'\baseinfo.ini');

  if SelectFolderDialog(Handle, 'ѡ������ͼƬ����ļ���', '', NewDir) then
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

   odfilebox.Filter:='Excel�ļ�(*.xls)|*.xls';
  if odfilebox.Execute then
    leProdcutsfile.Text:=odfilebox.FileName;

    ExcelApp := CreateOleObject( 'Excel.Application' );
    ExcelApp.WorkBooks.Open( leProdcutsfile.text );
    ExcelApp.WorkSheets[ '��������Ϣ' ].Activate;
    rowlast:=excelapp.Cells.SpecialCells(xlCellTypelastCell, EmptyParam).row;  //��ȡEXECEL�����һ��
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
   JSONObject: TJSONObject; // JSON��
            i: Integer; // ѭ������
         temp: string; // ��ʱʹ�ñ���
    jsonArray: TJSONArray; // JSON�������
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
     JSONObject.AddPair(tName[i],tData[i]);  // ��ģ���е��������� ģ��ֵJSON��
     end;
   showmessage(inttostr(JSONObject.Count));
 //  tp:=jsonobject.Get('��������');

   tp:=TJSONPair.Create('��������','asfasfasfsafasf');

//   JSONObject.SetPairs(tp);
//   JSONOBJECT.Values('��������').Value:='asfdafda';
   showmessage(JSONObject.GetValue('��������').ToString);

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
   vleSupplerInfo.Values['��˾��ƣ�']:=edtsupplyname.Text;
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
