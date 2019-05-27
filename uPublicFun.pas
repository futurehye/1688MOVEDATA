unit uPublicFun;
 //����Ԫ��Ҫ��������Ҫ�Ĺ��������͹���

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
     function DeleteDirectory(NowPath: string): Boolean; // ɾ������Ŀ¼(��ɾ�ļ��У�

     function SetProductsInfo(strcsvmodeinfo:string;iExcelRow:integer):string;
     function CompressImageFile(FileName: string;  Width, Height: integer; PressQuality:Integer= 90): Boolean;
     var
     arrExcelRecode:array of array of string;
     arrProdctusInfo:array[1..11] of string;  //��ŴӲ�Ʒҳ��õ�����Ϣ���������⡢�۸��

implementation



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

end;

//�����д�Сд�ַ�������ַ���
//���� strLong:integer Ϊ���ɵĴ�����
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

 //�����ִ���һ��ָ�����ȵ�����ַ���
 //����strLong����ȷ���ַ����ĳ���
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

//����SplitString���ڴ�Դ�ַ�����ͨ���ָ��ַ���ȡ�ֶε��ַ�������Ϊ�ַ��б���
//���磺SplitString('abc,bzcde,efg',','),�򷵻�abc  bzcde  efg��������
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
//����ļ������Ƿ��зǷ��ַ�����������ַ��滻Ϊ'&'�������޸ĺ���ļ���
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

 //���ڴӶԻ����л�ȡ�ļ���·����
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


//���ַ����л��ݻ�ñ�ʶ�ַ���֮����ַ�������STRBACKMARKΪ��������strSOUCE��ΪΨһ��
//������strSourceΪԴ����sBefoureMarkΪǰ���ʶ��(һ��Ϊһ���ַ�����STRBACKMARKΪ���ʶ����
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
            ExcelApp.WorkSheets[sSheetname].Activate; //����һ�����Sheet
            issheet:=1;
            break;
          end;
   end;
  if issheet=0 then      //����һ��SHEET
     begin
       ExcelApp.WorkSheets.Add;
       Excelapp.workbooks[1].sheets['sheet1'].name:=sSheetName;
      //  sheet:=Excelapp.workbooks[1].sheets['test'];
       ExcelApp.WorkSheets[sSheetname].Activate; //����һ�����Sheet
       //��ӱ�ͷ
       ExcelApp.ActiveSheet.cells[2,1]:='���';
       ExcelApp.ActiveSheet.cells[2,2]:='��Ʒ����';
       ExcelApp.ActiveSheet.cells[2,3]:='��Ʒ����';
       ExcelApp.ActiveSheet.cells[2,4]:='��������';
       ExcelApp.ActiveSheet.cells[2,5]:='Ʒ��';
       ExcelApp.ActiveSheet.cells[2,6]:='�ͺ�';
       ExcelApp.ActiveSheet.cells[2,7]:='�����۸�';
       ExcelApp.ActiveSheet.cells[2,8]:='���ۼ۸�';
       ExcelApp.ActiveSheet.cells[2,9]:='�Ƿ��л�';
       ExcelApp.ActiveSheet.cells[2,10]:='������ݷ�';
       ExcelApp.ActiveSheet.cells[2,11]:='������ַ';
       ExcelApp.ActiveSheet.cells[2,12]:='��ƷͼƬ';
       ExcelApp.ActiveSheet.cells[2,13]:='������ַ';

//    excelapp.activesheet.cells[12,5].value:='=HYPERLINK("http://www.126.com","ͼ")';
//    excelapp.activesheet.cells[12,6].value:='=HYPERLINK("c:\","�ļ���")';

      end;
         //ExcelApp.WorkSheets[SheetName].Delete;   //ɾ��
      //����Ʒ��Ϣд���Ʒ��ϢEXCEL�ĵ���
         rowlast :=excelapp.Cells.SpecialCells(xlCellTypelastCell, EmptyParam).row;  //��ȡEXECEL�����һ��
         for m:= low(arrExcelRecode) to high(arrExcelRecode) do
          for n := low(arrExcelRecode[m]) to high(arrExcelRecode[m]) do
             begin
              ExcelApp.ActiveSheet.cells[rowlast+m+1,n+1]:=arrExcelRecode[m,n];
             end;

     //����Ӧ����Ϣд���Ʒ��Ϣexcel�ĵ���

 ExcelApp.Activeworkbook.save;
 ExcelApp.WorkBooks.Close;
 ExcelApp.Quit;

end;

//�����ַ���ȷ�����������ļ���ϵͳ
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

 //����SetProductsinfo���ݵõ��Ĳ�Ʒ��Ϣҳ����Ϣ��arrProdctusInfo���飩������Ϣ���롰�Ա�����CSV�ĵ��ֶ�
//����˵����sProductUrlΪ��Ʒ��ַ�����ַ���
function SetProductsInfo(strcsvmodeinfo:string;iExcelRow:integer):string;
 var
   stmpall,stmp1,stmp2,stmp3,stmp4,stmp5,sUrl,sTitle,sPrice,sSavePicPath,sSavePriwePicPath,sProCode:string;
   aCsvrecode: array[1..50] of string;
   arrProjinxiaoInfo:array[1..11] of string;  //���ڴ����Ҫ��������ĵ�����Ϣ������0λ��Ź����̣�
   i,iItems:integer;
   sPageinfo,sCSVMODEINFO,slTmp:Tstringlist;

 begin
   sltmp:=tstringlist.Create;
   spageinfo:=Tstringlist.Create;
   SCSVMODEINFO:=TSTRINGLIST.Create;
   SCSVMODEINFO.CommaText:=strcsvmodeinfo;
//   sproducturl:= form1.IdHttpListPage.URL.URLEncode(sproducturl);
//   sPageinfo.Text :=form1.IdHttpListPage.Get(sProducturl); //��ʼ��ַ�õ��Ĳ�Ʒҳ��Դ��
   //  showmessage(spageinfo.Text);

     aCsvrecode[1]:=arrProdctusInfo[1];  //�õ�����'
     aCsvrecode[2]:=sCsvModeInfo[1];//form1.vleBaseProduceInfo.Values['������Ŀ'];   //������Ŀ
     aCsvrecode[3]:=sCsvModeInfo[2];   //������Ŀ
     if sCsvModeInfo[3]<>'' then aCsvrecode[4]:=sCsvModeInfo[3] else aCsvrecode[4]:='1';   //�¾ɳ̶�
     if sCsvModeInfo[4]<>'' then
      begin
      aCsvrecode[5]:=sCsvModeInfo[4];   //����ʡ��
      aCsvrecode[6]:=sCsvModeInfo[5];  //�������
       end;

     if sCsvModeInfo[6]<>'' then aCsvrecode[7]:=sCsvModeInfo[6] else aCsvrecode[7]:='1';   //�¾ɳ̶�
//     aCsvrecode[7]:=form1.vleBaseProduceInfo.Values['���۷�ʽ'];   //���۷�ʽ

   //  arrExcelRecode[iExcelRow,6]:=arrProdctusInfo[2]; //�õ������۸�
     aCsvrecode[8]:=arrProdctusInfo[3];  //����µļ۸��ۼۣ�

   //  arrExcelRecode[iExcelRow,7]:=aCsvrecode[8];   //���ۼ۸�
     aCsvrecode[9]:='';   //�Ӽ۷���
     if sCsvModeInfo[9]<>'' then aCsvrecode[10]:=sCsvModeInfo[9] else aCsvrecode[10]:='100';   //��Ʒ����

     aCsvrecode[11]:=sCsvModeInfo[10];//form1.vleBaseProduceInfo.Values['��Ч��'];;   //��Ч��
     aCsvrecode[12]:=sCsvModeInfo[11]; //    form1.cbSelPostage.Text;   //�˷ѳе�
     aCsvrecode[13]:=sCsvModeInfo[12]; ///form1.lePostnom.Text;   //ƽ��
     aCsvrecode[14]:=sCsvModeInfo[13];  //form1.lePostems.Text;   //EMS
     aCsvrecode[15]:=sCsvModeInfo[14];  //form1.lePostems.Text;   //���S
     aCsvrecode[16]:=sCsvModeInfo[15]; //    form1.cbSelPostage.Text;   //�˷ѳе�
     aCsvrecode[17]:=sCsvModeInfo[16]; ///form1.lePostnom.Text;   //ƽ��
     aCsvrecode[18]:=sCsvModeInfo[17];  //form1.lePostems.Text;   //EMS
     aCsvrecode[19]:=sCsvModeInfo[18];  //form1.lePostems.Text;   //���S
     aCsvrecode[20]:=datetostr(now)+timetostr(now);//form1.vleBaseProduceInfo.Values['��ʼʱ��'];   //��ʼʱ��

     aCsvrecode[21]:=arrProdctusInfo[4];    //���뱦������

//     showmessage('222');
     aCsvrecode[22]:=sCsvModeInfo[21];   //��������
     aCsvrecode[23]:=sCsvModeInfo[22];   //�ʷ�ģ��ID
     aCsvrecode[24]:=sCsvModeInfo[23];   //��Ա����
     aCsvrecode[25]:=sCsvModeInfo[24];   //�޸�ʱ��
     aCsvrecode[26]:=sCsvModeInfo[25];   // �ϴ�״̬
     aCsvrecode[27]:=sCsvModeInfo[26];   // ͼƬ״̬
     aCsvrecode[28]:=sCsvModeInfo[27];   // �������

     aCsvrecode[29]:=arrProdctusInfo[5];   //��ͼƬ

     aCsvrecode[30]:=sCsvModeInfo[29];   // ��Ƶ
     aCsvrecode[31]:=sCsvModeInfo[30];   // �����������
     aCsvrecode[32]:=sCsvModeInfo[31];   // �û�����ID��
     aCsvrecode[33]:=sCsvModeInfo[32];   // �û�������-ֵ��

     aCsvrecode[34]:=arrProdctusInfo[6];   // �̼ұ���


     for I := 35 to 43 do
      aCsvrecode[i]:=sCsvModeInfo[i-1];   {�ֱ�����������Ա�������������\����ID����ID
	�������ࡢ�˻����ơ�����״̬�����緢������Ʒ}
     aCsvrecode[44]:='';   //	ʳƷר��
     aCsvrecode[45]:='';   //	�����
     aCsvrecode[46]:='';   //	�������
     aCsvrecode[47]:='';   //	������
     aCsvrecode[48]:='';  // else aCsvrecode[48]:='0';   //�˻�����ŵ
     aCsvrecode[49]:='';   //	�������
     aCsvrecode[50]:='';   //	��������

     for i := 1 to 50 do  //ȡ�����еĶ��ű��
        stmpall:=stmpall+StringReplace(aCsvrecode[i],',','',[rfReplaceAll])+',';

      //����Ϣ����������¼��
  //    arrExcelRecode[iExcelRow,0]:=form1.cbSupplerName.items[form1.cbSupplerName.ItemIndex]; //����������
 {     arrExcelRecode[iExcelRow,1]:=arrProdctusInfo[6];   //��Ʒ���
      arrExcelRecode[iExcelRow,2]:=arrProdctusInfo[11]; //��Ʒ����
      arrExcelRecode[iExcelRow,3]:=arrProdctusInfo[1];
      if (length(arrProdctusInfo[7])>30) or (length(arrProdctusInfo[7])<1) then  arrExcelRecode[iExcelRow,4]:='' else  arrExcelRecode[iExcelRow,4]:=arrProdctusInfo[7];
      stmp5:=arrProdctusInfo[8];
      if (length(stmp5)>30) or (length(stmp5)<1) then  arrExcelRecode[iExcelRow,5]:='' else  arrExcelRecode[iExcelRow,5]:=stmp5;

      arrExcelRecode[iExcelRow,8]:='��';
      arrExcelRecode[iExcelRow,9]:=aCsvrecode[15];
      arrExcelRecode[iExcelRow,10]:=arrProdctusInfo[9];
      arrExcelRecode[iExcelRow,11]:=arrProdctusInfo[10];
  }
  //    excelapp.activesheet.cells[12,5].value:='=HYPERLINK("http://www.126.com","ͼ")';
  // excelapp.activesheet.cells[12,6].value:='=HYPERLINK("c:\","�ļ���")';

   result:=stmpall;
   spageinfo.Free;
   SCSVMODEINFO.Free;
   sltmp.Free;
  end;

  function DeleteDirectory(NowPath: string): Boolean; // ɾ������Ŀ¼(��ɾ�ļ��У�
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
  //removedir(NowPath); �����Ҫɾ���ļ��������
  result := True;
end;


/// <summary>
/// ѹ��ͼƬ(BMP��JPG��PNG)
/// </summary>
/// <param name="FileName">�ļ�·��</param>
/// <param name="Width">��Ҫѹ����Ŀ��</param>
/// <param name="Height">��Ҫѹ����ĸ߶�</param>
/// <param name="PressQuality">ѹ������</param>
/// <returns>�Ƿ�ѹ���ɹ�</returns>

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
          begin //��������С,���߶�����߶ȵ�
             RetHeight := NewHeight;
             RetWidth := Round(OldWidth *  (NewHeight/OldHeight));
          end
          else
          begin //��������С,����������ȵ�
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
