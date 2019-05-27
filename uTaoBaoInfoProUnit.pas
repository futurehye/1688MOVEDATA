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

 //����GetTaobaoProductsinfo���ݽ��ܲ�Ʒ��ҳ�õ��Ա���Ʒ����Ҫ��Ϣ�������ֵ�arrProdctusInfo������
//����˵����sProductUrlΪ��Ʒ��ַ�����ַ���
//��ȡ��Ϣԭ��
//���ڻ�ȡ����Ʒ�������ķ�������һ�������ݲ�Ʒ��ַ�õ���ҳԴ��
//                        �ڶ�������Դ����������g_config.dynamicScript("�����ַ��������������Եõ�����Ҳ����ַ
//                        �����������ݵõ���������ַ����ȡ��Դ���룬���������еġ�var desc='���͡�';���ַ���ȥ��

function GettaobaoProductsPageinfo(sProductUrl:string):string;
 var
   strFilePath,stmp2,sUrl,sTitle,sPrice,sSavePicPath,sSavePriwePicPath,sProCode:string;
//   aCsvrecode: array[1..50] of string;
//   arrProjinxiaoInfo:array[1..11] of string;  //���ڴ����Ҫ��������ĵ�����Ϣ������0λ��Ź����̣�
   i:integer;
   sPageinfo,sCSVMODEINFO:Tstringlist;
//   req: IXMLHttpRequest;
 begin
 for i := 1 to 11 do
    arrProdctusInfo[i]:='';

    spageinfo:=Tstringlist.Create;
    sproducturl:= form1.IdHttpListPage.URL.URLEncode(sproducturl);
    sPageinfo.Text :=form1.IdHttpListPage.Get(sProducturl); //��ʼ��ַ�õ��Ĳ�Ʒҳ��Դ��
   //  showmessage(spageinfo.Text);

     arrProdctusInfo[1]:=searchstring(spageinfo.Text,'<title>','-�Ա���</title>');  //�õ�����'
   //  arrProdctusInfo[2]:=searchstring(spageinfo.Text,'<strong id="J_StrPrice" ><em class="tb-rmb">&yen;</em><em class="tb-rmb-num">','</em></strong>'); //�õ������۸�
     arrProdctusInfo[2]:='0';
     arrProdctusInfo[3]:=inttostr(ceil(strtofloat(arrProdctusInfo[2])*strtofloat('1.'+form1.vleBaseProduceInfo.Values['�����۸�'])));  //����õ��µļ۸�

     stmp2:=searchstring(spageinfo.Text,'g_config.dynamicScript("','")');  //�õ����в�Ʒ���������ӵ�ַ
     if stmp2<>'' then
       begin

         strFilePath:=ModifyFileNameString(arrProdctusInfo[1]);
         if form1.lePicPath.text='' then
         sSavePicPath:=getcurrentdir+'\'+form1.cbSupplerName.Text+'\'+strFilePath+'\'  //�ڵ�ǰĿ¼�½���һ���Ա���Ϊ�����ļ������ڴ�Ų�Ʒ�����е�ͼƬ�ļ�
         else
         sSavePicPath:=form1.lePicPath.Text+form1.cbSupplerName.Text+'\'+strFilePath;  //�ڵ�ǰĿ¼�½���һ���Ա���Ϊ�����ļ������ڴ�Ų�Ʒ�����е�ͼƬ�ļ�
  //      showmessage(sSavepicpath);
         CreatMkDir(sSavePicPath);
    // if not DirectoryExists(sSavePicPath) then  MKDIR(sSavePicPath);
          //  showmessage('sSavepicpath');

     arrProdctusInfo[4]:=ProtaobaoProductShowInfo(stmp2,sSavePicPath);    //���뱦��������Ϣ�������Ѿ���ͼƬ·�������˴���
     end
     else
     arrProdctusInfo[4]:='';

     arrProdctusInfo[5]:=protaobaoProductPageRviewPic(sPageinfo.Text,'Addproducts');   //��ͼƬ
    //**************
    // sProCode:=GetProductsCode(' ');
    // if sprocode<>'' then   arrProdctusInfo[6]:=sprocode;   // �̼ұ���
     arrProdctusInfo[6]:='';
    //********************

     arrProdctusInfo[7]:='';//searchstring(spageinfo.Text,'Ʒ��:&nbsp;','</li>');
     arrProdctusInfo[8]:='';//searchstring(spageinfo.Text,'�ͺ�:&nbsp;','</li>');
     arrProdctusInfo[9]:='';//'=HYPERLINK("'+sProductUrl+'","��Ʒ��ַ")';
     arrProdctusInfo[10]:='';//'=HYPERLINK("'+sSavePicPath+'","ͼƬ����")';
   //   arrProdctusInfo[11]:=form1.cbProductsClass1.items[form1.cbProductsClass1.ItemIndex]+' '+form1.rgProductsclassSub.items[form1.rgProductsclassSub.ItemIndex]; //��Ʒ����

     arrProdctusInfo[11]:='';


     spageinfo.Free;
     result:='';

 //  SCSVMODEINFO.Free;
 //  sltmp.Free;
  end;


   //�õ�taobao��Ʒ�б�ҳ�еķ�ҳ���ӵ�ַ
//��������˵��������������һ����������������ҳ������һҳ������1����û��ҳ����0
//����з�ҳ��ͬʱ��ͨ������aPageUrlInfo������ÿ����ҳ����Ϣ
function GettaobaoListpageUrl(sSourstring:string):boolean;
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


//����GettaobaoListinfo�õ��Ա���Ʒ�б�ҳ�еĲ�Ʒ
//����˵����sPageUrlΪ��ַ���Ӵ�
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
      //�ҳ���Ʒ�б��еĴ�Χ��Ʒ�����б�
   //     stmpall:=searchstringsavemark(Spageinfo.text,'<ul class="offer-list-row">','</ul>');

       while pos('<dd class="detail">',stmpall)>0 do       //ĩβ��ǻ����ַ����о�ѭ��ִ��
       begin
       stmp1:=searchstringsavemark(stmpall,'<dd class="detail">','</dd>');
     //  showmessage(stmp1);
       sUrl:=searchstring(stmp1,'href="','" target="_blank">');
       sTitle:=searchstring(stmp1,'" target="_blank">','</a>');
       sPrice:=searchstring(stmp1,'class="c-price">','</span>');
       ipos1:=pos('</dl>',stmpall);
       stmpall:=copy(stmpall,ipos1+5,length(stmpall)-ipos1+5);  //�����5������</dl>���ֽ���

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
//����ProtaobaoProductShowInfo(sprourl,sPicPath:string)�����Ʒҳ��ͼƬ��Ϣ
//����spropageinfoΪ��Ʒ��ϢҳԴ����Ԥ�����Ĵ����루��Ҫ��ȥ����'\')
//��Ҫ����ԭ����Դ�����漰��ͼƬ���ص�ָ��Ŀ¼�У����滻Դ���е�ͼƬ·��
function  ProTaobaoProductShowInfo(sprourl,sPicPath:string):string;
 var
   sTmp,sDownPic,sDownPicFilename,spicfile,spropageinfo,sNewPageInfo:string;
   iPos1,inum,i:integer;
   arrPicFilepath:array of array of string;
   req: IXMLHttpRequest;
begin

   inum:=0;
   setlength(arrPicFilepath,100,2);  //����һ����ά����

   req := CoXMLHTTP.Create;
        req.open('Get',sprourl, False, EmptyParam, EmptyParam);
        req.send(EmptyParam);
        spropageinfo:=req.responseText;  //�õ�����ҳԴ����

 //  spropageinfo:=form1.IdHttpListPage.Get(sprourl);
   spropageinfo:=StringReplace (spropageinfo, '\', '', [rfReplaceAll]);//���ı��е�����'\'��''�滻
   spropageinfo:=StringReplace (spropageinfo, #13, '', [rfReplaceAll]);//���ı��е�����'\'��''�滻
   spropageinfo:=StringReplace (spropageinfo, #10, '', [rfReplaceAll]);//���ı��е�����'\'��''�滻
 //  showmessage('�ַ������ȣ�'+inttostr(length(spropageinfo)));
   sTmp:=sPropageinfo;

 //  spropageinfo:=searchstring(spropageinfo,'var offer_details={"content":"','"};');
 //  form1.Memo1.Text:=spropageinfo;
   while pos('src="',sTmp)>0 do
   begin
   inum:=inum+1;
//     showmessage(inttostr(pos(sTmp,'src="')));
     sDownPic:=searchstring(sTmp,'src="','"');
     ipos1:=pos('<img',sTmp);
//     snewtmp:=copy(stmp,1,ipos1);          //�õ�
     stmp:=copy(stmp,ipos1+4,length(stmp)-ipos1);
     spicfile:=StrRScan(pchar(sDownPic),'/');
   //  showmessage(sPicPath);
     if copy(sPicPath,length(sPicPath)-1,1)<>'\' then   sPicPath:=sPicPath+'\';

     sDownPicFilename:=sPicPath+copy(spicfile,2,length(spicfile)-1);
   //  showmessage(sdownpicfilename);
     UrlDownloadToFile(nil,pchar(sdownPic),pchar(sDownPicFilename), 0, nil);//��ͼƬ�ļ����ص�ָ��Ŀ¼
     arrPicFilepath[inum-1,0]:=sDownPic ; //��ԭԴ���е�ͼƬ·������������
     arrPicFilepath[inum-1,1]:='File:///'+sDownPicfilename;  //�����뱾��ͼƬ��·������������
    // showmessage(stmp);
   end;
   //��ԭԴ���е�����ͼƬ·��ȫ����Ϊ����ͼƬ·��
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

//����ProTaobaoProductPageRviewPic��Ҫ���ڴ����Ա���Ʒҳ�е�Ԥ��ͼƬ��
//һ�ǽ�ͼƬ��ȡ���ŵ���CSVͬ����Ŀ¼��savepath���У�
//�����������Ա������е���ͼƬ��ʽ�ַ��������ļ���:1:n;����NΪԤ����ʾ�еľ���λ��
//����str1Ϊ��Ʒҳ��Ϣ��Դ���룬savepathΪ��csv�ļ�ͬ�����ļ���
function  ProtaobaoProductPageRviewPic(str1,savepath:string):string;
 var
  sTmp,sdownpic,sCsvNewPicString,s1,sSavePriwePicPath,sDownPicFilename,sPicSize:string;
  ipos1,iitem:integer;
 begin
  sTmp:=str1;
  iitem:=0;
  sSavePriwePicPath:=getcurrentdir+'\'+savepath;  //�ڵ�ǰĿ¼�½���һ����csvͬ�����ļ������ڴ�Ų�Ʒ���ͼƬ�ļ�
  if not DirectoryExists(sSavePriwePicPath) then  MKDIR(sSavePriwePicPath);

  //sPicSize:=searchstring(sTmp,'<div class="tb-booth tb-pic tb-s','">');   //�õ�Ԥ����ͼ�ı߳��ߴ�
  //sPicSize:='.jpg_'+sPicSize+'x'+sPicSize+'.jpg';  //�γ�Ԥ��ͼƬ��־
  sPicSize:='.jpg_400x400.jpg';  //�γ�Ԥ��ͼƬ��־
  stmp:=searchstring(stmp,'<ul id="J_UlThumb" class="tb-thumb tb-clearfix">','</ul>');

  while pos('src="',sTmp)>0 do
   begin

     sDownPic:=searchstring(sTmp,'src="','.jpg_');

     sDownPic:=sDownPic+sPicSize;   //�����ص�Ԥ��ͼƬ�����ļ�ԭ��+�ߴ�x�ߴ�.JPG)
 //    form1.Memo1.Lines.Add('pic1:'+sDownpic);
     ipos1:=pos('src="',sTmp);
//     snewtmp:=copy(stmp,1,ipos1);          //�õ�
     stmp:=copy(stmp,ipos1+5,length(stmp)-ipos1);

 //    spicfile:=StrRScan(pchar(sDownPic),'/');
     s1:= CreatRandomstr(25);
     sDownPicFilename:=savePath+'\'+s1+'.tbi';
     UrlDownloadToFile(nil,pchar(sdownPic),pchar(sDownPicFilename), 0, nil);//��ͼƬ�ļ����ص�ָ��Ŀ¼

     sCsvNewPicString:=sCsvNewPicString+s1+':1:'+inttostr(iitem)+':|;';
     iitem:=iitem+1;
     s1:='';
   end;

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



//���Ա��еĲ�Ʒ��Ϣ������д��CSV�ļ�
//����SurlΪĳ���ļ������ӵ�ַ
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
  tStrCsvModeinfo:=Tstringlist.Create;  //���ڴ���csvģ���ļ�������
  tstr1:=tstringlist.Create;
  CsvLines := TStringList.Create;

  if form1.leModeFile.text<>'' then
  tStrCsvModeinfo.LoadFromFile(form1.leModeFile.text)
  else
  begin
  showmessage('������ģ���ļ���');
  exit;
  end;       //����CSV�ļ�ģ��

 { if (form1.vleBaseProduceInfo.Values['�����۸�']='') or (form1.vleBaseProduceInfo.Values['�����۸�']='0')   then
    begin
      showmessage('�����뱦���Ӽ۱�����');
      exit;
    end;}

  CsvLines.Add('version 1.00');   //��һ��
  CsvLines.Add(str1+str2+str3+str4+str5);   //�ڶ���
  CsvLines.Add('��������,������Ŀ,������Ŀ,�¾ɳ̶�,ʡ,����,���۷�ʽ,�����۸�,�Ӽ۷���,��������,��Ч��,�˷ѳе�,ƽ��,EMS,���,��Ʊ,����,����ֿ�,�����Ƽ�,��ʼʱ��,��������,��������,�ʷ�ģ��ID,��Ա����,�޸�ʱ��,�ϴ�״̬,ͼƬ״̬,�������,��ͼƬ,��Ƶ,�����������,�û�����ID��,�û�������-ֵ��,�̼ұ���,�������Ա���,��������,����ID,����ID,��������,�˻�����,����״̬,���緢��,��Ʒ,ʳƷר��,�����,�������,������,�˻�����ŵ,�������,��������');//������

  //GetAlibabaListinfo(surl);
{  for  i:=1 to form1.sgShowTitle.RowCount do
    begin
    if form1.sgShowTitle.Cells[1,i]='yes'  then
     iUsedRecode :=iUsedRecode+1;                  //�õ���Ҫ�������Ч��¼
    end;
 }
    //setlength(arrExcelRecode,iUsedRecode,13);    //���ö�̬����

    GettaobaoProductsPageinfo(sUrl);   //�õ��ƶ���ҳ���Ա�ҳ����Ϣ��������arrProdctusInfo������


    Csvlines.Add(SetProductsinfo(tStrCsvModeinfo[3],k));


  CsvLines.SaveToFile(getcurrentdir+'\Addproducts.csv'); // ���ɵ����Ա������CSV�ļ�
 // WirteDataToExcel(form1.leProdcutsFile.text,form1.cbSupplerName.items[form1.cbSupplerName.ItemIndex]);
 // showmessage('����ص����ݽ����޸ģ�');
  CsvLines.Free;
  tStrCsvModeinfo.Free;
  tstr1.Free;

end;

 //����SetTaoBaoProductsinfo���ݵõ��Ĳ�Ʒ��Ϣҳ����Ϣ��arrProdctusInfo���飩������Ϣ���롰�Ա�����CSV�ĵ��ֶ�
//����˵����sProductUrlΪ��Ʒ��ַ�����ַ���
function SetTaoBaoProductsInfo(strcsvmodeinfo:tstrings):string;
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
   showmessage(strcsvmodeinfo.CommaText);
   SCSVMODEINFO.CommaText:=strcsvmodeinfo.CommaText;
//   sproducturl:= form1.IdHttpListPage.URL.URLEncode(sproducturl);
//   sPageinfo.Text :=form1.IdHttpListPage.Get(sProducturl); //��ʼ��ַ�õ��Ĳ�Ʒҳ��Դ��
   //  showmessage(spageinfo.Text);
   showmessage('DDTT');
     aCsvrecode[1]:=arrProdctusInfo[1];  //�õ�����'
     aCsvrecode[2]:=sCsvModeInfo[1];//form1.vleBaseProduceInfo.Values['������Ŀ'];   //������Ŀ
     aCsvrecode[3]:=sCsvModeInfo[2];   //������Ŀ
     if sCsvModeInfo[3]<>'' then aCsvrecode[4]:=sCsvModeInfo[3] else aCsvrecode[4]:='1';   //�¾ɳ̶�
     if sCsvModeInfo[4]<>'' then
      begin
      aCsvrecode[5]:=sCsvModeInfo[4];   //����ʡ��
      aCsvrecode[6]:=sCsvModeInfo[5];  //�������
       end;
      showmessage('AAAA');
     if sCsvModeInfo[6]<>'' then aCsvrecode[7]:=sCsvModeInfo[6] else aCsvrecode[7]:='1';   //�¾ɳ̶�
//     aCsvrecode[7]:=form1.vleBaseProduceInfo.Values['���۷�ʽ'];   //���۷�ʽ

  //   arrExcelRecode[iExcelRow,6]:=arrProdctusInfo[2]; //�õ������۸�
     aCsvrecode[8]:=arrProdctusInfo[3];  //����µļ۸��ۼۣ�

   //  arrExcelRecode[iExcelRow,7]:=aCsvrecode[8];   //���ۼ۸�
     aCsvrecode[9]:='';   //�Ӽ۷���
     if sCsvModeInfo[9]<>'' then aCsvrecode[10]:=sCsvModeInfo[9] else aCsvrecode[10]:='100';   //��Ʒ����
           showmessage('bbb');
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

end.
