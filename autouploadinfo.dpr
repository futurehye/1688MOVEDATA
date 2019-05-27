program autouploadinfo;

uses
  Vcl.Forms,
  uMain in 'uMain.pas' {Form1},
  uPublicFun in 'uPublicFun.pas',
  uHelpInfo in 'uHelpInfo.pas',
  uTaoBaoInfoProUnit in 'uTaoBaoInfoProUnit.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
