program ComAdminTest;

uses
  Vcl.Forms,
  uFrmTest in 'uFrmTest.pas' {FrmTest},
  uComAdmin in 'uComAdmin.pas';

{$R *.res}

begin
  ReportMemoryLeaksOnShutdown := DebugHook <> 0;
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFrmTest, FrmTest);
  Application.Run;
end.
