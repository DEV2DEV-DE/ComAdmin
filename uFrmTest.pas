unit uFrmTest;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls;

type
  TFrmTest = class(TForm)
    btnComAdmin: TButton;
    tvTree: TTreeView;
    cmbServer: TComboBox;
    memoLog: TMemo;
    procedure btnComAdminClick(Sender: TObject);
  private
    procedure OnReadCOMObject(const AObjectType, AObjectName: string);
  public
    { Public-Deklarationen }
  end;

var
  FrmTest: TFrmTest;

implementation

uses
  uComAdmin;

{$R *.dfm}

procedure TFrmTest.btnComAdminClick(Sender: TObject);
var
  LCatalog: TComAdminCatalog;
  i, j: Integer;
  AppNode, RoleNode: TTreeNode;
  k: Integer;
begin
  tvTree.Items.Clear;
  memoLog.Clear;
  LCatalog := TComAdminCatalog.Create(cmbServer.Text, '', OnReadCOMObject);
  try
    //LCatalog.ExportApplication(20, 'D:\Test123.msi');
    for i := 0 to LCatalog.Applications.Count - 1 do
    begin
      AppNode := tvTree.Items.AddChild(nil, LCatalog.Applications[i].Name);
      for j := 0 to LCatalog.Applications[i].Roles.Count - 1 do
      begin
        RoleNode := tvTree.Items.AddChild(AppNode, LCatalog.Applications[i].Roles[j].Name);
        for k := 0 to LCatalog.Applications[i].Roles[j].Users.Count - 1 do
          tvTree.Items.AddChild(RoleNode, LCatalog.Applications[i].Roles[j].Users[k].Name);
      end;
    end;
    tvTree.AlphaSort;
  finally
    LCatalog.Free;
  end;
end;

procedure TFrmTest.OnReadCOMObject(const AObjectType, AObjectName: string);
begin
  memoLog.Lines.AddPair(AObjectType, AObjectName);
  memoLog.ScrollBy(0, 1)
end;

end.
