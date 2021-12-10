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
    procedure btnComAdminClick(Sender: TObject);
  private
    { Private-Deklarationen }
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
  LCatalog := TComAdminCatalog.Create(cmbServer.Text);
  try
    //LCatalog.ExportApplication(20, 'D:\Test123.msi');
    LCatalog.Filter := '*';
    for i := 0 to LCatalog.Applications.Count - 1 do
    begin
      AppNode := tvTree.Items.AddChild(nil, LCatalog.Applications[i].Name);
      for j := 0 to LCatalog.Applications[i].Roles.Count - 1 do
      begin
        RoleNode := tvTree.Items.AddChild(AppNode, LCatalog.Applications[i].Roles[j].Name);
        for k := 0 to LCatalog.Applications[i].Roles[j].Users.Count - 1 do
          tvTree.Items.AddChild(RoleNode, LCatalog.Applications[i].Roles[j].Users[k].Name);
      end;
      LCatalog.Applications[i].GetInstances;
    end;
    tvTree.AlphaSort;
  finally
    LCatalog.Free;
  end;
end;

end.
