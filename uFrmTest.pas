unit uFrmTest;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ExtCtrls, System.ImageList, Vcl.ImgList;

type
  TFrmTest = class(TForm)
    tvTree: TTreeView;
    memoLog: TMemo;
    pnlTop: TPanel;
    btnComAdmin: TButton;
    cmbServerFrom: TComboBox;
    btnSync: TButton;
    splSplitter: TSplitter;
    cmbServerTo: TComboBox;
    ilSmall: TImageList;
    txtFilter: TEdit;
    procedure btnComAdminClick(Sender: TObject);
    procedure btnSyncClick(Sender: TObject);
  private
    procedure OnReadCOMObject(const AObjectType, AObjectName: string);
    procedure OnDebugMessage(const AMessage: string);
  public
    { Public-Deklarationen }
  end;

var
  FrmTest: TFrmTest;

implementation

uses
  uComAdmin,
  COMAdmin_TLB;

{$R *.dfm}

procedure TFrmTest.btnComAdminClick(Sender: TObject);
var
  LCatalog: TComAdminCatalog;
  i, j: Integer;
  AppNode, CompNode, IntfNode, MethodNode, RoleNode, UserNode: TTreeNode;
  k: Integer;
  l: Integer;
  m: Integer;
begin
  tvTree.Items.Clear;
  memoLog.Clear;
  LCatalog := TComAdminCatalog.Create(cmbServerFrom.Text, txtFilter.Text, OnReadCOMObject);
  try
    LCatalog.LibraryPath := 'C:\Temp';
    //LCatalog.ExportApplication(20, 'D:\Test123.msi');
    for i := 0 to LCatalog.Applications.Count - 1 do
    begin
      AppNode := tvTree.Items.AddChild(nil, LCatalog.Applications[i].Name);
      AppNode.ImageIndex := 0;
      AppNode.SelectedIndex := 0;
      for j := 0 to LCatalog.Applications[i].Components.Count - 1 do
      begin
        CompNode := tvTree.Items.AddChild(AppNode, LCatalog.Applications[i].Components[j].Name);
        CompNode.ImageIndex := 1;
        CompNode.SelectedIndex := 1;
        for k := 0 to LCatalog.Applications[i].Components[j].Interfaces.Count - 1 do
        begin
          IntfNode := tvTree.Items.AddChild(CompNode, LCatalog.Applications[i].Components[j].Interfaces[k].Name);
          IntfNode.ImageIndex := 2;
          IntfNode.SelectedIndex := 2;
          for l := 0 to LCatalog.Applications[i].Components[j].Interfaces[k].Methods.Count - 1 do
          begin
            MethodNode := tvTree.Items.AddChild(IntfNode, LCatalog.Applications[i].Components[j].Interfaces[k].Methods[l].Name);
            MethodNode.ImageIndex := 3;
            MethodNode.SelectedIndex := 3;
            for m := 0 to LCatalog.Applications[i].Components[j].Interfaces[k].Methods[l].Roles.Count - 1 do
            begin
              RoleNode := tvTree.Items.AddChild(MethodNode, LCatalog.Applications[i].Components[j].Interfaces[k].Methods[l].Roles[m].Name);
              RoleNode.ImageIndex := 4;
              RoleNode.SelectedIndex := 4;
            end;
          end;
          for l := 0 to LCatalog.Applications[i].Components[j].Interfaces[k].Roles.Count - 1 do
          begin
            RoleNode := tvTree.Items.AddChild(IntfNode, LCatalog.Applications[i].Components[j].Interfaces[k].Roles[l].Name);
            RoleNode.ImageIndex := 4;
            RoleNode.SelectedIndex := 4;
          end;
        end;
        for k := 0 to LCatalog.Applications[i].Components[j].Roles.Count - 1 do
        begin
          RoleNode := tvTree.Items.AddChild(CompNode, LCatalog.Applications[i].Components[j].Roles[k].Name);
          RoleNode.ImageIndex := 4;
          RoleNode.SelectedIndex := 4;
        end;
      end;
      for j := 0 to LCatalog.Applications[i].Roles.Count - 1 do
      begin
        RoleNode := tvTree.Items.AddChild(AppNode, LCatalog.Applications[i].Roles[j].Name);
        RoleNode.ImageIndex := 4;
        RoleNode.SelectedIndex := 4;
        for k := 0 to LCatalog.Applications[i].Roles[j].Users.Count - 1 do
        begin
          UserNode := tvTree.Items.AddChild(RoleNode, LCatalog.Applications[i].Roles[j].Users[k].Name);
          UserNode.ImageIndex := 5;
          UserNode.SelectedIndex := 5;
        end;
      end;
    end;
    tvTree.AlphaSort(False);
  finally
    LCatalog.Free;
  end;
end;

procedure TFrmTest.btnSyncClick(Sender: TObject);
var
  LCatalog: TComAdminCatalog;
begin
  memoLog.Clear;
  LCatalog := TComAdminCatalog.Create(cmbServerFrom.Text, txtFilter.Text, OnReadCOMObject);
  try
    LCatalog.LibraryPath := 'C:\Temp';
    LCatalog.OnDebug := OnDebugMessage;
    LCatalog.CopyLibraries := False;
    LCatalog.SyncToServer(cmbServerTo.Text, '');
  finally
    LCatalog.Free;
  end;
end;

procedure TFrmTest.OnDebugMessage(const AMessage: string);
begin
  memoLog.Lines.Add(AMessage);
  memoLog.ScrollBy(0, 1);
  Application.ProcessMessages;
end;

procedure TFrmTest.OnReadCOMObject(const AObjectType, AObjectName: string);
begin
  memoLog.Lines.AddPair(AObjectType, AObjectName);
  memoLog.ScrollBy(0, 1);
  Application.ProcessMessages;
end;

end.
