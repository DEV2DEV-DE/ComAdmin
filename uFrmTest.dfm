object FrmTest: TFrmTest
  Left = 0
  Top = 0
  Caption = 'FrmTest'
  ClientHeight = 585
  ClientWidth = 433
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  DesignSize = (
    433
    585)
  PixelsPerInch = 96
  TextHeight = 13
  object btnComAdmin: TButton
    Left = 8
    Top = 8
    Width = 75
    Height = 25
    Caption = 'ComAdmin'
    TabOrder = 0
    OnClick = btnComAdminClick
  end
  object tvTree: TTreeView
    Left = 8
    Top = 39
    Width = 417
    Height = 538
    Anchors = [akLeft, akTop, akRight, akBottom]
    Indent = 19
    ReadOnly = True
    SortType = stText
    TabOrder = 1
  end
  object cmbServer: TComboBox
    Left = 89
    Top = 10
    Width = 192
    Height = 21
    TabOrder = 2
    Items.Strings = (
      ''
      'defthwa0063srv'
      'defthwa006dsrv'
      'defthwa006esrv'
      'defthwa006fsrv')
  end
end
