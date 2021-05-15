object MainFrm: TMainFrm
  Left = 0
  Top = 0
  Caption = 'Create OpenXML file'
  ClientHeight = 207
  ClientWidth = 271
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Shell Dlg 2'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Button1: TButton
    Left = 98
    Top = 91
    Width = 75
    Height = 25
    Caption = 'Create'
    TabOrder = 0
    OnClick = Button1Click
  end
  object XMLDocument1: TXMLDocument
    Left = 192
    Top = 32
    DOMVendorDesc = 'MSXML'
  end
end
