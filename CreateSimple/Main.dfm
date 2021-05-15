object MainFrm: TMainFrm
  Left = 0
  Top = 0
  Caption = 'Create OpenXML file'
  ClientHeight = 284
  ClientWidth = 418
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Shell Dlg 2'
  Font.Style = []
  OldCreateOrder = False
  DesignSize = (
    418
    284)
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 8
    Width = 26
    Height = 13
    Caption = 'Text:'
  end
  object Memo1: TMemo
    Left = 8
    Top = 27
    Width = 402
    Height = 206
    Anchors = [akLeft, akTop, akRight, akBottom]
    TabOrder = 0
  end
  object Button1: TButton
    Left = 172
    Top = 251
    Width = 75
    Height = 25
    Anchors = [akLeft, akBottom]
    Caption = 'Create'
    TabOrder = 1
    OnClick = Button1Click
  end
end
