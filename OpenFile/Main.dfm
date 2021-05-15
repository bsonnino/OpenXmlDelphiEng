object MainFrm: TMainFrm
  Left = 0
  Top = 0
  Caption = 'OpenXML with Delphi'
  ClientHeight = 396
  ClientWidth = 654
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Shell Dlg 2'
  Font.Style = []
  OldCreateOrder = False
  DesignSize = (
    654
    396)
  PixelsPerInch = 96
  TextHeight = 13
  object Button1: TButton
    Left = 571
    Top = 363
    Width = 75
    Height = 25
    Anchors = [akRight, akBottom]
    Caption = 'Open'
    TabOrder = 0
    OnClick = Button1Click
  end
  object Memo1: TMemo
    Left = 8
    Top = 8
    Width = 394
    Height = 341
    Anchors = [akLeft, akTop, akRight, akBottom]
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Consolas'
    Font.Style = []
    Lines.Strings = (
      'Memo1')
    ParentFont = False
    ScrollBars = ssBoth
    TabOrder = 1
    WordWrap = False
  end
  object ValueListEditor1: TValueListEditor
    Left = 408
    Top = 8
    Width = 237
    Height = 341
    Anchors = [akTop, akRight, akBottom]
    TabOrder = 2
    TitleCaptions.Strings = (
      'Property'
      'Value'
      '')
    ColWidths = (
      106
      125)
  end
  object OpenDialog1: TOpenDialog
    Filter = 
      'Arquivos Word (*.docx, *.docm)|*.docx;*.docm| Arquivos Excel (*.' +
      'xlsx, *.xlsm)|*.xlsx;*.xlsm| Arquivos Powerpoint (*.pptx, *.pptm' +
      ')|*.pptx;*.pptm'
    Left = 252
    Top = 204
  end
  object XMLDocument1: TXMLDocument
    Left = 320
    Top = 204
    DOMVendorDesc = 'MSXML'
  end
  object XMLDocument2: TXMLDocument
    Left = 352
    Top = 204
    DOMVendorDesc = 'MSXML'
  end
end
