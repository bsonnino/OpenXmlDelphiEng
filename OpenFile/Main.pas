unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, xmldom, XMLIntf, msxmldom, XMLDoc, Grids, ValEdit,
  System.Zip;

type
  TMainFrm = class(TForm)
    Button1: TButton;
    Memo1: TMemo;
    OpenDialog1: TOpenDialog;
    ValueListEditor1: TValueListEditor;
    XMLDocument1: TXMLDocument;
    XMLDocument2: TXMLDocument;
    procedure Button1Click(Sender: TObject);
  private
    procedure ReadProperties(ZipFile: TZipFile; const Arquivo: String);
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MainFrm: TMainFrm;

implementation

{$R *.dfm}

procedure TMainFrm.ReadProperties(ZipFile: TZipFile; const Arquivo: String);
var
  ZipStream: TStream;
  i: Integer;
  XmlNode: IXMLNode;
  LocalHeader: TZipHeader;
begin
  ZipFile.Read(Arquivo, ZipStream, LocalHeader);
  try
    ZipStream.Position := 0;
    XMLDocument2.LoadFromStream(ZipStream);
    // Read properties
    for i := 0 to XMLDocument2.DocumentElement.ChildNodes.Count - 1 do begin
      XmlNode := XMLDocument2.DocumentElement.ChildNodes.Nodes[i];
      try
        // Found new property - add
        ValueListEditor1.InsertRow(XmlNode.NodeName, XmlNode.NodeValue, True);
      except
        // Property is not a single value - ignore
        On EXMLDocError do;
        // Property is null - add null value
        On EVariantTypeCastError do
          ValueListEditor1.InsertRow(XmlNode.NodeName, '', True);
      end;
    end;
  finally
    ZipStream.Free;
  end;
end;

procedure TMainFrm.Button1Click(Sender: TObject);
var
  ZipStream: TStream;
  XmlNode: IXMLNode;
  i: Integer;
  AttType: String;
  ZipFile: TZipFile;
  LocalHeader: TZipHeader;
begin
  if OpenDialog1.Execute then
  begin
    ZipFile := TZipFile.Create();
    try
      ZipFile.Open(OpenDialog1.FileName, TZipMode.zmRead);

      // Read Relations
      ZipFile.Read('_rels/.rels', ZipStream, LocalHeader);
      try
        ZipStream.Position := 0;
        XMLDocument1.LoadFromStream(ZipStream);
        Memo1.Text := XMLDoc.FormatXMLData(XMLDocument1.XML.Text);
        ValueListEditor1.Strings.Clear;
        // Process nodes
        for i := 0 to XMLDocument1.DocumentElement.ChildNodes.Count - 1 do begin
          XmlNode := XMLDocument1.DocumentElement.ChildNodes.Nodes[i];
          // Get kind of relation
          // It's the final part of the Type attribute
          AttType := ExtractFileName(XmlNode.Attributes['Type']);
          if AttType.EndsWith('core-properties') or
            AttType.EndsWith('extended-properties') then
            // Add properties to ValueListEditor
            ReadProperties(ZipFile, XmlNode.Attributes['Target']);
        end;
      finally
        ZipStream.Free;
      end;
    finally
      ZipFile.Close();
      ZipFile.Free;
    end;
  end;
end;

end.
