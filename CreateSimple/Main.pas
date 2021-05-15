unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, xmldom, XMLIntf, msxmldom, XMLDoc, System.Zip;

type
  TMainFrm = class(TForm)
    Label1: TLabel;
    Memo1: TMemo;
    Button1: TButton;
    procedure Button1Click(Sender: TObject);
  private
    function CreateContentTypes(): TStream;
    function CreateDocument(): TStream;
    function CreateRels(): TStream;
    function CreateXml: IXmlDocument;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MainFrm: TMainFrm;

implementation

{$R *.dfm}

function TMainFrm.CreateXml: IXmlDocument;
begin
  Result := TXmlDocument.Create(nil);
  Result.Options := [doNodeAutoIndent];
  Result.Active := True;
  Result.Encoding := 'UTF-8';
  Result.Version := '1.0';
  Result.StandAlone := 'yes';
end;

function TMainFrm.CreateContentTypes(): TStream;
var
  Root: IXmlNode;
  Tipo: IXmlNode;
  XMLDoc: IXmlDocument;
begin
  Result := TMemoryStream.Create();
  XMLDoc := CreateXml;
  // Root node
  Root := XMLDoc.addChild('Types',
    'http://schemas.openxmlformats.org/package/2006/content-types');
  // Type definitions
  Tipo := Root.addChild('Default');
  Tipo.Attributes['Extension'] := 'rels';
  Tipo.Attributes['ContentType'] :=
    'application/vnd.openxmlformats-package.relationships+xml';
  Tipo := Root.addChild('Default');
  Tipo.Attributes['Extension'] := 'xml';
  Tipo.Attributes['ContentType'] :=
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
  // Save output stream
  XMLDoc.SaveToStream(Result);
  Result.Position := 0;
end;

function TMainFrm.CreateRels(): TStream;
var
  Root: IXmlNode;
  Rel: IXmlNode;
  XMLDoc: IXmlDocument;
begin
  Result := TMemoryStream.Create();
  XMLDoc := CreateXml;
  // Root node
  Root := XMLDoc.addChild('Relationships',
    'http://schemas.openxmlformats.org/package/2006/relationships');
  // Relation definitions
  Rel := Root.addChild('Relationship');
  Rel.Attributes['Id'] := 'rId1';
  Rel.Attributes['Type'] :=
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
  Rel.Attributes['Target'] := 'word/document.xml';
  // Save output stream
  XMLDoc.SaveToStream(Result);
  Result.Position := 0;
end;

function TMainFrm.CreateDocument(): TStream;
var
  Root: IXmlNode;
  XMLDoc: IXmlDocument;
begin
  Result := TMemoryStream.Create();
  XMLDoc := CreateXml;
  // Root node
  Root := XMLDoc.addChild('wordDocument',
    'http://schemas.openxmlformats.org/wordprocessingml/2006/main');
  // Save text
  Root.addChild('body').addChild('p').addChild('r').addChild('t').NodeValue :=
    Memo1.Text;
  // Save output stream
  XMLDoc.SaveToStream(Result);
  Result.Position := 0;
end;

procedure TMainFrm.Button1Click(Sender: TObject);
var
  zipFile: TZipFile;
  contentTypes: TStream;
  rels: TStream;
  doc: TStream;
begin
  zipFile := TZipFile.Create();
  try
    zipFile.Open('SimpleFile.docx', TZipMode.zmWrite);
    contentTypes := CreateContentTypes();
    try
      zipFile.Add(contentTypes, '[Content_Types].xml');
    finally
      contentTypes.Free;
    end;
    rels := CreateRels();
    try
      zipFile.Add(rels, '_rels\.rels');
    finally
      rels.Free;
    end;
    doc := CreateDocument();
    try
      zipFile.Add(doc, 'word\document.xml');
    finally
      doc.Free;
    end;
  finally
    zipFile.Close();
    zipFile.Free;
  end;
end;

end.
