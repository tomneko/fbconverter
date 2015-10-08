// -----------------------------------------------------------------------------
//                       Firebird MetaData Utils
// -----------------------------------------------------------------------------
// Note:
//
// -----------------------------------------------------------------------------
unit uFBMetaData;

interface

uses
  System.Classes, Data.DB, SYSTEM.Types, uZeosFBUtils, ZSqlMetadata;

type
  TFirebirdMetaData = class(TComponent)
  private
    { Private 宣言 }
    Firebird: TZeosFB;
    function GetTableArray: TStringDynArray;
  public
    { Public 宣言 }
    // Constructor / Destructor
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  end;

implementation


{ TFirebirdMetaData }

constructor TFirebirdMetaData.Create(AOwner: TComponent);
begin
  inherited;
  Firebird := TZeosFB.Create(nil);
end;

destructor TFirebirdMetaData.Destroy;
begin
  Firebird.Free;
  inherited;
end;

function TFirebirdMetaData.GetTableArray: TStringDynArray;
begin
  Firebird.Connect;
  try
    Firebird.MetaData.MetadataType := mdTables;
    Firebird.MetaData.Open;
    while not Firebird.MetaData.Eof do
      begin


        Firebird.MetaData.Next;
      end;
    Firebird.MetaData.Close;
  finally
    Firebird.Disconnect;
  end;
end;

end.
