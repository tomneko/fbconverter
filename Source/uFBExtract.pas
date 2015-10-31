{*******************************************************}
{                                                       }
{              Firebird Database Converter              }
{                                                       }
{*******************************************************}

// -----------------------------------------------------------------------------
// Note:
//   - Metadata Generator Unit
// -----------------------------------------------------------------------------
unit uFBExtract;

interface

uses
  System.Classes, System.SysUtils, System.Types, System.IOUtils, System.Math,
  System.RegularExpressions, WinAPI.Windows, Data.DB, ZDataset, uZeosFBUtils,
  uSQLBuilder, Dialogs, ZCompatibility;

type
  { TFieldType }
  TFieldType = record
    FieldType: Integer;
    TypeName: string;
    FieldSubType: Integer;
    FieldLength: Integer;
    CharcterLength: Integer;
    SegmentLength: Integer;
    Precision: Integer;
    Scale: Integer;
    CharSetID: Integer;
    DefaultValue: string;
    HasDefault: Boolean;
    DefaultSource: string;
    HasValidation: Boolean;
    ValidationSource: string;
    NotNull: Boolean;
    Dimensions: Integer;
  end;

  { TFBCharSet }
  TFBCharSet = record
    ID: Integer;
    Name: string;
    BytesPerCharacter: Byte;
    MinimumVersion: TFBVersion;
  end;

  { TIndexRec }
  TIndexRec =
  record
    TableName: string;
    Name: string;
    Unique: Boolean;
    IndexType: Integer;
  end;

  { TIndexDynArray }
  TIndexDynArray = array of TIndexRec;

  { TTriggerRec }
  TTriggerRec =
  record
    TableName: string;
    Name: string;
    Active: Boolean;
    TriggerType: Integer;
    Sequence: Integer;
    Source: string;
  end;

  { TTriggerDynArray }
  TTriggerDynArray = array of TTriggerRec;

  { TInsertRec }
  TInsertRec =
  record
    Fields: string;
    Values1: string;
    Values2: string;
    Comment: string;
  end;

  { TTriggerDynArray }
  TInsRecDynArray = array of TInsertRec;

  { TSplitDataRec }
  TSplitDataRec =
  packed record
    Index: Integer;
    RecordCount: Int64;
    SplitSize: Integer;
    SplitCount: Integer;
    LastSize: Integer;
  end;

  { TFBExtractor }
  TFBExtractor = class (TComponent)
  private
    FB: TZeosFB;
    FDefaultCharset: TFBCharSet;
    FTargetCharset: TFBCharSet;
    FBrobOutputDir: string;
    FIsExtractBlobData: Boolean;
    FConvertFromIntToBigInt: Boolean;
    function BuildEnumObjects(FIELD, FROM, WHERE, ORDER: string): TStringDynArray;
    function BuildArgumentString(Query: TZQuery): string;
    function BuildDimensionString(FieldName: string): string;
    function BuildFieldString(FieldName: string; DomainFlg: Boolean = False): string;
    function GetFieldSubTypeString(FieldSubType: Integer): string;
    function GetFieldTypeString(FieldType: Integer): string;
    function BuildUniqueKey(TableName: string; IsPrimary: Boolean = True): string;
    procedure SetBrobOutputDir(const Value: string);
    procedure SetIsExtractBlobData(const Value: Boolean);
  public
    // Constructor / Destructor
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    // Methods
    function BuildCheck(TableName: string): string;
    function BuildConstraint(TableName: string): string;
    function BuildDomain(aName: string): string;
    function BuildException(aName: string): string;
    function BuildFilter(aName: string): string;
    function BuildForeignKey(TableName: string): string;
    function BuildFunction(aName: string): string;
    function BuildGenerator(aName: string): string;
    function BuildGeneratorValue(aName: string): string;
    function BuildIndex(Index: TIndexRec): string;
    function BuildPrimaryKey(TableName: string): string;
    function BuildProcedure(aName: string; aDefine: Boolean): string;
    function BuildRole(aName: string): string;
    function BuildTable(aName: string): string;
    function BuildTrigger(Trigger: TTriggerRec): string;
    function BuildUnique(TableName: string): string;
    function BuildView(aName: string): string;
    function ConvertString(Bytes: TArray<Byte>; SrcCodePage: uInt32; DstCodePage: uInt32): string;
    function EnumAndBuildGrant_Execute: TStringDynArray;
    function EnumAndBuildGrant_Object: TStringDynArray;
    function EnumAndBuildGrant_Role: TStringDynArray;
    function EnumDomains: TStringDynArray;
    function EnumExceptions: TStringDynArray;
    function EnumFilters: TStringDynArray;
    function EnumFunctions: TStringDynArray;
    function EnumGenerators: TStringDynArray;
    function EnumGrants(IsView: Boolean): TStringDynArray;
    function EnumIndicies(TableName: string): TIndexDynArray;
    function EnumProcedures: TStringDynArray;
    function EnumRoles: TStringDynArray;
    function EnumTables: TStringDynArray;
    function EnumTriggers(TableName: string): TTriggerDynArray;
    function EnumViews: TStringDynArray;
    function GetCharsetInfoByID(CharSetID: Integer): TFBCharset;
    function GetCharsetInfoByName(CharSetName: string): TFBCharset;
    class function PreBuildTableData(aExtractor: TFBExtractor; TableName: string; SplitSize: Integer = 1000): TSplitDataRec;
    class function BuildTableData(aExtractor: TFBExtractor; TableName: string; TableIndex: Integer; var SplitDataRec: TSplitDataRec): TInsRecDynArray;
    procedure Init;
    // Properties
    property BrobOutputDir: string read FBrobOutputDir write SetBrobOutputDir;
    property DefaultCharset: TFBCharSet read FDefaultCharset;
    property Firebird: TZeosFB read FB write FB;
    property IsExtractBlobData: Boolean read FIsExtractBlobData write SetIsExtractBlobData;
    property TargetCharset: TFBCharSet read FTargetCharset write FTargetCharset;
    property ConvertFromIntToBigInt: Boolean read FConvertFromIntToBigInt write FConvertFromIntToBigInt;
  end;

  TFieldTypes = array [0..12] of TFieldType;
  TFieldSubTypes = array [0..2] of TFieldType;
  TColmnType = (ctDataType, ctComputed, ctDomain);
  TCharSets = array [0..51] of TFBCharSet;

  function FileToString(FileName: string; CodePage: uInt32 = CP_UTF8): string;
  function QuotedString(s: string): string;
  procedure StringToFile(s: string; FileName: string; CodePage: uInt32 = CP_UTF8);

const
  SEPARATOR_STR = ', ';

  // Field Type
  fft_SmaiiInt  = 7;   // SMALLINT
  fft_Integer   = 8;   // INTEGER
  fft_Quad      = 9;   // QUAD
  fft_Float     = 10;  // FLOAT
  fft_DFloat    = 11;
  fft_Date      = 12;  // DATE
  fft_Time      = 13;  // TIME
  fft_Char      = 14;  // CHAR(n)
  fft_BigInt    = 16;  // BIGINT (Dialect 3)
  fft_Double    = 27;  // DECIMAL(p, s) / DOUBLE PRECISION / NUMERIC(p,s)
  fft_TimeStamp = 35;  // TIMESTAMP
  fft_VarChar   = 37;  // VARCHAR(n)
  fft_CString   = 40;  // CSTRING
  fft_Blob      = 261; // BLOB

  // Field Sub Type (BLOB)
  ffst_blob_None = 0;
  ffst_blob_Text = 1;
  ffst_blob_BLR  = 2;

  // Field Sub Type (CHAR)
  ffst_char_None = 0;
  ffst_char_FixedBinary = 1;

  // Object Type
  fot_Relation         = 0;
  fot_View             = 1;
  fot_Trigger          = 2;
  fot_Computed         = 3;
  fot_Validation       = 4;
  fot_Procedure        = 5;
  fot_Expression_Index = 6;
  fot_Exception        = 7;
  fot_User             = 8;
  fot_Field            = 9;
  fot_Index            = 10;
  fot_Count            = 11;
  fot_User_Group       = 12;
  fot_SQL_Role         = 13;

 	FieldTypes: TFieldTypes = (
    (FieldType: fft_SmaiiInt;  TypeName: 'SMALLINT'        ),
    (FieldType: fft_Integer;   TypeName: 'INTEGER'         ),
    (FieldType: fft_Quad;      TypeName: 'QUAD'            ),
    (FieldType: fft_Float;     TypeName: 'FLOAT'           ),
    (FieldType: fft_Date;      TypeName: 'DATE'            ),
    (FieldType: fft_Time;      TypeName: 'TIME'            ),
    (FieldType: fft_Char;      TypeName: 'CHAR'            ),
    (FieldType: fft_BigInt;    TypeName: 'BIGINT'          ),
    (FieldType: fft_Double;    TypeName: 'DOUBLE PRECISION'),
    (FieldType: fft_TimeStamp; TypeName: 'TIMESTAMP'       ),
    (FieldType: fft_VarChar;   TypeName: 'VARCHAR'         ),
    (FieldType: fft_CString;   TypeName: 'CSTRING'         ),
    (FieldType: fft_Blob;      TypeName: 'BLOB'            )
  );

 	FieldSubTypes: TFieldSubTypes = (
    (FieldType: ffst_blob_None; TypeName: ''                ),
    (FieldType: ffst_blob_Text; TypeName: 'TEXT'            ),
    (FieldType: ffst_blob_BLR;  TypeName: 'BLR'             )
  );

  { Charset List: from Firebird 2.5 RDB$CHARACTER_SETS }
  CharSets: TCharSets = (
    (ID:  0; Name: 'NONE';        BytesPerCharacter: 1; MinimumVersion: fbv10),
    (ID:  1; Name: 'OCTETS';      BytesPerCharacter: 1; MinimumVersion: fbv10),
    (ID:  2; Name: 'ASCII';       BytesPerCharacter: 1; MinimumVersion: fbv10),
    (ID:  3; Name: 'UNICODE_FSS'; BytesPerCharacter: 3; MinimumVersion: fbv10),
    (ID:  4; Name: 'UTF8';        BytesPerCharacter: 4; MinimumVersion: fbv20), // Firebird 2.0
    (ID:  5; Name: 'SJIS_0208';   BytesPerCharacter: 2; MinimumVersion: fbv10),
    (ID:  6; Name: 'EUCJ_0208';   BytesPerCharacter: 2; MinimumVersion: fbv10),
    (ID:  9; Name: 'DOS737';      BytesPerCharacter: 1; MinimumVersion: fbv10),
    (ID: 10; Name: 'DOS437';      BytesPerCharacter: 1; MinimumVersion: fbv10),
    (ID: 11; Name: 'DOS850';      BytesPerCharacter: 1; MinimumVersion: fbv10),
    (ID: 12; Name: 'DOS865';      BytesPerCharacter: 1; MinimumVersion: fbv10),
    (ID: 13; Name: 'DOS860';      BytesPerCharacter: 1; MinimumVersion: fbv10),
    (ID: 14; Name: 'DOS863';      BytesPerCharacter: 1; MinimumVersion: fbv10),
    (ID: 15; Name: 'DOS775';      BytesPerCharacter: 1; MinimumVersion: fbv15), // Firebird 1.5
    (ID: 16; Name: 'DOS858';      BytesPerCharacter: 1; MinimumVersion: fbv15), // Firebird 1.5
    (ID: 17; Name: 'DOS862';      BytesPerCharacter: 1; MinimumVersion: fbv15), // Firebird 1.5
    (ID: 18; Name: 'DOS864';      BytesPerCharacter: 1; MinimumVersion: fbv15), // Firebird 1.5
    (ID: 19; Name: 'NEXT';        BytesPerCharacter: 1; MinimumVersion: fbv10),
    (ID: 21; Name: 'ISO8859_1';   BytesPerCharacter: 1; MinimumVersion: fbv10),
    (ID: 22; Name: 'ISO8859_2';   BytesPerCharacter: 1; MinimumVersion: fbv10),
    (ID: 23; Name: 'ISO8859_3';   BytesPerCharacter: 1; MinimumVersion: fbv15), // Firebird 1.5
    (ID: 34; Name: 'ISO8859_4';   BytesPerCharacter: 1; MinimumVersion: fbv15), // Firebird 1.5
    (ID: 35; Name: 'ISO8859_5';   BytesPerCharacter: 1; MinimumVersion: fbv15), // Firebird 1.5
    (ID: 36; Name: 'ISO8859_6';   BytesPerCharacter: 1; MinimumVersion: fbv15), // Firebird 1.5
    (ID: 37; Name: 'ISO8859_7';   BytesPerCharacter: 1; MinimumVersion: fbv15), // Firebird 1.5
    (ID: 38; Name: 'ISO8859_8';   BytesPerCharacter: 1; MinimumVersion: fbv15), // Firebird 1.5
    (ID: 39; Name: 'ISO8859_9';   BytesPerCharacter: 1; MinimumVersion: fbv15), // Firebird 1.5
    (ID: 40; Name: 'ISO8859_13';  BytesPerCharacter: 1; MinimumVersion: fbv15), // Firebird 1.5
    (ID: 44; Name: 'KSC_5601';    BytesPerCharacter: 2; MinimumVersion: fbv10),
    (ID: 45; Name: 'DOS852';      BytesPerCharacter: 1; MinimumVersion: fbv10),
    (ID: 46; Name: 'DOS857';      BytesPerCharacter: 1; MinimumVersion: fbv10),
    (ID: 47; Name: 'DOS861';      BytesPerCharacter: 1; MinimumVersion: fbv10),
    (ID: 48; Name: 'DOS866';      BytesPerCharacter: 1; MinimumVersion: fbv15), // Firebird 1.5
    (ID: 49; Name: 'DOS869';      BytesPerCharacter: 1; MinimumVersion: fbv15), // Firebird 1.5
    (ID: 50; Name: 'CYRL';        BytesPerCharacter: 1; MinimumVersion: fbv10),
    (ID: 51; Name: 'WIN1250';     BytesPerCharacter: 1; MinimumVersion: fbv10),
    (ID: 52; Name: 'WIN1251';     BytesPerCharacter: 1; MinimumVersion: fbv10),
    (ID: 53; Name: 'WIN1252';     BytesPerCharacter: 1; MinimumVersion: fbv10),
    (ID: 54; Name: 'WIN1253';     BytesPerCharacter: 1; MinimumVersion: fbv10),
    (ID: 55; Name: 'WIN1254';     BytesPerCharacter: 1; MinimumVersion: fbv10),
    (ID: 56; Name: 'BIG_5';       BytesPerCharacter: 2; MinimumVersion: fbv10),
    (ID: 57; Name: 'GB_2312';     BytesPerCharacter: 2; MinimumVersion: fbv10),
    (ID: 58; Name: 'WIN1255';     BytesPerCharacter: 1; MinimumVersion: fbv15), // Firebird 1.5
    (ID: 59; Name: 'WIN1256';     BytesPerCharacter: 1; MinimumVersion: fbv15), // Firebird 1.5
    (ID: 60; Name: 'WIN1257';     BytesPerCharacter: 1; MinimumVersion: fbv15), // Firebird 1.5
    (ID: 63; Name: 'KOI8R';       BytesPerCharacter: 1; MinimumVersion: fbv20), // Firebird 2.0
    (ID: 64; Name: 'KOI8U';       BytesPerCharacter: 1; MinimumVersion: fbv20), // Firebird 2.0
    (ID: 65; Name: 'WIN1258';     BytesPerCharacter: 1; MinimumVersion: fbv20), // Firebird 2.0
    (ID: 66; Name: 'TIS620';      BytesPerCharacter: 1; MinimumVersion: fbv21), // Firebird 2.1
    (ID: 67; Name: 'GBK';         BytesPerCharacter: 2; MinimumVersion: fbv21), // Firebird 2.1
    (ID: 68; Name: 'CP943C';      BytesPerCharacter: 2; MinimumVersion: fbv21), // Firebird 2.1
    (ID: 69; Name: 'GB18030';     BytesPerCharacter: 4; MinimumVersion: fbv25)  // Firebird 2.5
  );

  KeyString: array [Boolean] of string = ('UNIQUE', 'PRIMARY KEY');

implementation

function QuotedString(s: string): string;
begin
  Result := Format('"%s"', [s]);
end;

function FileToString(FileName: string; CodePage: uInt32 = CP_UTF8): string;
// ファイルを String に読み込み
var
  Buf: string;
  F: TextFile;
  OldValue: Byte;
begin
  OldValue := FileMode;
  try
    FileMode := fmOpenRead;
    AssignFile(F, FileName, CodePage);
    Reset(F);
    Read(F, Buf);
    CloseFile(F);
  finally
    FileMode := OldValue;
  end;
  Result := Buf;
end;

procedure StringToFile(s: string; FileName: string; CodePage: uInt32 = CP_UTF8);
// 文字列をファイルとして出力
var
  F: TextFile;
  OldValue: Byte;
begin
  OldValue := FileMode;
  try
    FileMode := fmOpenWrite;
    AssignFile(F, FileName, CodePage);
    Rewrite(F);
    Write(F, s);
    CloseFile(F);
  finally
    FileMode := OldValue;
  end;
end;

{ TFBExtractor }

constructor TFBExtractor.Create(AOwner: TComponent);
// コンストラクタ
begin
  inherited;
  FB := TZeosFB.Create(Self);
  FBrobOutputDir := '';
  FIsExtractBlobData := False;
  FConvertFromIntToBigInt := False;
end;

destructor TFBExtractor.Destroy;
// デストラクタ
begin
  FB.Free;
  inherited;
end;

function TFBExtractor.GetCharsetInfoByName(CharSetName: string): TFBCharset;
// 文字コード情報取得
var
  Charset: TFBCharSet;
begin
  Result := CharSets[0];
  for Charset in CharSets do
    begin
      if Charset.Name = CharSetName then
        begin
          Result := Charset;
          Break;
        end;
    end;
end;

function TFBExtractor.GetCharsetInfoByID(CharSetID: Integer): TFBCharset;
// 文字コード情報取得
var
  Charset: TFBCharSet;
begin
  Result := CharSets[0];
  for Charset in CharSets do
    begin
      if Charset.ID = CharSetID then
        begin
          Result := Charset;
          Break;
        end;
    end;
end;

procedure TFBExtractor.SetBrobOutputDir(const Value: string);
// Blob 出力先の指定
begin
  FBrobOutputDir := Value;
end;

procedure TFBExtractor.SetIsExtractBlobData(const Value: Boolean);
// Blob データファイルを出力するか？
begin
  FIsExtractBlobData := Value;
end;

procedure TFBExtractor.Init;
// 初期化 (基本情報の取得)
var
  SelectSQL: TSelectSQL;
begin
  if FB.Connect then
    begin
      with FB.Query[0] do
        begin
          SelectSQL.SELECT := 'B.RDB$CHARACTER_SET_ID, B.RDB$CHARACTER_SET_NAME, B.RDB$BYTES_PER_CHARACTER';
          SelectSQL.FROM   := 'RDB$DATABASE A LEFT JOIN RDB$CHARACTER_SETS B ON (B.RDB$CHARACTER_SET_NAME = A.RDB$CHARACTER_SET_NAME)';
          SQL.Text := SelectSQL.Build;
          Open;
          if IsEmpty then
            begin
              // NONE
              FDefaultCharset.ID                := 0;
              FDefaultCharset.Name              := DEFAULT_CHARSET;
              FDefaultCharset.BytesPerCharacter := 1;
            end
          else
            begin
              FDefaultCharset.ID                := Fields[0].AsInteger;
              FDefaultCharset.Name              := Fields[1].AsString;
              FDefaultCharset.BytesPerCharacter := Fields[2].AsInteger;
            end;
          Close;
        end;
      FB.Disconnect;
    end;
end;

function TFBExtractor.BuildFieldString(FieldName: string; DomainFlg: Boolean): string;
// フィールドの生成
const
  SelectSQL: TSelectSQL = (SELECT: '*'; FROM: 'RDB$FIELDS'; WHERE: '(RDB$FIELD_NAME = :_FIELD_NAME)');
var
  CharLen, Precision: Integer;
  CharsetInfo: TFBCharSet;
  FieldRec: TFieldType;
  FieldStr, FieldSubStr, RetStr: string;
begin
  RetStr := '';
  with FB.Query[1] do
    begin
      FieldStr := '';
      SQL.Text := SelectSQL.Build;
      ParamByName('_FIELD_NAME').AsString := FieldName;
      Open;
      if not IsEmpty then
        begin
          FieldRec.TypeName         := Trim(FieldByName('RDB$FIELD_NAME').AsString);
          FieldRec.FieldType        := FieldByName('RDB$FIELD_TYPE').AsInteger;
          FieldRec.FieldSubType     := FieldByName('RDB$FIELD_SUB_TYPE').AsInteger;
          FieldRec.FieldLength      := FieldByName('RDB$FIELD_LENGTH').AsInteger;
          if FieldByName('RDB$CHARACTER_LENGTH').IsNull then
            FieldRec.CharcterLength := FieldRec.FieldLength
          else
            FieldRec.CharcterLength := FieldByName('RDB$CHARACTER_LENGTH').AsInteger;
          FieldRec.SegmentLength    := FieldByName('RDB$SEGMENT_LENGTH').AsInteger;
          if FindField('RDB$FIELD_PRECISION') <> nil then
            FieldRec.Precision     := FieldByName('RDB$FIELD_PRECISION').AsInteger;
          FieldRec.Scale            := FieldByName('RDB$FIELD_SCALE').AsInteger;
          FieldRec.CharSetID        := FieldByName('RDB$CHARACTER_SET_ID').AsInteger;
          FieldRec.HasDefault       := (not FieldByName('RDB$DEFAULT_VALUE').IsNull);
          if FieldRec.HasDefault then
            FieldRec.DefaultSource := FieldByName('RDB$DEFAULT_SOURCE').AsString
          else
            FieldRec.DefaultSource := '';
          FieldRec.HasValidation    := (not FieldByName('RDB$VALIDATION_BLR').IsNull);
          if FieldRec.HasDefault then
            FieldRec.ValidationSource := FieldByName('RDB$VALIDATION_SOURCE').AsString
          else
            FieldRec.ValidationSource := '';
          FieldRec.Dimensions       := FieldByName('RDB$DIMENSIONS').AsInteger;
          FieldRec.NotNull          := (FieldByName('RDB$NULL_FLAG').AsInteger = 1);
          // カラムデータ種類
          if not FieldByName('RDB$COMPUTED_BLR').IsNull then
            begin
              // 計算カラム
              FieldStr := 'COMPUTED BY';
              if not FieldByName('RDB$COMPUTED_SOURCE').IsNull then
                FieldStr := FieldStr + Format(' (%s)', [Trim(FieldByName('RDB$COMPUTED_SOURCE').AsString)]);
            end
          else if DomainFlg then
            begin
              // ドメインカラム
              FieldStr := FieldStr + QuotedString(FieldName);
            end
          else
            begin
              // 通常カラム
              if FieldRec.Scale < 0 then
                begin
                  // 小数点を持つ型
                  if FindField('RDB$FIELD_PRECISION') = nil then
                    begin
                      // RDB$FIELD_PRECISION フィールドが存在しない
                      case FieldRec.FieldType of
                        fft_SmaiiInt:
                          Precision := 4;
                        fft_Integer:
                          Precision := 9;
                        fft_Double:
                          Precision := 15;
                      else
                        Precision := 0;
                      end;
                    end
                  else
                    begin
                      // RDB$FIELD_PRECISION フィールドが存在する
                      Precision := FieldRec.Precision;
                    end;
                  if Precision = 0 then
                    FieldStr := GetFieldTypeString(FieldRec.FieldType)
                  else  
                    FieldStr := Format('NUMERIC(%d, %d)', [Precision, -FieldRec.Scale]);
                end
              else
                begin
                  // 小数点を持たない型
                  FieldStr := GetFieldTypeString(FieldRec.FieldType);
                  case FieldRec.FieldType of
                    fft_Char,
                    fft_VarChar:
                      begin
                        CharsetInfo := GetCharsetInfoByID(FieldRec.CharSetID);
                        // 変換先の BPC が変換元の BPC より大きく、
                        // かつ "文字数×変換先 BPC" が 32765 を超えるようなら文字数を変換
                        if (CharsetInfo.BytesPerCharacter < TargetCharset.BytesPerCharacter) and
                           ((FieldRec.CharcterLength * TargetCharset.BytesPerCharacter) > 32765) then
                          CharLen := (FieldRec.CharcterLength div TargetCharset.BytesPerCharacter)
                        else
                          CharLen := FieldRec.CharcterLength;
                        FieldStr := FieldStr + Format('(%d)', [CharLen]);
                        // 文字コード
                        if FieldRec.CharSetID <> DefaultCharset.ID then
                          FieldStr := FieldStr + Format(' CHARACTER SET %s', [CharsetInfo.Name]);
                      end;
                    fft_Blob:
                      begin
                        FieldSubStr := GetFieldSubTypeString(FieldRec.FieldSubType);
                        // セグメントサイズ
                        FieldStr := FieldStr + Format(' SUB_TYPE %s SEGMENT SIZE %d', [FieldSubStr, FieldRec.SegmentLength]);
                        // 文字コード
                        if (FieldRec.FieldSubType = ffst_blob_Text) and (FieldRec.CharSetID <> DefaultCharset.ID) then
                          FieldStr := FieldStr + Format(' CHARACTER SET %s', [GetCharsetInfoByID(FieldRec.CharSetID).Name]);
                      end;
                    fft_Integer:
                      begin
                        if FConvertFromIntToBigInt then
                          FieldStr := GetFieldTypeString(fft_BigInt)
                        else
                          FieldStr := GetFieldTypeString(fft_Integer);
                      end
                  else
                    FieldStr := GetFieldTypeString(FieldRec.FieldType);
                  end;
                end;
              // 配列
              if FieldRec.Dimensions > 0 then
                FieldStr := FieldStr + BuildDimensionString(FieldName);
            end;
          if not DomainFlg then
            begin
              // デフォルト値
              if FieldRec.HasDefault then
                begin
                  FieldSubStr := FieldRec.DefaultSource;
                  if Pos('default', LowerCase(FieldSubStr)) = 1 then
                    begin
                      FieldSubStr := 'DEFAULT' + Copy(FieldSubStr, 8, Length(FieldSubStr));
                      FieldStr := FieldStr + Format(' %s', [FieldSubStr]);
                    end;
                end;
              // チェック制約
              if FieldRec.HasValidation then
                FieldStr := FieldStr + Format(' %s', [FieldRec.ValidationSource]);
              // NOT NULL
              if FieldRec.NotNull then
                FieldStr := FieldStr + ' NOT NULL';
            end;
          RetStr := RetStr + FieldStr;
        end;
      Close;
    end;
  Result := RetStr;
end;

function TFBExtractor.BuildArgumentString(Query: TZQuery): string;
// 引数の生成
var
  CharLen, Precision: Integer;
  CharsetInfo: TFBCharSet;
  FieldRec: TFieldType;
  FieldStr, FieldSubStr: string;
begin
  with Query do
    begin
      FieldStr := '';
      FieldRec.FieldType      := FieldByName('RDB$FIELD_TYPE').AsInteger;
      FieldRec.FieldSubType   := FieldByName('RDB$FIELD_SUB_TYPE').AsInteger;
      FieldRec.FieldLength    := FieldByName('RDB$FIELD_LENGTH').AsInteger;
      if FieldByName('RDB$CHARACTER_LENGTH').IsNull then
        FieldRec.CharcterLength := FieldRec.FieldLength
      else
        FieldRec.CharcterLength := FieldByName('RDB$CHARACTER_LENGTH').AsInteger;
      if FindField('RDB$FIELD_PRECISION') <> nil then
        FieldRec.Precision     := FieldByName('RDB$FIELD_PRECISION').AsInteger;
      FieldRec.Scale         := FieldByName('RDB$FIELD_SCALE').AsInteger;
      FieldRec.CharSetID     := FieldByName('RDB$CHARACTER_SET_ID').AsInteger;
      if FieldRec.Scale < 0 then
        begin
          // 小数点を持つ型
          if FindField('RDB$FIELD_PRECISION') = nil then
            begin
              // RDB$FIELD_PRECISION フィールドが存在しない
              case FieldRec.FieldType of
                fft_SmaiiInt:
                  Precision := 4;
                fft_Integer:
                  Precision := 9;
                fft_Double:
                  Precision := 15;
              else
                Precision := 0;
              end;
            end
          else
            begin
              // RDB$FIELD_PRECISION フィールドが存在する
              Precision := FieldRec.Precision;
            end;
          if Precision = 0 then
            FieldStr := GetFieldTypeString(FieldRec.FieldType)
          else
            FieldStr := Format('NUMERIC(%d, %d)', [Precision, -FieldRec.Scale]);
        end
      else
        begin
          // 小数点を持たない型
          FieldStr := GetFieldTypeString(FieldRec.FieldType);
          case FieldRec.FieldType of
            fft_Char,
            fft_VarChar,
            fft_CString:
              begin
                CharsetInfo := GetCharsetInfoByID(FieldRec.CharSetID);
                // 変換先の BPC が変換元の BPC より大きく、
                // かつ "文字数×変換先 BPC" が 32765 を超えるようなら文字数を変換
                if (CharsetInfo.BytesPerCharacter < TargetCharset.BytesPerCharacter) and
                   ((FieldRec.CharcterLength * TargetCharset.BytesPerCharacter) > 32765) then
                  CharLen := (FieldRec.CharcterLength div TargetCharset.BytesPerCharacter)
                else
                  CharLen := FieldRec.CharcterLength;
                FieldStr := FieldStr + Format('(%d)', [CharLen]);
                // 文字コード
                if FieldRec.CharSetID <> DefaultCharset.ID then
                  FieldStr := FieldStr + Format(' CHARACTER SET %s', [CharsetInfo.Name]);
              end;
            fft_Blob:
              begin
                FieldSubStr := GetFieldSubTypeString(FieldRec.FieldSubType);
                // 文字コード
                if (FieldRec.FieldSubType = ffst_blob_Text) and (FieldRec.CharSetID <> DefaultCharset.ID) then
                  FieldStr := FieldStr + Format(' CHARACTER SET %s', [GetCharsetInfoByID(FieldRec.CharSetID).Name]);
              end;
            fft_Integer:
              begin
                if FConvertFromIntToBigInt then
                  FieldStr := GetFieldTypeString(fft_BigInt)
                else
                  FieldStr := GetFieldTypeString(fft_Integer);
              end
          else
            FieldStr := GetFieldTypeString(FieldRec.FieldType);
          end;
        end;
      Result := FieldStr;
    end;
end;

function TFBExtractor.BuildDimensionString(FieldName: string): string;
// 配列の生成
const
  SelectSQL: TSelectSQL = (SELECT: '*'; FROM: 'RDB$FIELD_DIMENSIONS'; WHERE: '(RDB$FIELD_NAME = :_FIELD_NAME)'; ORDER: 'RDB$FIELD_NAME, RDB$DIMENSION');
var
  DimStr: string;
  Lower, Upper: Integer;
begin
  Result := '';
  with FB.Query[2] do
    begin
      SQL.Text := SelectSQL.Build;
      ParamByName('_FIELD_NAME').AsString := FieldName;
      DimStr := '';
      Open;
      while not EOF do
        begin
          Lower := FieldByName('RDB$LOWER_BOUND').AsInteger;
          Upper := FieldByName('RDB$UPPER_BOUND').AsInteger;
          if Lower <> 1 then
            DimStr := DimStr + Format('%d:', [Lower]);
          DimStr := DimStr + Format('%d', [Upper]);
          Next;
          if not Eof then
            DimStr := DimStr + SEPARATOR_STR;
        end;
      Close;
    end;
  if DimStr <> '' then
    Result := Format('[%s]', [Trim(DimStr)]);
end;

function TFBExtractor.BuildDomain(aName: string): string;
// ドメインの生成
begin
  Result := Format('CREATE DOMAIN %s AS %s%s', [QuotedString(aName), BuildFieldString(aName), TERM_CHAR]);
end;

function TFBExtractor.BuildException(aName: string): string;
// 例外の生成
const
  SelectSQL: TSelectSQL = (SELECT: '*'; FROM: 'RDB$EXCEPTIONS'; WHERE: '(RDB$EXCEPTION_NAME = :_EXCEPTION_NAME)');
var
  Msg: string;
begin
  Result := '';
  with FB.Query[0] do
    begin
      SQL.Text := SelectSQL.Build;
      ParamByName('_EXCEPTION_NAME').AsString := aName;
      Open;
      if not IsEmpty then
        begin
          Msg := FieldByName('RDB$MESSAGE').AsString;
          Result := Format('CREATE EXCEPTION %s ''%s''%s', [QuotedString(aName), Msg, TERM_CHAR]);
        end;
      Close;
    end;
end;

function TFBExtractor.BuildFilter(aName: string): string;
// BLOB フィルタの生成
const
  SelectSQL: TSelectSQL = (SELECT: '*'; FROM: 'RDB$FILTERS'; WHERE: '(RDB$FUNCTION_NAME = :_FUNCTION_NAME)');
var
  ENTRYPOINT, MODULE_NAME: string;
  INPUT_TYPE, OUTPUT_TYPE: Integer;
begin
  Result := Format('DECLARE FILTER %s', [QuotedString(aName)]) + sLineBreak;
  with FB.Query[0] do
    begin
      SQL.Text := SelectSQL.Build;
      ParamByName('_FUNCTION_NAME').AsString := aName;
      Open;
      if not IsEmpty then
        begin
          INPUT_TYPE  := FieldByName('RDB$INPUT_SUB_TYPE').AsInteger;
          OUTPUT_TYPE := FieldByName('RDB$OUTPUT_SUB_TYPE').AsInteger;
          ENTRYPOINT  := Trim(FieldByName('RDB$ENTRYPOINT').AsString);
          MODULE_NAME := Trim(FieldByName('RDB$MODULE_NAME').AsString);
          Result := Result + Format('INPUT_TYPE %d OUTPUT_TYPE %d%sENTRY_POINT ''%s'' MODULE_NAME ''%s''%s',
            [INPUT_TYPE, OUTPUT_TYPE, sLineBreak, ENTRYPOINT, MODULE_NAME, TERM_CHAR]);
        end;
      Close;
    end;
end;

function TFBExtractor.BuildFunction(aName: string): string;
// UDF の生成
var
  Argument, ENTRYPOINT, MODULE_NAME, RetVal: string;
  SelectSQL: TSelectSQL;
  MECHANISM, RETURN_ARGUMENT: Integer;
begin
  RETURN_ARGUMENT := -1;
  Result := Format('DECLARE EXTERNAL FUNCTION %s', [QuotedString(aName)]) + sLineBreak;
  with FB.Query[0] do
    begin
      SelectSQL.SELECT := '*';
      SelectSQL.FROM   := 'RDB$FUNCTIONS';
      SelectSQL.WHERE  := '(RDB$FUNCTION_NAME = :_FUNCTION_NAME)';
      SQL.Text := SelectSQL.Build;
      ParamByName('_FUNCTION_NAME').AsString := aName;
      Open;
      if not IsEmpty then
        begin
          MODULE_NAME     := Trim(FieldByName('RDB$MODULE_NAME').AsString);
          ENTRYPOINT      := Trim(FieldByName('RDB$ENTRYPOINT').AsString);
          RETURN_ARGUMENT := FieldByName('RDB$RETURN_ARGUMENT').AsInteger;
        end;
      Close;
    end;
  // 引数と戻り値
  with FB.Query[0] do
    begin
      // 引数
      Argument := '';
      SelectSQL.SELECT := '*';
      SelectSQL.FROM   := 'RDB$FUNCTION_ARGUMENTS';
      SelectSQL.WHERE  := '(RDB$FUNCTION_NAME = :_FUNCTION_NAME) and (RDB$ARGUMENT_POSITION <> 0)';
      SelectSQL.ORDER  := 'RDB$ARGUMENT_POSITION';
      SQL.Text := SelectSQL.Build;
      ParamByName('_FUNCTION_NAME'    ).AsString  := aName;
      Open;
      while not EOF do
        begin
          Argument := Argument + BuildArgumentString(FB.Query[0]);
          Next;
          if not EOF then
            Argument := Argument + SEPARATOR_STR;
        end;
      Close;
     // 戻り値
      RetVal := '';
      if RETURN_ARGUMENT <> -1 then
        begin
          SelectSQL.WHERE  := '(RDB$FUNCTION_NAME = :_FUNCTION_NAME) AND (RDB$ARGUMENT_POSITION = :_ARGUMENT_POSITION)';
          SQL.Text := SelectSQL.Build;
          ParamByName('_FUNCTION_NAME'    ).AsString  := aName;
          ParamByName('_ARGUMENT_POSITION').AsInteger := RETURN_ARGUMENT;
          Open;
          if not IsEmpty then
            begin
              if RETURN_ARGUMENT = 0 then
                RetVal := Format('RETURNS %s', [BuildArgumentString(FB.Query[0])])
              else
                RetVal := Format('RETURNS PARAMETER %d', [RETURN_ARGUMENT]);
              MECHANISM := FieldByName('RDB$MECHANISM').AsInteger;
              if MECHANISM = 0 then
                RetVal := RetVal + ' BY VALUE'
              else if MECHANISM < 0 then
                RetVal := RetVal + ' FREE_IT';
            end;
          Close;
        end;
    end;
  if Argument <> '' then
    Result := Result + Argument + sLineBreak;
  if RetVal <> '' then
    Result := Result + RetVal + sLineBreak;
  Result := Result + Format('ENTRY_POINT ''%s'' MODULE_NAME ''%s''%s', [ENTRYPOINT, MODULE_NAME, TERM_CHAR]);
end;

function TFBExtractor.BuildGenerator(aName: string): string;
// ジェネレータの生成
begin
  Result := Format('CREATE GENERATOR %s%s', [QuotedString(aName), TERM_CHAR]);
end;

function TFBExtractor.BuildGeneratorValue(aName: string): string;
// ジェネレータの値の生成
const
  SelectSQL: TSelectSQL = (SELECT: 'GEN_ID(%s, 0)'; FROM: 'RDB$DATABASE');
var
  Value: Integer;
begin
  Value := 0;
  with FB.Query[0] do
    begin
      SQL.Text := Format(SelectSQL.Build, [aName]);
      Open;
      if not IsEmpty then
        Value := Fields[0].AsInteger;
      Close;
    end;
  Result := Format('SET GENERATOR %s TO %d%s', [QuotedString(aName), Value, TERM_CHAR]);
end;

function TFBExtractor.EnumAndBuildGrant_Execute: TStringDynArray;
// GRANT (ストアド実行権限) の生成
const
  GRANT_OPSION_STR: array [Boolean] of string = ('', 'WITH GRANT OPTION');
var
  Cnt, RecCnt: Integer;
  ObjectStr: string;
  SelectSQL: TSelectSQL;
begin
  with FB.Query[1] do
    begin
      // GRANT の件数を得る
      SelectSQL.SELECT := 'Count(*)';
      SelectSQL.FROM   := 'RDB$USER_PRIVILEGES A LEFT JOIN RDB$PROCEDURES B ON (B.RDB$PROCEDURE_NAME = A.RDB$RELATION_NAME)';
      SelectSQL.WHERE  := '(A.RDB$PRIVILEGE = ''X'') AND (B.RDB$OWNER_NAME <> A.RDB$USER)';
      SQL.Text := SelectSQL.Build;
      Open;
      RecCnt := Fields[0].AsInteger;
      Close;

      // 戻り値の配列のサイズを変更
      SetLength(Result, RecCnt);

      // GRANT を列挙
      if RecCnt > 0 then
        begin
          Cnt := 0;
          SelectSQL.SELECT := 'A.RDB$RELATION_NAME, A.RDB$USER, A.RDB$USER_TYPE, A.RDB$GRANT_OPTION';
          SelectSQL.ORDER  := 'A.RDB$RELATION_NAME';
          SQL.Text := SelectSQL.Build;
          Open;
          while not EOF do
            begin
              case Fields[2].AsInteger of
                fot_View:
                  ObjectStr := 'VIEW';
                fot_Trigger:
                  ObjectStr := 'TRIGGER';
                fot_Procedure:
                  ObjectStr := 'PROCEDURE';
              else
                ObjectStr := '';
              end;
              if ObjectStr = '' then
                ObjectStr := ObjectStr + Trim(Fields[1].AsString)
              else
                ObjectStr := Format('%s %s', [ObjectStr, Trim(Fields[1].AsString)]);
              Result[Cnt] := Format('GRANT EXECUTE ON PROCEDURE "%s" TO %s%s%s',
                                    [Fields[0].AsString, ObjectStr, GRANT_OPSION_STR[Fields[3].AsInteger = 1], TERM_CHAR]);
              Inc(Cnt);
              Next;
              if not EOF then
                Result[Cnt] := Result[Cnt] + sLineBreak;
            end;
          Close;
        end;
    end;
end;

function TFBExtractor.EnumAndBuildGrant_Object: TStringDynArray;
// GRANT (オブジェクト) の生成
var
  Grants: TStringDynArray;
  Index: Integer;

  { BuildGrant_Object BEGIN }
  function BuildGrant_Object(aGrants: TStringDynArray; var retGrants: TStringDynArray; aIndex: Integer): Integer;
  type
    TPrivType = (privSelect, privDelete, privInsert, privUpdate, privReferences);
    TPrivTypes = set of TPrivType;
  const
    GRAND_OPTION_STR: array [Boolean] of string = ('', ' WITH GRANT OPTION');
  var
    Cnt: Integer;
    GrantOption_Flg: Boolean;
    Privilege, PrivilegeStr, s, u: string;
    PrivTypes: TPrivTypes;
    ReferrencesFieldList, UpdateFieldList: TStringList;
    SelectSQL: TSelectSQL;

    { BuildPrivilegeFieldString BEGIN }
    function BuildPrivilegeFieldString(aSL: TStringList): string;
    var
      i: Integer;
    begin
      Result := '';
      if aSL.Count > 0 then
        begin
          for i:=0 to aSL.Count-1 do
            begin
              Result := Result + QuotedString(Trim(aSL[i]));
              if i < (aSL.Count-1) then
                Result := Result + SEPARATOR_STR;
            end;
          Result := Format(' (%s)', [Result]);
        end;
    end;
    { BuildPrivilegeFieldString END }
  begin
    with FB.Query[1] do
      begin
        SelectSQL.SELECT := 'A.RDB$USER';
        SelectSQL.FROM   := 'RDB$USER_PRIVILEGES A LEFT JOIN RDB$RELATIONS B ON (B.RDB$RELATION_NAME = A.RDB$RELATION_NAME)';
        SelectSQL.WHERE  := '(A.RDB$USER <> B.RDB$OWNER_NAME) AND (A.RDB$PRIVILEGE <> ''M'') AND (A.RDB$PRIVILEGE <> ''X'') AND (A.RDB$RELATION_NAME = :_RELATION_NAME)';
        SelectSQL.GROUP  := 'A.RDB$USER';
        SelectSQL.ORDER  := '1';
        SQL.Text := SelectSQL.Build;
      end;
    with FB.Query[2] do
      begin
        SelectSQL.SELECT := 'A.RDB$GRANT_OPTION, A.RDB$FIELD_NAME, A.RDB$USER, A.RDB$USER_TYPE, A.RDB$PRIVILEGE, A.RDB$RELATION_NAME';
        SelectSQL.FROM   := 'RDB$USER_PRIVILEGES A LEFT JOIN RDB$RELATIONS B ON (B.RDB$RELATION_NAME = A.RDB$RELATION_NAME)';
        SelectSQL.WHERE  := '(A.RDB$USER <> B.RDB$OWNER_NAME) AND (A.RDB$PRIVILEGE <> ''M'') AND (A.RDB$PRIVILEGE <> ''X'') AND (A.RDB$RELATION_NAME = :_RELATION_NAME) AND (A.RDB$USER = :_USER_NAME)';
        SelectSQL.GROUP  := '';
        SelectSQL.ORDER  := 'A.RDB$PRIVILEGE';
        SQL.Text := SelectSQL.Build;
      end;

    Cnt := 0;
    UpdateFieldList := TStringList.Create;
    ReferrencesFieldList := TStringList.Create;
    try
      for s in Grants do
        begin
          with FB.Query[1] do
            begin
              ParamByName('_RELATION_NAME').AsString := s;
              Open;
              while not EOF do
                begin
                  u := Trim(Fields[0].AsString);
                  with FB.Query[2] do
                    begin
                      ParamByName('_RELATION_NAME').AsString := s;
                      ParamByName('_USER_NAME'    ).AsString := u;
                      Open;
                      PrivTypes := [];
                      GrantOption_Flg := True;
                      UpdateFieldList.Clear;
                      ReferrencesFieldList.Clear;
                      while not EOF do
                        begin
                          if BOF then
                            GrantOption_Flg := (not FieldByName('RDB$GRANT_OPTION').IsNull) and (FieldByName('RDB$GRANT_OPTION').AsInteger = 1);
                          Privilege := FieldByName('RDB$PRIVILEGE').AsString + ' ';
                          case Ord(Privilege[1]) of
                            Ord('S'):
                              Include(PrivTypes, privSelect);
                            Ord('D'):
                              Include(PrivTypes, privDelete);
                            Ord('I'):
                              Include(PrivTypes, privInsert);
                            Ord('U'):
                              begin
                                Include(PrivTypes, privUpdate);
                                if not FieldByName('RDB$FIELD_NAME').IsNull then
                                  UpdateFieldList.Add(Trim(FieldByName('RDB$FIELD_NAME').AsString));
                              end;
                            Ord('R'):
                              begin
                                Include(PrivTypes, privReferences);
                                if not FieldByName('RDB$FIELD_NAME').IsNull then
                                  ReferrencesFieldList.Add(Trim(FieldByName('RDB$FIELD_NAME').AsString));
                              end;
                          end;
                          Next;
                        end;
                      Close;
                    end;
                  PrivilegeStr := '';
                  if ([privSelect, privDelete, privInsert, privUpdate, privReferences] = PrivTypes) and
                     (UpdateFieldList.Count = 0) and (ReferrencesFieldList.Count = 0) then
                    PrivilegeStr := 'ALL'
                  else
                    begin
                      if privSelect in PrivTypes then
                        PrivilegeStr := PrivilegeStr + 'SELECT, ';
                      if privDelete in PrivTypes then
                        PrivilegeStr := PrivilegeStr + 'DELETE, ';
                      if privInsert in PrivTypes then
                        PrivilegeStr := PrivilegeStr + 'INSERT, ';
                      if privUpdate in PrivTypes then
                        PrivilegeStr := PrivilegeStr + Format('UPDATE%s, ', [BuildPrivilegeFieldString(UpdateFieldList)]);
                      if privReferences in PrivTypes then
                        PrivilegeStr := PrivilegeStr + Format('REFERENCES%s, ', [BuildPrivilegeFieldString(ReferrencesFieldList)]);
                      PrivilegeStr := Trim(PrivilegeStr);
                      if Copy(PrivilegeStr, Length(PrivilegeStr), 1) = ',' then
                        PrivilegeStr := Copy(PrivilegeStr, 1, Length(PrivilegeStr)-1);
                    end;
                  retGrants[aIndex + Cnt] := Format('GRANT %s ON "%s" TO %s%s%s', [PrivilegeStr, s, u, GRAND_OPTION_STR[GrantOption_Flg], TERM_CHAR]);
                  Inc(Cnt);
                  Next;
                end;
              Close;
            end;
        end;
    finally
      ReferrencesFieldList.Free;
      UpdateFieldList.Free;
    end;
    Result := Cnt;
  end;
  { BuildGrant_Object END }
begin
  Index := 0;

  // テーブル
  Grants := EnumGrants(False);
  SetLength(Result, Length(Grants));
  Index := Index + BuildGrant_Object(Grants, Result, Index);
  SetLength(Result, Index);

  // ビュー
  Grants := EnumGrants(True);
  SetLength(Result, Length(Result) + Length(Grants));
  Index := Index + BuildGrant_Object(Grants, Result, Index);
  SetLength(Result, Index);
end;

function TFBExtractor.EnumAndBuildGrant_Role: TStringDynArray;
// GRANT (ロール) の生成
  const
    GRAND_OPTION_STR: array [Boolean] of string = ('', ' WITH ADMIN OPTION');
var
  SelectSQL: TSelectSQL;
  Cnt, RecCnt: Integer;
begin
  with FB.Query[1] do
    begin
      // GRANT の件数を得る
      SelectSQL.SELECT := 'Count(*)';
      SelectSQL.FROM   := 'RDB$USER_PRIVILEGES';
      SelectSQL.WHERE  := '(RDB$OBJECT_TYPE = 13) AND (RDB$USER_TYPE = 8) AND (RDB$PRIVILEGE = ''M'')';
      SQL.Text := SelectSQL.Build;
      Open;
      RecCnt := Fields[0].AsInteger;
      Close;

      // 戻り値の配列のサイズを変更
      SetLength(Result, RecCnt);

      // GRANT を列挙
      if RecCnt > 0 then
        begin
          SelectSQL.SELECT := 'RDB$RELATION_NAME, RDB$USER, RDB$GRANT_OPTION';
          SelectSQL.ORDER  := 'RDB$RELATION_NAME, RDB$USER';
          SQL.Text := SelectSQL.Build;
          Open;
          Cnt := 0;
          while not EOF do
            begin
              Result[Cnt] := Format('GRANT %s TO %s%s%s',
                               [QuotedString(Trim(FieldByName('RDB$RELATION_NAME').AsString)),
                                Trim(FieldByName('RDB$USER').AsString),
                                GRAND_OPTION_STR[FieldByName('RDB$GRANT_OPTION').AsInteger = 1],
                                TERM_CHAR]);
              Inc(Cnt);
              Next;
            end;
          Close;
        end;
    end;
end;

function TFBExtractor.BuildIndex(Index: TIndexRec): string;
// インデックスの生成
const
  SelectSQL: TSelectSQL = (SELECT: 'RDB$FIELD_NAME'; FROM: 'RDB$INDEX_SEGMENTS'; WHERE: 'RDB$INDEX_NAME = :_INDEX_NAME'; ORDER: 'RDB$FIELD_POSITION');
  SortStr: array [Boolean] of string = ('', ' DESCENDING');
  UniqueStr: array [Boolean] of string = ('', ' UNIQUE');
var
  dFields: string;
begin
  with FB.Query[1] do
    begin
      SQL.Text := SelectSQL.Build;
      ParamByName('_INDEX_NAME').AsString := Index.Name;
      Open;
      dFields := '';
      while not EOF do
        begin
          dFields := dFields + QuotedString(Trim(Fields[0].AsString));
          Next;
          if not EOF then
            dFields := dFields + SEPARATOR_STR;
        end;
      Close;
    end;
  Result := Format('CREATE%S%S INDEX %s ON %s (%s)%s',
              [UniqueStr[Index.Unique], SortStr[Index.IndexType = 1], QuotedString(Index.Name), QuotedString(Index.TableName), dFields, TERM_CHAR]);
end;

function TFBExtractor.BuildProcedure(aName: string; aDefine: Boolean): string;
// ストアドプロシージャの生成
const
  ProcString: array [Boolean] of string = ('ALTER', 'CREATE');
var
  Argument, CharSetStr, ProcSource: string;
  SelectSQL: TSelectSQL;
begin
  Result := Format('%s PROCEDURE %s', [ProcString[aDefine], QuotedString(aName)]) + sLineBreak;
  // プロシージャ引数
  with FB.Query[0] do
    begin
      SelectSQL.SELECT := '*';
      SelectSQL.FROM   := 'RDB$PROCEDURE_PARAMETERS A LEFT JOIN RDB$FIELDS B ON (B.RDB$FIELD_NAME = A.RDB$FIELD_SOURCE)';
      SelectSQL.WHERE  := '(A.RDB$PROCEDURE_NAME = :_PROCEDURE_NAME) AND (A.RDB$PARAMETER_TYPE = :_PARAMETER_TYPE)';
      SelectSQL.ORDER  := 'A.RDB$PARAMETER_NUMBER';
      SQL.Text := SelectSQL.Build;
      ParamByName('_PROCEDURE_NAME').AsString  := aName;
      ParamByName('_PARAMETER_TYPE').AsInteger := 0; // Arguments
      Open;
      while not EOF do
        begin
          if BOF then
            Result := Result + '(' + sLineBreak;
          Argument := '';
          Argument := Argument + Format('  %s %s', [QuotedString(Trim(FieldByName('RDB$PARAMETER_NAME').AsString)), BuildArgumentString(FB.Query[0])]);
          Result := Result + Argument;
          Next;
          if EOF then
            Result := Result + sLineBreak + ')' + sLineBreak
          else
            Result := Result + ',' + sLineBreak
        end;
      Close;
    end;
  // プロシージャ戻り値
  with FB.Query[0] do
    begin
      ParamByName('_PROCEDURE_NAME').AsString  := aName;
      ParamByName('_PARAMETER_TYPE').AsInteger := 1; // Returns
      Open;
      while not EOF do
        begin
          if BOF then
            begin
              Result := Result + 'RETURNS' + sLineBreak;
              Result := Result + '(' + sLineBreak;
            end;
          Argument := '';
          Argument := Argument + Format('  %s %s', [QuotedString(Trim(FieldByName('RDB$PARAMETER_NAME').AsString)), BuildArgumentString(FB.Query[0])]);
          Result := Result + Argument;
          Next;
          if EOF then
            Result := Result + sLineBreak + ')' + sLineBreak
          else
            Result := Result + ',' + sLineBreak
        end;
      Close;
    end;
  // AS
  Result := Result + 'AS' + sLineBreak;
  // プロシージャソース
  if aDefine then
    begin
      Result := Result + 'BEGIN' + sLineBreak;
      Result := Result + #$09 + 'SUSPEND' + TERM_CHAR + sLineBreak;
      Result := Result + 'END' + sLineBreak;
    end
  else
    begin
      with FB.Query[0] do
        begin
          SelectSQL.SELECT := '*';
          SelectSQL.FROM   := 'RDB$PROCEDURES';
          SelectSQL.WHERE  := '(RDB$PROCEDURE_NAME = :_PROCEDURE_NAME)';
          SelectSQL.ORDER  := '';
          SQL.Text := SelectSQL.Build;
          ParamByName('_PROCEDURE_NAME').AsString := aName;
          Open;
          if not IsEmpty then
            begin
              ProcSource := FieldByName('RDB$PROCEDURE_SOURCE').AsString;
              // デフォルト文字コードと同じ文字コードの定義が含まれたらその部分を削除する
              CharsetStr := Format(' *CHARACTER +SET +%s', [Self.DefaultCharset.Name]);
              ProcSource := TRegEx.Replace(ProcSource, CharSetStr, '', [roIgnoreCase]);
              Result := Result + ProcSource + sLineBreak;
            end;
          Close;
        end;
    end;
end;

function TFBExtractor.BuildRole(aName: string): string;
// ロールの生成
begin
  Result := Format('CREATE ROLE %s%s', [QuotedString(aName), TERM_CHAR]);
end;

function TFBExtractor.BuildTable(aName: string): string;
// テーブルの生成
var
  EXTERNAL_FILE, FIELD_NAME, FIELD_TYPE,
  Check, Constraint, FieldStr, PrimaryKey, Unique: string;
  HasConstraint, HasPrimaryKey, HasUnique, HasCheck, IsDomain: Boolean;
  SelectSQL: TSelectSQL;
  TableConstraintCount: Integer;
  SL: TStringList;
  i: Integer;
begin
  Result := Format('CREATE TABLE %s%s', [QuotedString(aName), sLineBreak]);
  // 外部ファイル
  with FB.Query[0] do
    begin
      SelectSQL.SELECT := '*'; 
      SelectSQL.FROM   := 'RDB$RELATIONS'; 
      SelectSQL.WHERE  := '(RDB$RELATION_NAME = :_RELATION_NAME)';
      SQL.Text := SelectSQL.Build;
      ParamByName('_RELATION_NAME').AsString := aName;
      Open;
      if not IsEmpty then
        begin
          if not FieldByName('RDB$EXTERNAL_FILE').IsNULL then
            begin
              EXTERNAL_FILE := FieldByName('RDB$EXTERNAL_FILE').AsString;
              Result := Result + Format(' EXTERNAL FILE ''%s''%s', [EXTERNAL_FILE, sLineBreak]);
            end;
        end;
      Close;
    end;
  Result := Result + '(' + sLineBreak;
  // プライマリキー (表制約)
  PrimaryKey := BuildPrimaryKey(aName);
  HasPrimaryKey := (PrimaryKey <> '');
  // プライマリキー (表制約)
  Unique := BuildUnique(aName);
  HasUnique := (Unique <> '');
  // 制約 (表制約)
  Constraint := BuildConstraint(aName);
  HasConstraint := (Constraint <> '');
  // CHECK 制約 (表制約)
  Check := BuildCheck(aName);
  HasCheck      := (Check <> '');
  // 表制約が存在するか？
  TableConstraintCount := Integer(HasPrimaryKey) +  Integer(HasUnique) + Integer(HasConstraint) + Integer(HasCheck);
  // フィールド
  with FB.Query[0] do
    begin
      SelectSQL.SELECT := '*';
      SelectSQL.FROM   := 'RDB$RELATION_FIELDS'; 
      SelectSQL.WHERE  := '(RDB$RELATION_NAME = :_RELATION_NAME)';
      SelectSQL.ORDER  := 'RDB$FIELD_POSITION'; // <> FIELD_ID
      SQL.Text := SelectSQL.Build;
      ParamByName('_RELATION_NAME').AsString := aName;
      Open;
      while not EOF do
        begin
          // フィールド
          FieldStr := '';
          IsDomain := (Pos('RDB$', FieldByName('RDB$FIELD_SOURCE').AsString) <> 1);
          FIELD_NAME := Trim(FieldByName('RDB$FIELD_SOURCE').AsString);
          FIELD_TYPE := BuildFieldString(FIELD_NAME, IsDomain);
          FIELD_NAME := Trim(FieldByName('RDB$FIELD_NAME').AsString);
          FieldStr := Format('  %s %s', [QuotedString(FIELD_NAME), FIELD_TYPE]);
          // DEFAULT
          if not FieldByName('RDB$DEFAULT_VALUE').IsNull then
            FieldStr := FieldStr + ' ' + FieldByName('RDB$DEFAULT_SOURCE').AsString;  
          // NOT NULL
          if FieldByName('RDB$NULL_FLAG').AsInteger = 1 then
            FieldStr := FieldStr + ' NOT NULL';
          // 列制約
          // COLATE
          Next;
          if (not EOF) or (TableConstraintCount > 0) then
            FieldStr := FieldStr + ',';
          Result := Result + FieldStr + sLineBreak;
        end;
      Close;
    end;
  // プライマリキー (表制約)
  if HasPrimaryKey then
    begin
      Dec(TableConstraintCount);
      Result := Result + '  ' + PrimaryKey;
      if TableConstraintCount > 0 then
        Result := Result + ',';
      Result := Result + sLineBreak;
    end;
  // Unique 制約 (表制約)
  if HasUnique then
    begin
      Dec(TableConstraintCount);
      Result := Result + '  ' + Unique;
      if TableConstraintCount > 0 then
        Result := Result + ',';
      Result := Result + sLineBreak;
    end;
  // Constraint (表制約)
  if HasConstraint then
    begin
      Dec(TableConstraintCount);
      SL := TStringList.Create;
      try
        SL.Text := Constraint;
        for i:=0 to SL.Count-1 do
          begin
            Result := Result + '  ' + SL[i];
            if i < SL.Count-1 then
              Result := Result + sLineBreak;
          end;
      finally
        SL.Free;
      end;
      if TableConstraintCount > 0 then
        Result := Result + ',';
      Result := Result + sLineBreak;
    end;
  // Check 制約 (表制約)
  if HasCheck then
    begin
      Dec(TableConstraintCount);
      Result := Result + '  ' + Check;
      if TableConstraintCount > 0 then
        Result := Result + ',';
      Result := Result + sLineBreak;
    end;
  Result := Result + ')' + TERM_CHAR;
end;

class function TFBExtractor.PreBuildTableData(aExtractor: TFBExtractor; TableName: string; SplitSize: Integer): TSplitDataRec;
// テーブルデータの生成
var
  SelectSQL: TSelectSQL;
begin
  with aExtractor.Firebird.Query[0] do
    begin
      // レコード件数を得る
      SelectSQL.SELECT := 'Count(*)';
      SelectSQL.FROM   := TableName;
      SQL.Text := SelectSQL.Build;
      Open;
      Result.Index       := 0;
      Result.RecordCount := Fields[0].AsInteger;
      Result.SplitSize   := SplitSize;
      Result.SplitCount  := (Max(0, Result.RecordCount - 1)) div SplitSize + 1;
      Result.LastSize    := Result.RecordCount mod SplitSize;
      Close;
    end;
end;

class function TFBExtractor.BuildTableData(aExtractor: TFBExtractor; TableName: string; TableIndex: Integer; var SplitDataRec: TSplitDataRec): TInsRecDynArray;
// テーブルデータの生成
var
  BlobFileName, dFields, dComment, Dmy, Dmy2, dValues, dValues2, TextFileName: string;
  C: Char;
  Cnt: Int64;
  HasControlChar: Boolean;
  i, l: Integer;
  LineBreak: Boolean;
  SelectSQL: TSelectSQL;
  Stream: TMemoryStream;
begin
  with aExtractor.Firebird.Query[0] do
    begin
      // 配列のサイズを変更
      if (SplitDataRec.Index = SplitDataRec.SplitCount - 1) then
        SetLength(Result, SplitDataRec.LastSize)
      else
        SetLength(Result, SplitDataRec.SplitSize);

      // データ処理
      SelectSQL.SELECT := Format('FIRST %d SKIP %d *', [SplitDataRec.SplitSize, SplitDataRec.Index * SplitDataRec.SplitSize]);
      SelectSQL.FROM   := TableName;
      SQL.Text := SelectSQL.Build;
      Open;
      Cnt := 0;
      while not EOF do
        begin
          dFields  := '';
          dValues  := '';
          dValues2 := '';
          dComment := '';
          for i:=0 to Fields.Count-1 do
            begin
              HasControlChar := False;
              dFields := dFields + QuotedString(Trim(Fields[i].FieldName));
              if Fields[i].IsNull then
                begin
                  dValues  := dValues  + NULL_STRING;
                  dValues2 := dValues2 + NULL_STRING;
                end
              else
                begin
                  case Fields[i].DataType of
                    ftMemo, ftWideMemo,
                    ftString, ftWideString,
                    ftDate, ftTime, ftDateTime:
                      begin
                        Dmy := Fields[i].AsString;
                        Dmy := StringReplace(Dmy, '''', '''''', [rfReplaceAll]);
                        Dmy := QuotedStr(Dmy);
                        // コントロール文字が含まれるか？
                        for C in Dmy do
                          begin
                            if C < #$0020 then
                              begin
                                HasControlChar := True;
                                Break;
                              end;
                          end;
                        // メタデータ用 SQL 文字列の生成
                        if HasControlChar then
                          begin
                            // コントロール文字が含まれていれば ansi_char() で置換
                            Dmy2 := '';
                            LineBreak := False;
                            for l:=1 to Length(Dmy) do
                              begin
                                if Dmy[l] < #$0020 then
                                  begin
                                    if not LineBreak then
                                      Dmy2 := Dmy2 + '''';
                                    Dmy2 := Dmy2 + Format(' || ascii_char(%d)', [Ord(Dmy[l])]);
                                    LineBreak := True;
                                  end
                                else
                                  begin
                                    if LineBreak and (l > 2) then
                                      Dmy2 := Dmy2 + ' || ''';
                                    Dmy2 := Dmy2 + Dmy[l];
                                    LineBreak := False;
                                  end;
                              end;
                            dValues  := dValues + Dmy2;
                          end
                        else
                          begin
                            // コントロール文字が含まれていない
                            dValues  := dValues + Dmy;
                          end;
                        // 実行用 SQL 文字列の生成
                        if HasControlChar then
                          begin
                            // コントロール文字が含まれていればパラメータに変換
                            Dmy := Format('%.4x%.4x%.8x', [TableIndex, i, Cnt]);
                            dValues2 := dValues2 + ':PT' + Dmy;
                            TextFileName := Dmy + '.text';
                            dComment := dComment + Format('%.4d: ', [i]) + TextFileName;

                            StringToFile(Fields[i].AsString, TPath.Combine(aExtractor.BrobOutputDir, TextFileName));
                          end
                        else
                          begin
                            // コントロール文字が含まれていない
                            dValues2 := dValues2 + Dmy;
                          end;
                      end;
                    ftBlob:
                      begin
                        dValues := dValues + NULL_STRING;
                        Dmy := Format('%.4x%.4x%.8x', [TableIndex, i, Cnt]);
                        dValues2 := dValues2 + ':PB' + Dmy;
                        BlobFileName := Dmy + '.blob';
                        dComment := dComment + Format('%.4d: ', [i]) + BlobFileName;
                        if aExtractor.IsExtractBlobData then
                          begin
                            Stream := TMemoryStream.Create;
                            try
                              aExtractor.Firebird.LoadFromBlob(Stream, TBlobField(Fields[i]));
                              Stream.SaveToFile(TPath.Combine(aExtractor.BrobOutputDir, BlobFileName));
                            finally
                              Stream.Free;
                            end;
                          end;
                      end
                  else
                    begin
                      dValues  := dValues  + Fields[i].AsString;
                      dValues2 := dValues2 + Fields[i].AsString;
                    end;
                  end;
                end;
              if i < Fields.Count-1 then
                begin
                  dFields  := dFields  + SEPARATOR_STR;
                  dValues  := dValues  + SEPARATOR_STR;
                  dValues2 := dValues2 + SEPARATOR_STR;
                  if (Fields[i].DataType = ftBlob) or HasControlChar then
                    dComment := dComment + SEPARATOR_STR;
                end;
            end;
          dComment := Trim(dComment);
          if (dComment <> '') and (Copy(dComment, Length(dComment), 1) = ',') then
            dComment := Copy(dComment, 1, Length(dComment)-1);
          Result[Cnt].Fields  := dFields;
          Result[Cnt].Values1 := dValues;
          Result[Cnt].Values2 := dValues2;
          Result[Cnt].Comment := dComment;
          Inc(Cnt);
          Next;
        end;
      Close;
    end;
  SplitDataRec.Index := SplitDataRec.Index + 1;
end;

function TFBExtractor.BuildTrigger(Trigger: TTriggerRec): string;
// トリガの生成
var
  CharSetStr, ProcSource: string;
const
  TriggerActive: array [Boolean] of string = ('INACTIVE', 'ACTIVE');
  TriggerPos: array [Boolean] of string = ('AFTER', 'BEFORE');
  TriggerAction: array [0..2] of string = ('INSERT', 'UPDATE', 'DELETE');

  { GetTriggerTypeString BEGIN }
  function GetTriggerTypeString(TriggerType: Integer): string;
  begin
    Result := '';
    if TriggerType = 0 then
      Exit;
    Result := Format('%s %s', [TriggerPos[(TriggerType mod 2) = 1], TriggerAction[(TriggerType div 3)]]);
  end;
  { GetTriggerTypeString END }
begin
  Result := Format('CREATE TRIGGER %s FOR %s ',
              [QuotedString(Trigger.Name),
               QuotedString(Trigger.TableName)]) + sLineBreak;
  Result := Result + Format('%s %s POSITION %d ',
              [TriggerActive[Trigger.Active],
               GetTriggerTypeString(Trigger.TriggerType),
               Trigger.Sequence]) + sLineBreak;
  ProcSource := Trigger.Source;
  // デフォルト文字コードと同じ文字コードの定義が含まれたらその部分を削除する
  CharsetStr := Format(' *CHARACTER +SET +%s', [Self.DefaultCharset.Name]);
  ProcSource := TRegEx.Replace(ProcSource, CharSetStr, '', [roIgnoreCase]);
  Result := Result + ProcSource;
end;

function TFBExtractor.BuildView(aName: string): string;
// ビューの生成
var
  Argument: string;
  SelectSQL: TSelectSQL;
begin
  // 引数
  Argument := '';
  with FB.Query[0] do
    begin
      SelectSQL.SELECT := '*';
      SelectSQL.FROM   := 'RDB$RELATION_FIELDS';
      SelectSQL.WHERE  := '(RDB$RELATION_NAME = :_RELATION_NAME)';
      SelectSQL.ORDER  := 'RDB$FIELD_POSITION';
      SQL.Text := SelectSQL.Build;
      ParamByName('_RELATION_NAME').AsString := aName;
      Open;
      while not EOF do
        begin
          Argument := Argument + QuotedString(Trim(FieldByName('RDB$FIELD_NAME').AsString));
          Next;
          if not EOF then
            Argument := Argument + SEPARATOR_STR;
        end;
      Close;
    end;
  Result := Format('CREATE VIEW %s (%s) AS', [QuotedString(aName), Argument]) + sLineBreak;
  // ソース
  with FB.Query[0] do
    begin
      SelectSQL.SELECT := '*';
      SelectSQL.FROM   := 'RDB$RELATIONS';
      SelectSQL.WHERE  := '(RDB$RELATION_NAME = :_RELATION_NAME)';
      SelectSQL.ORDER  := '';
      SQL.Text := SelectSQL.Build;
      ParamByName('_RELATION_NAME').AsString := aName;
      Open;
      if not IsEmpty then
        Result := Result + Trim(FieldByName('RDB$VIEW_SOURCE').AsString);
      Close;
    end;
  if (Length(Result) > 0) and (Result[Length(Result)] <> TERM_CHAR) then
    Result := Result + TERM_CHAR;
end;

function TFBExtractor.BuildUniqueKey(TableName: string; IsPrimary: Boolean): string;
// ユニークキー (表制約) の生成
var
  Cnt, RecCnt: Integer;
  IsUnique: Boolean;
  KeyFields: string;
  SelectSQL: TSelectSQL;
begin
  KeyFields := '';
  with FB.Query[1] do
    begin
      SelectSQL.SELECT := 'B.RDB$FIELD_NAME, A.RDB$CONSTRAINT_TYPE, A.RDB$CONSTRAINT_NAME';
      SelectSQL.FROM   := 'RDB$RELATION_CONSTRAINTS A LEFT JOIN RDB$INDEX_SEGMENTS B ON (B.RDB$INDEX_NAME = A.RDB$INDEX_NAME)';
      SelectSQL.WHERE  := '(A.RDB$CONSTRAINT_TYPE = :_CONSTRAINT_TYPE) AND (A.RDB$RELATION_NAME = :_RELATION_NAME)';
      SelectSQL.ORDER  := 'RDB$FIELD_POSITION';
      SQL.Text := SelectSQL.Build;
      ParamByName('_RELATION_NAME').AsString := TableName;
      ParamByName('_CONSTRAINT_TYPE').AsString := KeyString[IsPrimary];
      Open;
      RecCnt := 0;
      while not EOF do
        begin
          IsUnique := (Pos('INTEG_', FieldByName('RDB$CONSTRAINT_NAME').AsString) = 1);
          if IsUnique then
            Inc(RecCnt);
          Next;
        end;
      First;
      Cnt := 0;
      while not EOF do
        begin
          IsUnique := (Pos('INTEG_', FieldByName('RDB$CONSTRAINT_NAME').AsString) = 1);
          if IsUnique then
            begin
              KeyFields := KeyFields + QuotedString(Trim(FieldByName('RDB$FIELD_NAME').AsString));
              Inc(Cnt);
            end;
          Next;
          if IsUnique and (Cnt < RecCnt) then
            KeyFields := KeyFields + SEPARATOR_STR;
        end;
      Close;
    end;
  if KeyFields <> '' then
    KeyFields := Format('%s (%s)', [KeyString[IsPrimary], KeyFields]);
  Result := KeyFields;
end;

function TFBExtractor.BuildPrimaryKey(TableName: string): string;
// プライマリキー (表制約) の生成
begin
  Result := BuildUniqueKey(TableName, True);
end;

function TFBExtractor.BuildUnique(TableName: string): string;
// ユニーク (表制約) の生成
begin
  Result := BuildUniqueKey(TableName, False);
end;

function TFBExtractor.BuildConstraint(TableName: string): string;
// Constraint (表制約) の生成
var
  SL: TStringList;
  PrimaryKeyFields, UniqueKeyFields: string;
  { _BuildConstraint BEGIN }
  function _BuildConstraint(TableName: string; IsPrimary: Boolean): string;
  var
    ConstraintName, FieldNames, KeyFields: string;
    i: Integer;
    SelectSQL: TSelectSQL;
  begin
    // Constraint 名を取得
    KeyFields := '';
    SL.Clear;
    with FB.Query[1] do
      begin
        SelectSQL.SELECT := 'A.RDB$CONSTRAINT_NAME';
        SelectSQL.FROM   := 'RDB$RELATION_CONSTRAINTS A LEFT JOIN RDB$INDEX_SEGMENTS B ON (B.RDB$INDEX_NAME = A.RDB$INDEX_NAME)';
        SelectSQL.WHERE  := '(A.RDB$CONSTRAINT_TYPE = :_CONSTRAINT_TYPE) AND (A.RDB$RELATION_NAME = :_RELATION_NAME)';
        SelectSQL.GROUP  := 'A.RDB$CONSTRAINT_NAME';
        SelectSQL.ORDER  := '';
        SQL.Text := SelectSQL.Build;
        ParamByName('_RELATION_NAME').AsString := TableName;
        ParamByName('_CONSTRAINT_TYPE').AsString := KeyString[IsPrimary];
        Open;
        while not EOF do
          begin
            if (Pos('INTEG_', Fields[0].AsString) <> 1) then
              SL.Add(Trim(Fields[0].AsString));
            Next;
          end;
        Close;
      end;
    // Constraint を列挙
    with FB.Query[1] do
      begin
        SelectSQL.SELECT := 'B.RDB$FIELD_NAME, A.RDB$CONSTRAINT_TYPE, A.RDB$CONSTRAINT_NAME';
//      SelectSQL.FROM   := 'RDB$RELATION_CONSTRAINTS A LEFT JOIN RDB$INDEX_SEGMENTS B ON (B.RDB$INDEX_NAME = A.RDB$INDEX_NAME)';
        SelectSQL.WHERE  := '(A.RDB$CONSTRAINT_NAME = :_CONSTRAINT_NAME) AND (A.RDB$CONSTRAINT_TYPE = :_CONSTRAINT_TYPE) AND (A.RDB$RELATION_NAME = :_RELATION_NAME)';
        SelectSQL.GROUP  := '';
        SelectSQL.ORDER  := 'B.RDB$FIELD_POSITION';
        SQL.Text := SelectSQL.Build;
      end;
    for i:=0 to SL.Count-1 do
      begin
        FieldNames := '';
        with FB.Query[1] do
          begin
            ParamByName('_CONSTRAINT_NAME').AsString := SL[i];
            ParamByName('_RELATION_NAME').AsString := TableName;
            ParamByName('_CONSTRAINT_TYPE').AsString := KeyString[IsPrimary];
            Open;
            while not EOF do
              begin
                FieldNames := FieldNames + QuotedString(Trim(FieldByName('RDB$FIELD_NAME').AsString));
                ConstraintName := QuotedString(FieldByName('RDB$CONSTRAINT_NAME').AsString);
                Next;
                if not EOF then
                  FieldNames := FieldNames + SEPARATOR_STR;
              end;
            Close;
          end;
        if FieldNames <> '' then
          KeyFields := KeyFields + Format('CONSTRAINT %s %s (%s)', [ConstraintName, KeyString[IsPrimary], FieldNames]);
        if i < (SL.Count-1) then
          KeyFields := KeyFields + ',' + sLineBreak;
      end;
    result := KeyFields;
  end;
  { _BuildConstraint END }
begin
  PrimaryKeyFields := '';
  UniqueKeyFields  := '';
  SL := TStringList.Create;
  try
    // CONSTRAINT PRIMARY
    PrimaryKeyFields := _BuildConstraint(TableName, True);
    // CONSTRAINT UNIQUE
    UniqueKeyFields  := _BuildConstraint(TableName, False);
  finally
    SL.Free;
  end;
  Result := PrimaryKeyFields;
  if (PrimaryKeyFields <> '') and (UniqueKeyFields <> '') then
    Result := Result + ',' + sLineBreak;
  Result := Result + UniqueKeyFields;
end;

function TFBExtractor.BuildForeignKey(TableName: string): string;
// 外部キー (表制約) の生成 : ALTER TABLE で表現
var
  DeleteRule, ForeignKeys, RefIdx, RefTable, RefKeys, UpdateRule, Constraint: string;
  SelectSQL: TSelectSQL;
begin
  Result := '';

  SelectSQL.SELECT := 'Max(C.RDB$CONST_NAME_UQ), Max(C.RDB$CONSTRAINT_NAME)';
  SelectSQL.FROM   := 'RDB$RELATION_CONSTRAINTS A ' +
                      'LEFT JOIN RDB$INDEX_SEGMENTS B ON ' + 
                      '(B.RDB$INDEX_NAME = A.RDB$INDEX_NAME) ' +
                      'LEFT JOIN RDB$REF_CONSTRAINTS C ON ' + 
                      '(C.RDB$CONSTRAINT_NAME = A.RDB$CONSTRAINT_NAME)';
  SelectSQL.WHERE  := '(A.RDB$RELATION_NAME = :_RELATION_NAME) AND (A.RDB$CONSTRAINT_TYPE = :_CONSTRAINT_TYPE)';
  SelectSQL.GROUP  := 'C.RDB$CONST_NAME_UQ';
  FB.Query[0].SQL.Text := SelectSQL.Build;

  SelectSQL.SELECT := 'A.RDB$CONSTRAINT_TYPE, A.RDB$RELATION_NAME, B.RDB$FIELD_NAME, C.RDB$CONST_NAME_UQ, C.RDB$UPDATE_RULE, C.RDB$DELETE_RULE';
  SelectSQL.FROM   := 'RDB$RELATION_CONSTRAINTS A ' +
                      'LEFT JOIN RDB$INDEX_SEGMENTS B ON ' + 
                      '(B.RDB$INDEX_NAME = A.RDB$INDEX_NAME) ' + 
                      'LEFT JOIN RDB$REF_CONSTRAINTS C ON ' +
                      '(C.RDB$CONSTRAINT_NAME = A.RDB$CONSTRAINT_NAME)';
  SelectSQL.WHERE  := '(A.RDB$RELATION_NAME = :_RELATION_NAME) AND (A.RDB$CONSTRAINT_TYPE = :_CONSTRAINT_TYPE) AND (C.RDB$CONST_NAME_UQ = :_CONST_NAME_UQ)';
  SelectSQL.GROUP  := '';
  SelectSQL.ORDER  := 'B.RDB$FIELD_POSITION';
  FB.Query[1].SQL.Text := SelectSQL.Build;

  SelectSQL.SELECT := 'A.RDB$CONSTRAINT_TYPE, A.RDB$RELATION_NAME, B.RDB$FIELD_NAME';
  SelectSQL.FROM   := 'RDB$RELATION_CONSTRAINTS A ' + 
                      'LEFT JOIN RDB$INDEX_SEGMENTS B ON ' + 
                      '(B.RDB$INDEX_NAME = A.RDB$INDEX_NAME)'; 
  SelectSQL.WHERE  := '(A.RDB$CONSTRAINT_NAME = :_CONSTRAINT_NAME)';
  SelectSQL.GROUP  := '';
  SelectSQL.ORDER  := 'B.RDB$FIELD_POSITION';
  FB.Query[2].SQL.Text := SelectSQL.Build;

  with FB.Query[0] do
    begin
      ParamByName('_RELATION_NAME').AsString := TableName;
      ParamByName('_CONSTRAINT_TYPE').AsString := 'FOREIGN KEY';
      Open;
      while not EOF  do
        begin                                                                             
          ForeignKeys := '';
          RefKeys     := '';
          RefIdx      := FB.Query[0].Fields[0].AsString;
          Constraint  := Trim(FB.Query[0].Fields[1].AsString);
          if Pos('INTEG_', Constraint) = 1 then
            Constraint := ''
          else
            Constraint := 'CONSTRAINT ' + QuotedString(Constraint) + ' ';
          // 外部キーの列挙
          with FB.Query[1] do
            begin
              ParamByName('_RELATION_NAME').AsString   := TableName;
              ParamByName('_CONSTRAINT_TYPE').AsString := 'FOREIGN KEY';
              ParamByName('_CONST_NAME_UQ').AsString   := RefIdx;
              Open;
              while not EOF  do
                begin                                                                             
                  ForeignKeys := ForeignKeys + QuotedString(Trim(FieldByName('RDB$FIELD_NAME').AsString));
                  UpdateRule := Trim(FieldByName('RDB$UPDATE_RULE').AsString);
                  DeleteRule := Trim(FieldByName('RDB$DELETE_RULE').AsString);
                  Next;
                  if not EOF then
                    ForeignKeys := ForeignKeys + SEPARATOR_STR;
                end;
              Close;
            end;
          // 参照元のキーを列挙
          with FB.Query[2] do
            begin
              ParamByName('_CONSTRAINT_NAME').AsString := RefIdx;
              RefKeys := '';
              Open;
              while not EOF  do
                begin
                  RefTable := Trim(FieldByName('RDB$RELATION_NAME').AsString);
                  RefKeys := RefKeys + QuotedString(Trim(FieldByName('RDB$FIELD_NAME').AsString));
                  Next;
                  if not EOF then
                    RefKeys := RefKeys + SEPARATOR_STR;
                end;
              Close;
            end;
          Result := Result + Format('ALTER TABLE %s ADD %sFOREIGN KEY (%s) REFERENCES %s (%s)',
            [QuotedString(TableName), Constraint, ForeignKeys, QuotedString(RefTable), RefKeys]);
          if (DeleteRule <> '') and (DeleteRule <> 'RESTRICT') then
            Result := Result + Format(' ON DELETE %s', [DeleteRule]);
          if (UpdateRule <> '') and  (UpdateRule <> 'RESTRICT') then
            Result := Result + Format(' ON UPDATE %s', [UpdateRule]);
          Result := Result + TERM_CHAR;
          Next;
          if not EOF then
            Result := Result + sLineBreak;
        end;
      Close;
    end;
end;

function TFBExtractor.BuildCheck(TableName: string): string;
// CHECK 制約 (表制約) の生成
begin
  Result := '';
end;

function TFBExtractor.BuildEnumObjects(FIELD, FROM, WHERE, ORDER: string): TStringDynArray;
// オブジェクトの列挙
var
  Cnt, RecCnt: Integer;
  SelectSQL: TSelectSQL;
  IgnoreSystem: Boolean;
begin
  with FB.Query[0] do
    begin
      SelectSQL.Init;
      IgnoreSystem := True;
      if Pos('RDB$ROLES', FROM) > 0 then
        begin
          SelectSQL.SELECT := '*';
          SelectSQL.FROM   := FROM;
          SQL.Text := SelectSQL.Build;
          Open;
          IgnoreSystem := (Fields.FindField('RDB$SYSTEM_FLAG') <> nil);
          Close;
        end;
      SelectSQL.SELECT := 'Count(*)';
      SelectSQL.FROM   := FROM;
      if IgnoreSystem then
        SelectSQL.WHERE  := '((RDB$SYSTEM_FLAG = 0) OR (RDB$SYSTEM_FLAG is NULL))';
      if IgnoreSystem and (WHERE <> '') then
        SelectSQL.WHERE := SelectSQL.WHERE + ' AND ';
      if WHERE <> '' then
        SelectSQL.WHERE := SelectSQL.WHERE + WHERE;
      // レコード数を得る
      SQL.Text := SelectSQL.Build;
      Open;
      RecCnt := Fields[0].AsInteger;
      Close;
      // 戻り値のサイズを変更
      SetLength(Result, RecCnt);
      // レコードを取得
      SelectSQL.SELECT := FIELD;
      if ORDER <> '' then
        SelectSQL.ORDER := ORDER
      else
        SelectSQL.ORDER  := FIELD;
      SQL.Text := SelectSQL.Build;
      Open;
      Cnt := 0;
      while not EOF do
        begin
          Result[Cnt] := Trim(Fields[0].AsString);
          Inc(Cnt);
          Next;
        end;
      Close;
    end;
end;

function TFBExtractor.EnumDomains: TStringDynArray;
// ドメインの列挙
begin
  Result := BuildEnumObjects('RDB$FIELD_NAME',
                             'RDB$FIELDS',
                             '(NOT RDB$FIELD_NAME STARTING WITH ''RDB$'')',
                             '');
end;

function TFBExtractor.EnumExceptions: TStringDynArray;
// 例外の列挙
begin
  Result := BuildEnumObjects('RDB$EXCEPTION_NAME',
                             'RDB$EXCEPTIONS',
                             '',
                             '');
end;

function TFBExtractor.EnumFilters: TStringDynArray;
// BLOB フィルタの列挙
begin
  Result := BuildEnumObjects('RDB$FUNCTION_NAME',
                             'RDB$FILTERS',
                             '',
                             '');
end;

function TFBExtractor.EnumFunctions: TStringDynArray;
// UDF の列挙
begin
  Result := BuildEnumObjects('RDB$FUNCTION_NAME',
                             'RDB$FUNCTIONS',
                             '',
                             '');
end;

function TFBExtractor.EnumGenerators: TStringDynArray;
// ジェネレータの列挙
begin
  Result := BuildEnumObjects('RDB$GENERATOR_NAME',
                             'RDB$GENERATORS',
                             '',
                             '');
end;

function TFBExtractor.EnumGrants(IsView: Boolean): TStringDynArray;
// 権限の列挙
const
  OBJECT_STRING: array [Boolean] of string = ('', 'NOT ');
begin
  Result := BuildEnumObjects('RDB$RELATION_NAME',
                             'RDB$RELATIONS',
                             '(RDB$VIEW_BLR IS ' + OBJECT_STRING[IsView] + 'NULL) AND (RDB$SECURITY_CLASS STARTING WITH ''SQL$'')',
                             '');
end;

function TFBExtractor.EnumIndicies(TableName: string): TIndexDynArray;
// インデックスの列挙
var
  SelectSQL: TSelectSQL;
  Cnt, RecCnt: Integer;
begin
  with FB.Query[0] do
    begin
      // インデックス数の取得
      SelectSQL.SELECT := 'Count(*)';
      SelectSQL.FROM   := 'RDB$INDICES A ' +
                          'LEFT JOIN RDB$RELATIONS B ON ' +
                          '(B.RDB$RELATION_NAME = A.RDB$RELATION_NAME) ' +
                          'LEFT JOIN RDB$RELATION_CONSTRAINTS C ON ' +
                          '(C.RDB$INDEX_NAME = A.RDB$INDEX_NAME)';
      SelectSQL.WHERE  := '(A.RDB$RELATION_NAME = :_RELATION_NAME) AND ' +
                          '((B.RDB$SYSTEM_FLAG = 0) OR (B.RDB$SYSTEM_FLAG IS NULL)) AND ' +
                          '(C.RDB$INDEX_NAME IS NULL)';
      SQL.Text := SelectSQL.Build;
      ParamByName('_RELATION_NAME').AsString := TableName;
      Open;
      RecCnt := Fields[0].AsInteger;
      Close;
      SetLength(Result, RecCnt);
      // インデックスの列挙
      SelectSQL.SELECT := 'A.RDB$INDEX_NAME, A.RDB$UNIQUE_FLAG, A.RDB$INDEX_TYPE';
      SelectSQL.ORDER  := 'A.RDB$INDEX_NAME';
      SQL.Text := SelectSQL.Build;
      ParamByName('_RELATION_NAME').AsString := TableName;
      Cnt := 0;
      Open;
      while not EOF  do
        begin
          Result[Cnt].TableName := Trim(TableName);
          Result[Cnt].Name      := Trim(Fields[0].AsString);
          Result[Cnt].Unique    := (Fields[1].AsInteger = 1);
          Result[Cnt].IndexType := Fields[2].AsInteger;
          Inc(Cnt);
          Next;
        end;
      Close;
    end;
end;

function TFBExtractor.EnumProcedures: TStringDynArray;
// ストアドプロシージャの列挙
begin
  Result := BuildEnumObjects('RDB$PROCEDURE_NAME',
                             'RDB$PROCEDURES',
                             '',
                             '');
end;

function TFBExtractor.EnumRoles: TStringDynArray;
// ロールの列挙
begin
  Result := BuildEnumObjects('RDB$ROLE_NAME',
                             'RDB$ROLES',
                             '',
                             ''); // 古い ODS にはロールに RDB$SYSTEM_FLAG がない
end;

function TFBExtractor.EnumTables: TStringDynArray;
// テーブルの列挙
begin
  Result := BuildEnumObjects('RDB$RELATION_NAME',
                             'RDB$RELATIONS',
                             '(RDB$VIEW_BLR IS NULL)',  // 古い ODS には RDB$RELATION_TYPE が存在しない
                             '');
end;

function TFBExtractor.EnumTriggers(TableName: string): TTriggerDynArray;
// トリガの列挙
var
  SelectSQL: TSelectSQL;
  Cnt, RecCnt: Integer;
begin
  with FB.Query[0] do
    begin
      // トリガ数の取得
      SelectSQL.SELECT := 'Count(*)';
      SelectSQL.FROM   := 'RDB$TRIGGERS A ' +
                          'LEFT JOIN RDB$RELATIONS B ON ' +
                          '(B.RDB$RELATION_NAME = A.RDB$RELATION_NAME) ' +
                          'LEFT JOIN RDB$CHECK_CONSTRAINTS C ON ' +
                          '(C.RDB$TRIGGER_NAME = A.RDB$TRIGGER_NAME)';
      SelectSQL.WHERE  := '(A.RDB$RELATION_NAME = :_RELATION_NAME) AND ' +
                          '((B.RDB$SYSTEM_FLAG = 0) OR (B.RDB$SYSTEM_FLAG IS NULL)) AND ' +
                          '(C.RDB$TRIGGER_NAME IS NULL)';
      SQL.Text := SelectSQL.Build;
      ParamByName('_RELATION_NAME').AsString := TableName;
      Open;
      RecCnt := Fields[0].AsInteger;
      Close;
      SetLength(Result, RecCnt);
      // トリガの列挙
      SelectSQL.SELECT := 'A.RDB$TRIGGER_NAME, A.RDB$TRIGGER_SEQUENCE, A.RDB$TRIGGER_TYPE, A.RDB$TRIGGER_INACTIVE, A.RDB$TRIGGER_SOURCE';
      SelectSQL.ORDER  := 'A.RDB$TRIGGER_SEQUENCE, A.RDB$TRIGGER_NAME';
      SQL.Text := SelectSQL.Build;
      ParamByName('_RELATION_NAME').AsString := TableName;
      Cnt := 0;
      Open;
      while not EOF  do
        begin
          Result[Cnt].TableName   := Trim(TableName);
          Result[Cnt].Name        := Trim(Fields[0].AsString);
          Result[Cnt].Sequence    := Fields[1].AsInteger;
          Result[Cnt].TriggerType := Fields[2].AsInteger;
          Result[Cnt].Active      := (not Fields[3].IsNull) and (Fields[3].AsInteger <> 1);
          Result[Cnt].Source      := Fields[4].AsString;
          Inc(Cnt);
          Next;
        end;
      Close;
    end;
end;

function TFBExtractor.EnumViews: TStringDynArray;
// ビューの列挙
begin
  Result := BuildEnumObjects('RDB$RELATION_NAME',
                             'RDB$RELATIONS',
                             '(NOT RDB$VIEW_BLR IS NULL)', // 古い ODS には RDB$RELATION_TYPE が存在しない
                             'RDB$RELATION_ID');           // View は作成順でソートしないと破綻する
end;

function TFBExtractor.GetFieldTypeString(FieldType: Integer): string;
// フィールドタイプの検索
var
  FT: TFieldType;
begin
  Result := '';
  for FT in FieldTypes do
    begin
      if FT.FieldType =  FieldType then
        begin
          Result := FT.TypeName;
          break;
        end;
    end;
end;

function TFBExtractor.GetFieldSubTypeString(FieldSubType: Integer): string;
// フィールドサブタイプの検索
var
  FT: TFieldType;
begin
  Result := '';
  for FT in FieldSubTypes do
    begin
      if FT.FieldType =  FieldSubType then
        begin
          Result := FT.TypeName;
          break;
        end;
    end;
  if Result = '' then
    Result := IntToStr(FieldSubType);
end;

function TFBExtractor.ConvertString(Bytes: TArray<Byte>; SrcCodePage,
  DstCodePage: uInt32): string;
// 文字コード変換
var
  DstEnc, SrcEnc: TEncoding;
begin
  SrcEnc := TEncoding.GetEncoding(SrcCodePage);
  DstEnc := TEncoding.GetEncoding(DstCodePage);
  try
    Result := StringOf(TEncoding.Convert(SrcEnc, DstEnc, Bytes));
  finally
    DstEnc.Free;
    SrcEnc.Free;
  end;
end;

end.
