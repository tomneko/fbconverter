{*******************************************************}
{                                                       }
{              Firebird Database Converter              }
{                                                       }
{*******************************************************}

// -----------------------------------------------------------------------------
// Note:
//   - Firebird Utility Unit for ZeosDBO
// -----------------------------------------------------------------------------
unit uZeosFBUtils;

interface

{$DEFINE USEINDY}

uses
  System.Classes, System.SysUtils, System.Types, System.RegularExpressions,
  Windows, Data.DB, ZAbstractConnection, ZConnection, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, ZSqlMetadata, ZDbcIntfs, ZCompatibility,
  ZPlainLoader, ZPlainFirebirdDriver, ZPlainFirebirdInterbaseConstants,
  System.SyncObjs, uFBRawAccess, uSQLBuilder
  {$IFDEF USEINDY}, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdExceptionCore{$ENDIF};

type
  TFBQueryAray = array of TZQuery;
  TFBReadOnlyQueryAray = array of TZReadOnlyQuery;
  TDataSourceAray = array of TDataSource;
  TFBVersion = (fbvUnknown, fbv10, fbv15, fbv20, fbv21, fbv25, fbv30);
  TFBServerType = (fbtUnknown, fbtClientServer, fbtEmbedded);
  TFBVersions = set of TFBVersion;

  { TConnectionRec }
  TConnectionRec =
  packed record
    ServerTYpe: TFBServerType;
    Version: TFBVersion;
    LibraryLocation: string;
    HostName: string;
    DataBaseName: string;
    Port: uInt16;
    Charset: string;
    ServerCharset: string;
    ClientCharset: string;
    UserName: string;
    Password: string;
    TransactIsolationLevel: TZTransactIsolationLevel;
  end;

  { TVersion }
  TVersion =
  record
    MajorVersion: Integer;
    MinorVersion: Integer;
    ReleaseVersion: Integer;
    BuildVersion: Integer;
  end;

  { TFBDatabaseInfo }
  TFBDatabaseInfo =
  record
    ODS: TVersion;
    PageSize: uInt32;
    Dialect: Byte;
    Msg: string;
    Success: Boolean;
    function GetODSString: string;
  end;

  { TFBServerInfo }
  TFBServerInfo =
  record
    Version: TVersion;
    VersionStr: string;
    RootDir: string;
    ManegerVersion: Integer;
    Msg: string;
    Success: Boolean;
    function GetVersionString: string;
  end;

  { TFBClientInfo }
  TFBClientInfo =
  record
    Version: TVersion;
    VersionStr: string;
    function GetVersionString: string;
  end;

const
  DEFAULT_ADMIN_USERNAME = 'SYSDBA';
  DEFAULT_ADMIN_PASSWORD = 'masterkey';
  DEFAULT_HOSTNAME       = 'localhost';
  NONE_CHARSET           = 'NONE';
  DEFAULT_CHARSET        = NONE_CHARSET;
  DEFAULT_GDS_PORT       = 3050;
  DEFAULT_PAGE_SIZE      = 4096;
  DEFAULT_DIALECT        = 3;

type
  TZeosFB = class(TComponent)
  private
    { Private 宣言 }
    FClientCharset: string;
    FConnection: TZConnection;
    FDatabaseName: string;
    FDataSource: TDataSourceAray;
    FFBServerType: TFBServerType;
    FFBVersion: TFBVersion;
    FHostName: string;
    FInternalQuery: TZQuery;
    FInternalSQL: TStringList;
    FLibraryLocation: string;
    FMetaData: TZSQLMetadata;
    FPassword: string;
    FPort: UInt32;
    FQuery: TFBQueryAray;
    FQueryCount: Integer;
    FReadOnlyQuery: TFBReadOnlyQueryAray;
    FServerCharset: string;
    FTransactIsolationLevel: TZTransactIsolationLevel;
    FUserName: string;
    FWorkSQL: TStringList;
    function FBVersionToString(Version: TFBVersion): string;
    procedure FreeQueryArray;
    function GetConnectonInfo: TConnectionRec;
    procedure InitQueryArray;
    procedure SetClientCharset(const Value: string);
    procedure SetConnectonInfo(const Value: TConnectionRec);
    procedure SetDatabaseName(const Value: string);
    procedure SetFBServerType(const Value: TFBServerType);
    procedure SetFBVersion(const Value: TFBVersion);
    procedure SetHostName(const Value: string);
    procedure SetLibraryLocation(const Value: string);
    procedure SetPassword(const Value: string);
    procedure SetPort(const Value: UInt32);
    procedure SetQueryCount(const Value: Integer);
    procedure SetServerCharset(const Value: string);
    procedure SetTransactIsolationLevel(const Value: TZTransactIsolationLevel);
    procedure SetUserName(const Value: string);
  public
    { Public 宣言 }
    // Constructor / Destructor
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    // DB Method
    class function BuildFileSpec(aHostName: string; aPort: uInt16; aDatabaseName: string): string;
    procedure Commit;
    function Connect: Boolean;
    function CreateDatabase(Version: TFBVersion; LibName: string; HostName: string; Port: uInt32; DBFileName: string;
      CharSet: string = DEFAULT_CHARSET; UserName: string = DEFAULT_ADMIN_USERNAME; Password: string = DEFAULT_ADMIN_PASSWORD;
      PageSize: uInt32 = DEFAULT_PAGE_SIZE; Dialect: Byte = DEFAULT_DIALECT): string;
    procedure Disconnect;
    class function GetClientInfo(aDllName: string): TFBClientInfo;
    function GetDatabaseInfo(QueryFBServer: Boolean = True): TFBDatabaseInfo;
    function GetServerInfo: TFBServerInfo;
    procedure LoadFromBlob(Stream: TStream; Blob: TBlobField);
    procedure RollBack;
    function SimpleConnectionCheck(var Msg: string): Boolean;
    procedure StartTransaction;
    class function VersionToString(Version: TFBVersion; ExcludeDecimalPoint: Boolean = False): string;
    class function VersionToFBVersion(Version: TVersion): TFBVersion;
    // Properties
    property ClientCharset: string read FClientCharset write SetClientCharset;
    property Connection: TZConnection read FConnection write FConnection;
    property ConnectonInfo: TConnectionRec read GetConnectonInfo write SetConnectonInfo;
    property DatabaseName: string read FDatabaseName write SetDatabaseName;
    property HostName: string read FHostName write SetHostName;
    property LibraryLocation: string read FLibraryLocation write SetLibraryLocation;
    property MetaData: TZSQLMetadata read FMetaData write FMetaData;
    property Password: string read FPassword write SetPassword;
    property Port: UInt32 read FPort write SetPort default DEFAULT_GDS_PORT;
    property Query: TFBQueryAray read FQuery write FQuery;
    property QueryCount: Integer read FQueryCount write SetQueryCount;
    property ReadOnlyQuery: TFBReadOnlyQueryAray read FReadOnlyQuery write FReadOnlyQuery;
    property ServerCharset: string read FServerCharset write SetServerCharset;
    property ServerType: TFBServerType read FFBServerType write SetFBServerType;
    property TransactIsolationLevel: TZTransactIsolationLevel read FTransactIsolationLevel write SetTransactIsolationLevel;
    property UserName: string read FUserName write SetUserName;
    property Version: TFBVersion read FFBVersion write SetFBVersion;
    property WorkSQL: TStringList read FWorkSQL write FWorkSQL;
  end;

const
  VERSION_NUMBER: array [TFBVersion] of Single = (0, 1.0, 1.5, 2.0, 2.1, 2.5, 3.0);

  NULL_STRING   = 'NULL';
  TERM_CHAR     = ';';
  ALT_TERM_CHAR = '^';

  SET_TERM       = 'SET TERM ';
  SET_TERM_BEGIN = SET_TERM + TERM_CHAR + ALT_TERM_CHAR; // SET TERM ;^
  SET_TERM_END   = SET_TERM + ALT_TERM_CHAR + TERM_CHAR; // SET TERM ^;
  COMMIT_WORK    = 'COMMIT WORK' + TERM_CHAR;

  SERVICE_MANAGER_STR = 'service_mgr';

  DmyVersion: TVersion = (MajorVersion: 0; MinorVersion: 0; ReleaseVersion: 0; BuildVersion: 0);

implementation

// 内部使用メソッド
// -----------------------------------------------------------------------------

function GetVersionInfo(const aFileName: TFileName): TVSFixedFileInfo;
// バージョンリソースの取得
var
  Buf, FIBuf: Pointer;
  dwHandle, VISize: DWORD;
begin
  VISize := GetFileVersionInfoSize(PChar(aFileName), dwHandle);
  if VISize > 0 then
    begin
      GetMem(Buf, VISize);
      if not GetFileVersionInfo(PChar(aFileName), dwHandle, VISize, Buf) then
        RaiseLastOSError;
      if not VerQueryValue(Buf, '\', FIBuf, VISize) then
        RaiseLastOSError;
      Result := PVSFixedFileInfo(FIBuf)^;
      FreeMem(Buf);
    end;
end;

function ExtractVersion(s: string): TVersion;
// バージョン文字列からバージョンを取得
const
  REGEXP_VERSION_EXP = '(?<MAJOR>\d{1,})\.(?<MINOR>\d{1,})\.(?<RELEASE>\d{1,})\.(?<BUILD>\d{1,})';
var
  Match: TMatch;
begin
  Result := DmyVersion;
  Match := TRegEx.Match(s, REGEXP_VERSION_EXP, []);
  if Match.Success then
    begin
      Result.MajorVersion   := StrToIntDef(Match.Groups[1].Value, 0);
      Result.MinorVersion   := StrToIntDef(Match.Groups[2].Value, 0);
      Result.ReleaseVersion := StrToIntDef(Match.Groups[3].Value, 0);
      Result.BuildVersion   := StrToIntDef(Match.Groups[4].Value, 0);
    end;
end;

{ TZeosFB }

// コンストラクタ / デストラクタ
// -----------------------------------------------------------------------------

constructor TZeosFB.Create(AOwner: TComponent);
// コンストラクタ
begin
  inherited;
  // Params
  FClientCharset          := DEFAULT_CHARSET;
  FServerCharset          := DEFAULT_CHARSET;
  FFBVersion              := fbv21;
  FHostName               := DEFAULT_HOSTNAME;
  FPassword               := DEFAULT_ADMIN_PASSWORD;
  FPort                   := DEFAULT_GDS_PORT;
  FTransactIsolationLevel := tiReadCommitted;
  FUserName               := DEFAULT_ADMIN_USERNAME;

  // SQL Work
  FWorkSQL := TStringList.Create;
  FInternalSQL := TStringList.Create;

  // Connection
  FConnection := TZConnection.Create(Self);
  FConnection.LoginPrompt            := False;
  FConnection.HostName               := FHostName;
  FConnection.Port                   := FPort;
  FConnection.User                   := FUserName;
  FConnection.Password               := FPassword;
  FConnection.ControlsCodePage       := cCP_UTF16;
  FConnection.ClientCodepage         := FClientCharset;
  FConnection.TransactIsolationLevel := FTransactIsolationLevel;

  // Query (TIBQuery)
  FQueryCount := 5; // 初期値
  InitQueryArray;

  // MetaData
  FMetaData := TZSQLMetadata.Create(Self);
  FMetaData.Connection := FConnection;

  // Internal Query
  FInternalQuery := TZQuery.Create(Self);
  FInternalQuery.Connection := FConnection;
end;

destructor TZeosFB.Destroy;
// デストラクタ
begin
  // Internal Query
  FInternalQuery.Free;
  // MetaData
  FMetaData.Free;
  // Query
  FreeQueryArray;
  // Connection
  if Assigned(FConnection) and FConnection.Connected then
    if FConnection.InTransaction then
      FConnection.Rollback; // トランザクション中ならば強制的にロールバックする。
  FConnection.Free;
  // SQL Work
  FInternalSQL.Free;
  FWorkSQL.Free;
  inherited;
end;

// クエリコンポーネント配列
// -----------------------------------------------------------------------------

procedure TZeosFB.InitQueryArray;
// クエリコンポーネント配列を初期化
var
  i: Integer;
begin
  FreeQueryArray;
  SetLength(FQuery, FQueryCount);
  SetLength(FReadOnlyQuery, FQueryCount);
  for i:=Low(FQuery) to High(FQuery) do
    begin
      FQuery[i] := TZQuery.Create(Self);
      FQuery[i].Connection := FConnection;
      FReadOnlyQuery[i] := TZReadOnlyQuery.Create(Self);
      FReadOnlyQuery[i].Connection := FConnection;
    end;
  SetLength(FDataSource, FQueryCount);
  for i:=Low(FDataSource) to High(FDataSource) do
    begin
      FDataSource[i] := TDataSource.Create(Self);
      FDataSource[i].DataSet := FQuery[i];
    end;
end;

procedure TZeosFB.FreeQueryArray;
// クエリコンポーネント配列を解放
var
  i: Integer;
begin
  for i:=Low(FDataSource) to High(FDataSource) do
    FDataSource[i].Free;
  for i:=Low(FQuery) to High(FQuery) do
   begin
     FReadOnlyQuery[i].Free;
     FQuery[i].Free;
   end;
end;

// プロパティ: アクセスメソッド
// -----------------------------------------------------------------------------

function TZeosFB.GetConnectonInfo: TConnectionRec;
// 接続情報
begin
  with Result do
    begin
      HostName               := Self.HostName;
      Port                   := Self.Port;
      DataBaseName           := Self.DatabaseName;
      UserName               := Self.UserName;
      Password               := Self.Password;
      ServerCharset          := Self.ServerCharset;
      ClientCharset          := Self.ClientCharset;
      Version                := Self.Version;
      LibraryLocation        := Self.LibraryLocation;
      TransactIsolationLevel := Self.TransactIsolationLevel;
    end;
end;

procedure TZeosFB.SetDatabaseName(const Value: string);
// Firebird データベース名指定
begin
  FDatabaseName := Value;
  FConnection.Database := FDatabaseName;
end;

procedure TZeosFB.SetFBVersion(const Value: TFBVersion);
// Firebird バージョン指定
begin
  FFBVersion := Value;
  FConnection.Protocol := FBVersionToString(Value);
end;

procedure TZeosFB.SetPort(const Value: UInt32);
// Firebird ポート指定
begin
  if Value = 0 then
    FPort := DEFAULT_GDS_PORT
  else
    FPort := Value;
  FConnection.Port := FPort;
end;

procedure TZeosFB.SetHostName(const Value: string);
// Firebird ホスト名指定
begin
  FHostName := Value;
  FConnection.HostName := FHostName;
end;

procedure TZeosFB.SetLibraryLocation(const Value: string);
// Firebird ライブラリ位置指定
begin
  FLibraryLocation := Value;
  FConnection.LibraryLocation := FLibraryLocation;
end;

procedure TZeosFB.SetPassword(const Value: string);
// Firebird パスワード指定
begin
  FPassword := Value;
  FConnection.Password := FPassword;
end;

procedure TZeosFB.SetQueryCount(const Value: Integer);
// クエリコンポーネント配列の数を設定
begin
  FQueryCount := Value;
  InitQueryArray;
end;

procedure TZeosFB.SetServerCharset(const Value: string);
// Firebird 文字コード (サーバ)
begin
  FServerCharset := Value;
  FConnection.Properties.Values['lc_ctype'] := FServerCharset;
end;

procedure TZeosFB.SetFBServerType(const Value: TFBServerType);
// Firebird サーバ種類
begin
  FFBServerType := Value;
end;

procedure TZeosFB.SetTransactIsolationLevel(
  const Value: TZTransactIsolationLevel);
// Firebird トランザクション ISO レベル指定
begin
  FTransactIsolationLevel := Value;
  FConnection.TransactIsolationLevel := FTransactIsolationLevel;
end;

procedure TZeosFB.SetClientCharset(const Value: string);
// Firebird 文字コード (クライアント)
begin
  FClientCharset := Value;
  FConnection.ClientCodepage := FClientCharset;
end;

procedure TZeosFB.SetConnectonInfo(const Value: TConnectionRec);
// 接続情報
begin
  Self.HostName               := Value.HostName;
  Self.Port                   := Value.Port;
  Self.DataBaseName           := Value.DatabaseName;
  Self.UserName               := Value.UserName;
  Self.Password               := Value.Password;
  Self.ServerCharset          := Value.ServerCharset;
  Self.ClientCharset          := Value.ClientCharset;
  Self.Version                := Value.Version;
  Self.LibraryLocation        := Value.LibraryLocation;
  Self.TransactIsolationLevel := Value.TransactIsolationLevel;
end;

procedure TZeosFB.SetUserName(const Value: string);
// Firebird ユーザ名指定
begin
  FUserName := Value;
  FConnection.User := FUserName;
end;

// メソッド
// -----------------------------------------------------------------------------

class function TZeosFB.VersionToFBVersion(Version: TVersion): TFBVersion;
// バージョン -> Firebird バージョン
var
  v: TFBVersion;
begin
  result := fbvUnknown;
  for v := Low(TFBVersion) to High(TFBVersion) do
    begin
      if FormatFloat('#0.0', VERSION_NUMBER[v]) = Format('%d.%d', [Version.MajorVersion, Version.MinorVersion]) then
        begin
          result := v;
          Break;
        end;
    end;
end;

class function TZeosFB.VersionToString(Version: TFBVersion; ExcludeDecimalPoint: Boolean): string;
// バージョン -> バージョン文字列
begin
  Result := FormatFloat('#0.0', VERSION_NUMBER[Version]);
  if ExcludeDecimalPoint then
    Result := StringReplace(Result, '.', '', [rfReplaceAll]);
end;

function TZeosFB.FBVersionToString(Version: TFBVersion): string;
// バージョン -> バージョン文字列
begin
  Result := 'firebird-' + VersionToString(Version);
end;

procedure TZeosFB.LoadFromBlob(Stream: TStream; Blob: TBlobField);
// BLOB フィールドの読み込み
begin
  Blob.SaveToStream(Stream);
  Stream.Seek(0, soFromBeginning);
end;

function TZeosFB.SimpleConnectionCheck(var Msg: string): Boolean;
// サーバ接続確認 (TCP/IP)
{$IFDEF USEINDY}
var
  TCP: TIdTCPClient;
{$ENDIF}
begin
{$IFDEF USEINDY}
  if Self.HostName = '' then
    begin
      Result := True;
      Exit;
    end;
  Result := False;
  TCP := TIdTCPClient.Create(Self);
  try
    TCP.Host := Self.HostName;
    TCP.Port := Self.Port;
    try
      TCP.Connect;
    except
      on E: EIdSocksRequestFailed do
        Msg := E.Message;
      on E: EIdSocksRequestServerFailed do
        Msg := E.Message;
      on E: EIdSocksServerRespondError do
        Msg := E.Message;
      on E: Exception do
        Msg := E.Message;
    end;
    if TCP.Connected then
      begin
        TCP.Disconnect;
        Result := True;
      end;
  finally
    TCP.Free;
  end;
{$ELSE}
  Result := True;
{$ENDIF}
end;

function TZeosFB.GetDatabaseInfo(QueryFBServer: Boolean): TFBDatabaseInfo;
// データベース情報の取得
// -----------------------------------------------------------------------------
// http://edn.embarcadero.com/article/25494
// -----------------------------------------------------------------------------
var
  DatabaseHandle: TISC_DB_HANDLE;
  DatabaseName, DPB, db_items, mes_buffer, res_buffer: TBytes;
  FBVersion: TFBVersion;
  isc_attach_database: Tisc_attach_database;
  isc_database_info: Tisc_database_info;
  isc_detach_database: Tisc_detach_database;
  isc_interprete: Tisc_interprete;
  isc_vax_integer: Tisc_vax_integer;
  Item: Byte;
  ItemLen: Integer;
  LibraryLoader: TZNativeLibraryLoader;
  P: PByte;
  SelectSQL: TSelectSQL;
  Status: ISC_STATUS_VECTOR;
  PStatus: PSTATUS_VECTOR;
  ret: ISC_STATUS;
begin
  Result.Success   := False;
  Result.PageSize := DEFAULT_PAGE_SIZE;
  Result.Dialect  := DEFAULT_DIALECT;
  Result.ODS.MajorVersion := 10;
  Result.ODS.MinorVersion := 0;
  if QueryFBServer then
    begin
      // サーバに問い合わせてデータベース情報を取得する
      LibraryLoader := TZNativeLibraryLoader.Create([Self.LibraryLocation]);
      try
        // ライブラリのロード
        LibraryLoader.LoadNativeLibrary;
        isc_attach_database := LibraryLoader.GetAddress('isc_attach_database');
        isc_detach_database := LibraryLoader.GetAddress('isc_detach_database');
        isc_database_info   := LibraryLoader.GetAddress('isc_database_info');
        isc_interprete      := LibraryLoader.GetAddress('isc_interprete');
        isc_vax_integer     := LibraryLoader.GetAddress('isc_vax_integer');

        // データベースに接続
        DatabaseHandle := 0;
        DPB := Build_PB(Self.UserName, Self.Password);
        DatabaseName := Build_DatabaseName(Self.HostName, Self.Port, Trim(Self.DatabaseName));
        ret := isc_attach_database(@Status, Length(DatabaseName), PByte(DatabaseName), @DatabaseHandle, Length(DPB), PByte(DPB));
        // メッセージバッファの確保
        SetLength(mes_buffer, 1000);
        PStatus := @Status;
        isc_interprete(PByte(mes_buffer), @PStatus);
        Result.Msg := Format('[%s]: %s', [SERVICE_MANAGER_STR, TEncoding.ANSI.GetString(mes_buffer, 0, Length(mes_buffer))]);
        Result.Success := (ret = 0);
        if Result.Success then
          begin
            // 問い合わせ
            SetLength(db_items, 5);
            db_items[0] := isc_info_page_size;         // データベースのページサイズ
            db_items[1] := isc_info_ods_version;       // データベースの ODS メジャーバージョン
            db_items[2] := isc_info_ods_minor_version; // データベースの ODS マイナーバージョン
            db_items[3] := isc_info_db_SQL_dialect;    // データベースのダイアレクト
            db_items[4] := isc_info_end;               // 終端
            SetLength(res_buffer, 1000);
            if isc_database_info(@Status, @DatabaseHandle, Length(db_items), PByte(db_items), Length(res_buffer), PByte(res_buffer)) = 0 then
              begin
                P := PByte(res_buffer);
                while P^ <> isc_info_end do
                  begin
                    Item := P^;
                    Inc(P);
                    ItemLen := isc_vax_integer(P, 2);
                    Inc(P, 2);
                    case Item of
                      isc_info_page_size:
                        Result.PageSize         := isc_vax_integer(P, ItemLen);
                      isc_info_ods_version:
                        Result.ODS.MajorVersion := isc_vax_integer(P, ItemLen);
                      isc_info_ods_minor_version:
                        Result.ODS.MinorVersion := isc_vax_integer(P, ItemLen);
                      isc_info_db_SQL_dialect:
                        Result.Dialect          := isc_vax_integer(P, ItemLen);
                    end;
                    Inc(P, ItemLen);
                  end;
              end;
            // データベースから切断
            isc_detach_database(@Status, @DatabaseHandle);
          end;
      finally
        LibraryLoader.Free;
      end;
    end
  else
    begin
      // サーバに問い合わせずにデータベース情報を取得する
      with FInternalQuery do
        begin
          // 初期値
          FBVersion := fbv10;
          // SQL のビルド
          SelectSQL.SELECT := 'Count(*)';
          SelectSQL.FROM   := 'RDB$CHARACTER_SETS';
          SelectSQL.WHERE  := '(RDB$CHARACTER_SET_NAME = :_CHARACTER_SET_NAME)';
          SQL.Text := SelectSQL.Build;
          // FB 1.5 以上か？
          Params[0].AsString := 'ISO8859_3';
          Open;
          Result.Success := True;
          if Fields[0].AsInteger > 0 then
            begin
              FBVersion := fbv15;
              Result.ODS.MajorVersion := 10;
              Result.ODS.MinorVersion := 1;
            end;
          Close;
          // FB 2.0 以上か？
          if FBVersion = fbv15 then
            begin
              Params[0].AsString := 'UTF8';
              Open;
              if Fields[0].AsInteger > 0 then
                begin
                  FBVersion := fbv20;
                  Result.ODS.MajorVersion := 11;
                  Result.ODS.MinorVersion := 0;
                end;
              Close;
            end;
          // FB 2.1 以上か？
          if FBVersion = fbv20 then
            begin
              Params[0].AsString := 'CP943C';
              Open;
              if Fields[0].AsInteger > 0 then
                FBVersion := fbv21;
              Close;
            end;
          // FB 2.1 以上の詳細を調べる
          if FBVersion = fbv21 then
            begin
              // ODS
              SelectSQL.Init;
              SelectSQL.SELECT := '*';
              SelectSQL.FROM   := 'MON$DATABASE';
              SQL.Text := SelectSQL.Build;
              Open;
              Result.ODS.MajorVersion := FieldByName('MON$ODS_MAJOR').AsInteger;
              Result.ODS.MinorVersion := FieldByName('MON$ODS_MINOR').AsInteger;
              Result.PageSize         := FieldByName('MON$PAGE_SIZE').AsInteger;
              Result.Dialect          := FieldByName('MON$SQL_DIALECT').AsInteger;
              Close;
            end;
        end;
    end;
end;

function TZeosFB.GetServerInfo: TFBServerInfo;
// サーバ情報の取得 (要 Super Server)
// -----------------------------------------------------------------------------
// http://edn.embarcadero.com/article/27002
// https://www.ibphoenix.com/resources/documents/design/doc_178
// https://www.ibphoenix.com/resources/documents/design/doc_179
// https://www.ibphoenix.com/resources/documents/design/doc_180
// -----------------------------------------------------------------------------
var
  i: Integer;
  isc_service_attach: Tisc_service_attach;
  isc_service_detach: Tisc_service_detach;
  isc_service_query: Tisc_service_query;
  isc_interprete: Tisc_interprete;
  isc_vax_integer: Tisc_vax_integer;
  Item: Byte;
  ItemLen, SubItemLen: Integer;
  ItemStr: string;
  LibraryLoader: TZNativeLibraryLoader;
  ItemPtr, P, SP : PByte;
  ServiceHandle: PISC_SVC_HANDLE;
  db_items, res_buffer, mes_buffer, ServiceName, SPB: TBytes;
  Status: ISC_STATUS_VECTOR;
  PStatus: PSTATUS_VECTOR;
  ret: ISC_STATUS;
//SubItem: Byte;
begin
  Result.VersionStr := '';
  Result.Version := DmyVersion;
  for i:= Low(Status) to High(Status) do
    Status[i] := 0;
  ServiceHandle := nil;

  LibraryLoader := TZNativeLibraryLoader.Create([Self.FLibraryLocation]);
  try
    // ライブラリのロード
    LibraryLoader.LoadNativeLibrary;
    isc_service_attach := LibraryLoader.GetAddress('isc_service_attach');
    isc_service_detach := LibraryLoader.GetAddress('isc_service_detach');
    isc_service_query  := LibraryLoader.GetAddress('isc_service_query');
    isc_interprete     := LibraryLoader.GetAddress('isc_interprete');
    isc_vax_integer    := LibraryLoader.GetAddress('isc_vax_integer');

    // サーバに接続
    SPB := Build_PB(Self.UserName, Self.Password);
    ServiceName := Build_DatabaseName(Self.HostName, Self.Port, SERVICE_MANAGER_STR);
    ret := isc_service_attach(@Status, Length(ServiceName), PByte(ServiceName), @ServiceHandle, Length(SPB), PByte(SPB));
    // メッセージバッファの確保
    SetLength(mes_buffer, 1000);
    PStatus := @Status;
    isc_interprete(PByte(mes_buffer), @PStatus);
    Result.Msg := Format('[%s]: %s', [SERVICE_MANAGER_STR, TEncoding.ANSI.GetString(mes_buffer, 0, Length(mes_buffer))]);
    Result.Success := (ret = 0);
    if Result.Success then
      begin
        // 問い合わせ
        SetLength(db_items, 4);
        db_items[0] := isc_info_svc_server_version; // サーバーのバージョン
        db_items[1] := isc_info_svc_version;        // サーバーマネージャのバージョン
        db_items[2] := isc_info_svc_get_env;        // サーバーのルートパス
        db_items[3] := isc_info_end;                // 終端
        SetLength(res_buffer, 1000);
        if isc_service_query(@Status, @ServiceHandle, nil, 0, nil, Length(db_items), PByte(db_items), Length(res_buffer), PByte(res_buffer)) = 0 then
          begin
            P := PByte(res_buffer);
            while P^ <> isc_info_end do
              begin
                Item := P^;
                Inc(P);
                ItemLen := isc_vax_integer(P, 2);
                Inc(P, 2);
                ItemPtr := P;
                ItemStr := '';
                case Item of
                  isc_info_svc_server_version:
                    begin
                      repeat
                        SP := ItemPtr;
//                      SubItem := SP^;
                        Inc(SP);
                        SubItemLen := isc_vax_integer(SP, 1);
                        Inc(SP);
                        ItemStr := ItemStr + TEncoding.ANSI.GetString(res_buffer, SP - PByte(res_buffer), SubItemLen);
                        Inc(SP, SubItemLen);
                      until (SP >= ItemPtr + ItemLen);
                      Result.VersionStr := ItemStr;
                      Result.Version := ExtractVersion(Result.VersionStr);
                    end;
                  isc_info_svc_get_env:
                    Result.RootDir := TEncoding.ANSI.GetString(res_buffer, P - PByte(res_buffer), ItemLen);
                  isc_info_svc_version:
                    Result.ManegerVersion := isc_vax_integer(P, SizeOf(UInt32));
                end;
                Inc(P, ItemLen);
              end;
          end;
        // サーバから切断
        isc_service_detach(@Status, @ServiceHandle);
      end;
  finally
    LibraryLoader.Free;
  end;
end;

class function TZeosFB.GetClientInfo(aDllName: string): TFBClientInfo;
// クライアント情報の取得
var
  Buf: TBytes;
  FileInfo: TVSFixedFileInfo;
  isc_get_client_version: Tisc_get_client_version;
  LibraryLoader: TZNativeLibraryLoader;
  LoadFlg: Boolean;
begin
  Result.VersionStr := '';
  Result.Version := DmyVersion;
  LibraryLoader := TZNativeLibraryLoader.Create([aDllName]);
  try
    try
      LibraryLoader.LoadNativeLibrary;
      LoadFlg := True;
    except
      LoadFlg := False;
    end;
    if not LoadFlg then
      Exit;
    isc_get_client_version := LibraryLoader.GetAddress('isc_get_client_version');
    if @isc_get_client_version <> nil then
      begin
        SetLength(Buf, 1000);
        isc_get_client_version(PByte(Buf));
        result.VersionStr := Trim(TEncoding.ANSI.GetString(Buf));
        result.Version := ExtractVersion(result.VersionStr);
      end
    else
      begin
        // Firebird 1.0.x
        FileInfo := GetVersionInfo(aDllName);
        result.VersionStr := Format('WI-V%d.%d.%d.%d Firebird %d.%d',
                                        [HiWord(FileInfo.dwFileVersionMS), LoWord(FileInfo.dwFileVersionMS),
                                         HiWord(FileInfo.dwFileVersionLS), LoWord(FileInfo.dwFileVersionLS),
                                         HiWord(FileInfo.dwProductVersionMS), LoWord(FileInfo.dwProductVersionMS)]);
        result.Version := ExtractVersion(result.VersionStr);
      end;
  finally
    LibraryLoader.Free;
  end;
end;

// データベースメソッド
// -----------------------------------------------------------------------------

function TZeosFB.Connect: Boolean;
// DB に接続
begin
  if FConnection.Connected then
    FConnection.Connected := False;
  FConnection.Connected := True;
  Result := FConnection.Connected;
end;

procedure TZeosFB.Disconnect;
// DB から切断
begin
  if Assigned(FConnection) and FConnection.Connected then
    FConnection.Connected := False;
end;

procedure TZeosFB.StartTransaction;
// トランザクションの開始
begin
  if not FConnection.InTransaction then
    FConnection.StartTransaction;
end;

class function TZeosFB.BuildFileSpec(aHostName: string; aPort: uInt16;
  aDatabaseName: string): string;
// Filespec (DB 接続文字列) のビルド
begin
  result := Build_FileSpec(aHostName, aPort, aDatabaseName);
end;

procedure TZeosFB.Commit;
// トランザクションのコミット
begin
  if FConnection.InTransaction then
    FConnection.Commit;
end;

procedure TZeosFB.RollBack;
// トランザクションのロールバック
begin
  if FConnection.InTransaction then
    FConnection.Rollback;
end;

function TZeosFB.CreateDatabase(Version: TFBVersion; LibName, HostName: string; Port: uInt32; DBFileName,
  CharSet, UserName, Password: string; PageSize: uInt32; Dialect: Byte): string;
// 新規データベースの作成
const
  CREATE_DB = 'CREATE DATABASE ''%S'' USER ''%S'' PASSWORD ''%S'' PAGE_SIZE %d DEFAULT CHARACTER SET %S';
var
  dConnection: TZConnection;
  DSQL: string;
begin
  dConnection := TZConnection.Create(Self);
  try
    DSQL := Format(CREATE_DB, [DBFileName, UserName, Password, PageSize, CharSet]);
    dConnection.Protocol := FBVersionToString(Version);
    dConnection.HostName := HostName;
    dConnection.Port := Port;
    dConnection.Database := DBFileName;
    dConnection.LibraryLocation := LibName;
    dConnection.Properties.Clear;
    dConnection.Properties.Values['dialect'] := IntToStr(Dialect);
    dConnection.Properties.Values['CreateNewDatabase'] := DSQL;
    dConnection.User     := UserName;
    dConnection.Password := Password;
    dConnection.Connect;
    dConnection.Disconnect;
  finally
    dConnection.Free;
  end;
  Result := DSQL;
end;

// メソッド (高度なレコード型)
// -----------------------------------------------------------------------------

{ TFBDatabaseInfo }

function TFBDatabaseInfo.GetODSString: string;
begin
  Result := Format('%d.%d', [Self.ODS.MajorVersion, Self.ODS.MinorVersion]);
end;

{ TFBClientInfo }

function TFBClientInfo.GetVersionString: string;
begin
  Result := Format('%d.%d.%d.%d',
    [Self.Version.MajorVersion, Self.Version.MinorVersion, Self.Version.ReleaseVersion, Self.Version.BuildVersion]);
end;

{ TFBServerInfo }

function TFBServerInfo.GetVersionString: string;
begin
  Result := Format('%d.%d.%d.%d',
    [Self.Version.MajorVersion, Self.Version.MinorVersion, Self.Version.ReleaseVersion, Self.Version.BuildVersion]);
end;
end.
