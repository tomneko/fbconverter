{*******************************************************}
{                                                       }
{              Firebird Database Converter              }
{                                                       }
{*******************************************************}

// -----------------------------------------------------------------------------
// Note:
//   - Main Form
// -----------------------------------------------------------------------------
unit frmuMain;

interface

uses
  System.Types, System.Classes, System.SysUtils, System.Variants, System.Actions,
  System.IOUtils, System.UITypes, System.Math, System.StrUtils, System.RegularExpressions,
  System.IniFiles, Vcl.Dialogs, Vcl.ActnList, Vcl.ExtCtrls, Vcl.StdCtrls, Vcl.Buttons,
  System.Generics.Collections, Vcl.Forms, Vcl.Graphics, Vcl.Controls,
  Winapi.Windows, Winapi.Messages, Winapi.ShellAPI, Data.DB, ZClasses, ZConnection,
  ZDbcIntfs, ZSysUtils, ZDataset, uZeosFBUtils, uFBExtract, uFBConvConsts, uIoUtilsEx,
  uSQLBuilder, AsyncCalls, ZAbstractConnection, ZCompatibility;

type
  TDropFilesEvent = procedure (Sender: TObject; FileNames: TStringDynArray; Count: Integer) of object;
  TConnectionTestType = (cttSource, cttDestination);

  { TFBBulkRec }

  TFBBulkRec =
    packed record
      TableName: string;
      SplitDataRec: TSplitDataRec;
      SrcConRec: TConnectionRec;
      DstConRec: TConnectionRec;
      StartTime: TDateTime;
      IsMultiThread: Boolean;
    end;
  PFBBulkRec = ^TFBBulkRec;

  { TFBVersionType }
  TFBVersionType =
    record
      ServerType: TFBServerType;
      Version: TFBVersion;
    end;

  { TMessageRec }
  TMessageRec =
    packed record
      StartBlock: Integer;
      EndBlock: Integer;
      StartTime: TDateTime;
    end;
  PMessageRec = ^TMessageRec;

  { TfrmMain }
  TfrmMain = class(TForm)
    AL_AutoDetect: TAction;
    AL_Cancel: TAction;
    AL_ConnectionTest_Dst: TAction;
    AL_ConnectionTest_Src: TAction;
    AL_Convert: TAction;
    AL_Information: TAction;
    AL_OpenDstFile: TAction;
    AL_OpenSrcFile: TAction;
    alMain: TActionList;
    bbAutoDetect: TBitBtn;
    bbCancel: TBitBtn;
    bbConvert: TBitBtn;
    bbDstConnectionTest: TBitBtn;
    bbDstDatabase: TBitBtn;
    bbInformation: TBitBtn;
    bbSrcConnectionTest: TBitBtn;
    bbSrcDatabase: TBitBtn;
    cbDstCharset: TComboBox;
    cbDstVersion: TComboBox;
    cbSrcCharset: TComboBox;
    cbSrcVersion: TComboBox;
    chbEmptyDatabase: TCheckBox;
    chbIntToBigInt: TCheckBox;
    chbMetadataOnly: TCheckBox;
    chbStrictCheck: TCheckBox;
    chbVerboseMode: TCheckBox;
    edDstDatabase: TEdit;
    edDstHostName: TEdit;
    edDstPassword: TEdit;
    edDstPort: TEdit;
    edDstUser: TEdit;
    edSrcDatabase: TEdit;
    edSrcHostName: TEdit;
    edSrcPassword: TEdit;
    edSrcPort: TEdit;
    edSrcUser: TEdit;
    gbDestination: TGroupBox;
    gbSource: TGroupBox;
    lblDstCharset: TLabel;
    lblDstDatabase: TLabel;
    lblDstHostName: TLabel;
    lblDstPassword: TLabel;
    lblDstPort: TLabel;
    lblDstUser: TLabel;
    lblDstVersion: TLabel;
    lblSrcCharset: TLabel;
    lblSrcDatabase: TLabel;
    lblSrcHostName: TLabel;
    lblSrcPassword: TLabel;
    lblSrcPort: TLabel;
    lblSrcUser: TLabel;
    lblSrcVersion: TLabel;
    odFile: TOpenDialog;
    pnlBottom: TPanel;
    sdFile: TSaveDialog;
    // Events
    procedure AL_AutoDetectExecute(Sender: TObject);
    procedure AL_CancelExecute(Sender: TObject);
    procedure AL_ConnectionTest_DstExecute(Sender: TObject);
    procedure AL_ConnectionTest_SrcExecute(Sender: TObject);
    procedure AL_ConvertExecute(Sender: TObject);
    procedure AL_InformationExecute(Sender: TObject);
    procedure AL_OpenDstFileExecute(Sender: TObject);
    procedure AL_OpenSrcFileExecute(Sender: TObject);
    procedure cbDstVersionChange(Sender: TObject);
    procedure DropFiles(Sender: TObject; FileNames: TStringDynArray; Count: Integer);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private 宣言 }
    // Fields
    FBlobDir: string;
    FClientDir: string;
    FClientLatestVersion: TFBVersion;
    FClientVersions: TFBVersions;
    FConvertFromIntToBigInt: Boolean;
    FDatabaseInfo: TFBDatabaseInfo;
    FDefaultCharset: string;
    FDstConnectionRec: TConnectionRec;
    FEmbeddedDir: string;
    FEmbeddedLatestVersion: TFBVersion;
    FEmbeddedVersions: TFBVersions;
    FIsConvertRunning: Boolean;
    FIsEmptyDatabase: Boolean;
    FMaxErrorLine: uint16;
    FMetadataOnly: Boolean;
    FNumberOfImportThread: Byte;
    FPostDataSplitSize: uInt16;
    FSrcConnectionRec: TConnectionRec;
    FStrictCheck: Boolean;
    FUseMultiThread: Boolean;
    FVerboseMode: Boolean;
    FWorkDir: string;
    FOnDropFiles: TDropFilesEvent;
    // Access Methods
    function GetDestinationCharset: string;
    function GetDestinationDataBaseName: string;
    function GetDestinationServerType: TFBServerType;
    function GetDestinationVersion: TFBVersion;
    function GetEmbeddedDllName: string;
    function GetEmbeddedVersion: TFBVersion;
    function GetSourceHostName: string;
    function GetIsEmptyDatabase: Boolean;
    function GetMetadataOnly: Boolean;
    function GetSourcePassword: string;
    function GetSourcePort: uInt16;
    function GetSourceCharset: string;
    function GetSourceDataBaseName: string;
    function GetStrictCheck: Boolean;
    function GetSourceUserName: string;
    procedure SetDestinationCharset(const Value: string);
    procedure SetDestinationDataBaseName(const Value: string);
    procedure SetSourceHostName(const Value: string);
    procedure SetIsEmptyDatabase(const Value: Boolean);
    procedure SetMetadataOnly(const Value: Boolean);
    procedure SetSourcePassword(const Value: string);
    procedure SetSourcePort(const Value: uInt16);
    procedure SetSourceCharset(const Value: string);
    procedure SetSourceDataBaseName(const Value: string);
    procedure SetStrictCheck(const Value: Boolean);
    procedure SetSourceUserName(const Value: string);
    // Functions
    procedure AutoDetectCharset(ShowErrorMsg: Boolean);
    function CheckError: Boolean;
    procedure CleanupWorkDir;
    procedure ConnectionTest(aType: TConnectionTestType);
    procedure ExitProgram;
    function GetClientDllName: string;
    function GetConvertFromIntToBigInt: Boolean;
    function GetDestinationHostName: string;
    function GetDestinationPassword: string;
    function GetDestinationPort: uInt16;
    function GetDestinationUserName: string;
    function GetSourceClientDllName: string;
    function GetSourceVersion: TFBVersion;
    function GetVerboseMode: Boolean;
    procedure OpenDestinationFile;
    procedure OpenSourceFile;
    procedure ProcessConvert;
    procedure SetConvertFromIntToBigInt(const Value: Boolean);
    procedure SetDestinationCharSetList;
    procedure SetDestinationHostName(const Value: string);
    procedure SetDestinationPassword(const Value: string);
    procedure SetDestinationPort(const Value: uInt16);
    procedure SetDestinationUserName(const Value: string);
    procedure SetDestinationVersion(const Value: TFBVersionType);
    procedure SetFirebirdProperty(aFirebird: TZeosFB; IsSetCharset: Boolean = False; aDisConnect: Boolean = True);
    procedure SetSourceVersion(const Value: TFBVersionType);
    procedure SetVerboseMode(const Value: Boolean);
    procedure SetWorkDir(aDir: string);
    procedure ShowInformation;
    function ValueToVersionType(aValue: uInt32): TFBVersionType;
    function VersionTypeToValue(aVersuionType: TFBVersionType): uInt32;
    // AppMessage Event
    procedure AppMessage(var Msg: TMsg; var Handled: Boolean);
  protected
    { Protected 宣言 }
    procedure UpdateActions; override;
  public
    { Public 宣言 }
    FBExtract: TFBExtractor;
    // Properties
    property BlobDir: string read FBlobDir;
    property ClientDir: string read FClientDir;
    property ClientDllName: string read GetClientDllName;
    property ClientVersions: TFBVersions read FClientVersions;
    property DatabaseInfo: TFBDatabaseInfo read FDatabaseInfo;
    property DefaultCharset: string read FDefaultCharset;
    property DestinationCharset: string read GetDestinationCharset write SetDestinationCharset;
    property DestinationDataBaseName: string read GetDestinationDataBaseName write SetDestinationDataBaseName;
    property DestinationHostName: string read GetDestinationHostName write SetDestinationHostName;
    property DestinationPassword: string read GetDestinationPassword write SetDestinationPassword;
    property DestinationPort: uInt16 read GetDestinationPort write SetDestinationPort;
    property DestinationServerType: TFBServerType read GetDestinationServerType;
    property DestinationUserName: string read GetDestinationUserName write SetDestinationUserName;
    property DestinationVersion: TFBVersion read GetDestinationVersion;
    property EmbeddedDir: string read FEmbeddedDir;
    property EmbeddedDllName: string read GetEmbeddedDllName;
    property EmbeddedVersions: TFBVersions read FEmbeddedVersions;
    property ConvertFromIntToBigInt: Boolean read GetConvertFromIntToBigInt write SetConvertFromIntToBigInt;
    property IsConvertRunning: Boolean read FIsConvertRunning;
    property IsEmptyDatabase: Boolean read GetIsEmptyDatabase write SetIsEmptyDatabase;
    property MaxErrorLine: uint16 read FMaxErrorLine;
    property MetadataOnly: Boolean read GetMetadataOnly write SetMetadataOnly;
    property NumberOfImportThread: Byte read FNumberOfImportThread;
    property PostDataSplitSize: uint16 read FPostDataSplitSize;
    property SourceCharset: string read GetSourceCharset write SetSourceCharset;
    property SourceClientDllName: string read GetSourceClientDllName;
    property SourceDataBaseName: string read GetSourceDataBaseName write SetSourceDataBaseName;
    property SourceHostName: string read GetSourceHostName write SetSourceHostName;
    property SourcePassword: string read GetSourcePassword write SetSourcePassword;
    property SourcePort: uInt16 read GetSourcePort write SetSourcePort;
    property SourceUserName: string read GetSourceUserName write SetSourceUserName;
    property SourceVersion: TFBVersion read GetSourceVersion;
    property StrictCheck: Boolean read GetStrictCheck write SetStrictCheck;
    property UseMultiThread: Boolean read FUseMultiThread;
    property VerboseMode: Boolean read GetVerboseMode write SetVerboseMode;
    property WorkDir: string read FWorkDir;
    // Events
    property OnDropFiles: TDropFilesEvent read FOnDropFiles write FOnDropFiles;
  end;

const
  DmyByte: Byte = 0;

  MAX_ERROR_LINE = 5;          // エラーの最大行数 [初期値]
  POST_DATA_SPLIT_SIZE = 1000; // インポートデータの処理サイズ (n 件単位で処理) [初期値]
  NUMBER_OF_IMPORT_THREAD = 4; // インポート用スレッドの数

  EMBEDDED_LIBNAME = 'fbembed.dll';
  CLIENT_LIBNAME = 'fbclient%s.dll';
  FB_FILE          = 'firebird';
  FB_CONFIG_FILE   = FB_FILE + '.conf';
  FB_METADATA      = 'metadata';
  FB_METADATA_FILE = FB_METADATA + '.ddl';
  FB_METADATA_SQL  = FB_METADATA + '.sql';
  FB_ERROR_FILE    = FB_METADATA + '.err';

  INI_SECTION_GLOBALS = 'Globals';

  procedure ImportData(aBulkRec: PFBBulkRec);

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}

uses
  frmuLogViewer;

procedure TfrmMain.FormCreate(Sender: TObject);
// フォーム作成時
type
  TInternalCharSet =
  record
    CodePage: uInt16;
    DispName: string;
  end;
const
  EMBEDDED_DIR = 'Embedded';
  CLIENT_DIR = 'Client';
  FB_VERSION_FORMAT = 'Firebird %s%s';
  FB_VERSION_TYPE = ' (%s)';
var
  FBVer: TFBVersion;
  FBVerTYpe: TFBVersionType;
  i: Integer;
  Ini: TMemIniFile;
  IniFileName, dDefaultCharset, dWorkDir: string;
begin
  // ---------------------------------------------------------------------------
  // 2014/12/30 時点:
  // ・ZeosLib 7.1.4 stable は Firebird 1.5 以前で問題が発生する
  // ・ZeosLib 7.1.3a stable は Firebird 1.5 以前で問題が発生する
  // ・ZeosLib 7.1.3 stable は Firebird 1.5 以前で問題が発生する
  // ---------------------------------------------------------------------------
  {$IF ZEOS_VERSION <> '7.1.2-stable'}
    {$MESSAGE ERROR 'Firebird Database Converter is may not work correctly.'}
  {$IFEND}
  // ---------------------------------------------------------------------------

  ClientWidth  := 622;
  ClientHeight := 612;

  FIsConvertRunning := False;

  // コピー元の文字コード
//dDefaultCharset := DEFAULT_CHARSET;  // NONE
  dDefaultCharset := CharSets[5].Name; // SJIS_0208
  for i:=Low(CharSets) to High(CharSets) do
    cbSrcCharset.Items.Add(CharSets[i].Name);
  cbSrcCharset.Sorted := True;

  // Embedded の存在チェック
  FEmbeddedDir := TPath.Combine(TPath.GetDirectoryName(ParamStr(0)), EMBEDDED_DIR);
  FEmbeddedVersions := [];
  for FBVer := Low(TFBVersion) to High(TFBVersion) do
    if TFile.Exists(TPath.Combine(EmbeddedDir, TZeosFB.VersionToString(FBVer)) + PathDelim + EMBEDDED_LIBNAME) then
      begin
        Include(FEmbeddedVersions, FBVer);
        FBVerType.ServerType := fbtEmbedded;
        FBVerType.Version    := FBVer;
        cbDstVersion.Items.AddObject(Format(FB_VERSION_FORMAT, [TZeosFB.VersionToString(FBVer), Format(FB_VERSION_TYPE, [EMBEDDED_DIR])]), TObject(VersionTypeToValue(FBVerType)));
        FEmbeddedLatestVersion := FBVer;
      end;

  // クライアントの存在チェック
  FClientDir := TPath.Combine(TPath.GetDirectoryName(ParamStr(0)), CLIENT_DIR);
  FClientVersions := [];
  for FBVer := Low(TFBVersion) to High(TFBVersion) do
    if TFile.Exists(TPath.Combine(ClientDir, Format(CLIENT_LIBNAME, [TZeosFB.VersionToString(FBVer, True)]))) then
      begin
        Include(FClientVersions, FBVer);
        FBVerType.ServerType := fbtClientServer;
        FBVerType.Version    := FBVer;
        cbSrcVersion.Items.AddObject(Format(FB_VERSION_FORMAT, [TZeosFB.VersionToString(FBVer), '']), TObject(VersionTypeToValue(FBVerType)));
        cbDstVersion.Items.AddObject(Format(FB_VERSION_FORMAT, [TZeosFB.VersionToString(FBVer), Format(FB_VERSION_TYPE, [CLIENT_DIR])]), TObject(VersionTypeToValue(FBVerType)));
        FClientLatestVersion := FBVer;
      end;

  // TFBExtractor の準備
  FBExtract := TFBExtractor.Create(Self);

  // 設定の読み込み
  IniFileName := ChangeFileExt(ParamStr(0), '.ini');
  Ini := TMemIniFile.Create(IniFileName, TEncoding.UTF8);
  try
    // コピー元文字コードのデフォルト
    FDefaultCharset        := Ini.ReadString (INI_SECTION_GLOBALS, 'DEFAULT_CHARSET'         , dDefaultCharset        );

    // エラー行の最大値
    FMaxErrorLine          := Ini.ReadInteger(INI_SECTION_GLOBALS, 'MAX_ERROR_LINE'          , MAX_ERROR_LINE         );

    // インポートデータの分割サイズ (レコード数)
    FPostDataSplitSize     := Ini.ReadInteger(INI_SECTION_GLOBALS, 'POST_DATA_SPLIT_SIZE'    , POST_DATA_SPLIT_SIZE   );
    FPostDataSplitSize := Max(FPostDataSplitSize, 100);

    // (可能な場合) インポート時にマルチスレッドを使うか？
    FUseMultiThread        := Ini.ReadBool   (INI_SECTION_GLOBALS, 'USE_MULTI_THREAD'        , True                   );

    // インポート時のマルチスレッド数
    FNumberOfImportThread  := Ini.ReadInteger(INI_SECTION_GLOBALS, 'NUMBER_OF_IMPORT_THREAD' , NUMBER_OF_IMPORT_THREAD);
    FNumberOfImportThread := Max(FNumberOfImportThread, 2);
    FNumberOfImportThread := Min(FNumberOfImportThread, FPostDataSplitSize div 2);

    // ワークフォルダ (バーボーズモードで使用)
    dWorkDir               := Ini.ReadString (INI_SECTION_GLOBALS, 'WORK_DIR'                , GetHomePath             );
  finally
    Ini.Free;
  end;

  // ワークフォルダの設定
  SetWorkDir(dWorkDir);

  // Drag&Drop の許可
  DragAcceptFiles(edSrcDatabase.Handle, True);
  DragAcceptFiles(edDstDatabase.Handle, True);
  Application.OnMessage:= AppMessage;
  Self.OnDropFiles := DropFiles;
end;

procedure TfrmMain.FormShow(Sender: TObject);
// フォーム表示時
var
  VersionType: TFBVersionType;
begin
  OnShow := nil;

  // Source Version
  if cbSrcVersion.Items.Count <> 0 then
    begin
      VersionType.ServerType := fbtClientServer;
      VersionType.Version    := FClientLatestVersion;
      SetSourceVersion(VersionType);
    end;

  // Source HostName
  SourceHostName := DEFAULT_HOSTNAME;

  // Source Port
  SourcePort := DEFAULT_GDS_PORT;

  // UserName / Password
  SourceUserName := DEFAULT_ADMIN_USERNAME;
  SourcePassword := DEFAULT_ADMIN_PASSWORD;

  // Source Charset
  SourceCharset := DefaultCharset;

  // Destination Version
  if cbDstVersion.Items.Count <> 0 then
    begin
      if FEmbeddedVersions <> [] then
        begin
          VersionType.ServerType := fbtEmbedded;
          VersionType.Version    := FEmbeddedLatestVersion;
        end
      else
        begin
          VersionType.ServerType := fbtClientServer;
          VersionType.Version    := FClientLatestVersion;
        end;
      SetDestinationVersion(VersionType);
    end;

  // Destination HostName
  DestinationHostName := DEFAULT_HOSTNAME;

  // Destination Port
  DestinationPort := DEFAULT_GDS_PORT;

  // UserName / Password
  DestinationUserName := DEFAULT_ADMIN_USERNAME;
  DestinationPassword := DEFAULT_ADMIN_PASSWORD;

  // Destination Charset
  cbDstVersionChange(nil);

  // Open / Save Dialog
  odFile.Filter := 'Firebird Database File|*.fdb;*.gdb;*.idb|All Files (*.*)|*.*';
  odFile.FilterIndex := 0;
  odFile.DefaultExt := 'fdb';
  odFile.Options := [ofHideReadOnly, ofPathMustExist, ofFileMustExist, ofEnableSizing];

  sdFile.Filter := odFile.Filter;
  sdFile.FilterIndex := 0;
  sdFile.DefaultExt := odFile.DefaultExt;
  sdFile.Options := [ofOverwritePrompt, ofHideReadOnly, ofPathMustExist, ofEnableSizing];

  // Position
  bbAutoDetect.Top := cbSrcCharset.Top;
  bbAutoDetect.Height := cbSrcCharset.Height;
  bbAutoDetect.Left := cbSrcCharset.Left + cbSrcCharset.Width + 4;

  bbSrcDatabase.Top := edSrcDatabase.Top;
  bbSrcDatabase.Height := edSrcDatabase.Height;
  bbSrcDatabase.Left := edSrcDatabase.Left + edSrcDatabase.Width + 4;

  bbDstDatabase.Top := edDstDatabase.Top;
  bbDstDatabase.Height := edDstDatabase.Height;
  bbDstDatabase.Left := edDstDatabase.Left + edDstDatabase.Width + 4;
end;

procedure TfrmMain.FormDestroy(Sender: TObject);
// フォーム破棄時
begin
//CleanupWorkDir;
  Application.OnMessage:= nil;
  DragAcceptFiles(edSrcDatabase.Handle, False);
  DragAcceptFiles(edDstDatabase.Handle, False);
  FBExtract.Free;
end;

procedure TfrmMain.UpdateActions;
// アイドル時
var
  i: Integer;
  OldState, NewState: Boolean;
begin
  inherited;
  if IsConvertRunning then
    Exit;

  OldState := AL_Information.Enabled;
  NewState := (SourceDataBaseName <> '') and
              (SourceUserName     <> '') and
              (SourcePassword     <> '');
  if OldState <> NewState then
    AL_Information.Enabled := NewState;

  OldState := AL_Convert.Enabled;
  NewState := AL_Information.Enabled and
              ((DestinationDataBaseName <> '') or MetadataOnly);
  if OldState <> NewState then
    AL_Convert.Enabled := NewState;

  OldState := lblDstHostName.Enabled;
  NewState := (DestinationServerType = fbtClientServer);
  if OldState <> NewState then
    begin
      lblDstHostName.Enabled        := NewState;
      edDstHostName.Enabled         := NewState;
      lblDstPort.Enabled            := NewState;
      edDstPort.Enabled             := NewState;
      lblDstUser.Enabled            := NewState;
      edDstUser.Enabled             := NewState;
      lblDstPassword.Enabled        := NewState;
      edDstPassword.Enabled         := NewState;
      AL_ConnectionTest_Dst.Enabled := NewState;
    end;

  OldState := AL_AutoDetect.Enabled;
  NewState := (SourceDataBaseName <> '') and
              (SourceUserName     <> '') and
              (SourcePassword     <> '');
  if OldState <> NewState then
    AL_AutoDetect.Enabled := NewState;

  OldState := chbStrictCheck.Enabled;
  NewState := (not Self.MetadataOnly) and (DestinationVersion >= fbv20);
  if OldState <> NewState then
    chbStrictCheck.Enabled := NewState;

  OldState := chbVerboseMode.Enabled;
  NewState := not IsEmptyDatabase;
  if OldState <> NewState then
    chbVerboseMode.Enabled := NewState;

  OldState := chbIntToBigInt.Enabled;
  NewState := (DestinationVersion >= fbv15);
  if OldState <> NewState then
    chbIntToBigInt.Enabled := NewState;

  for i:=0 to alMain.ActionCount-1 do
    begin
      OldState := alMain.Actions[0].Enabled;
      NewState := (not IsConvertRunning);
      if (alMain.Actions[i] = AL_Information) then
        NewState := NewState and AL_Information.Enabled
      else if (alMain.Actions[i] = AL_Convert) then
        NewState := NewState and AL_Convert.Enabled
      else if (alMain.Actions[i] = AL_AutoDetect) then
        NewState := NewState and AL_AutoDetect.Enabled
      else if (alMain.Actions[i] = AL_ConnectionTest_Dst) then
        NewState := NewState and AL_ConnectionTest_Dst.Enabled;
      if OldState <> NewState then
        alMain.Actions[i].Enabled := NewState;
    end;
end;

procedure TfrmMain.AppMessage(var Msg: TMsg; var Handled: Boolean);
// アプリケーション: メッセージ処理
var
  Buf: string;
  i, BufLen, FileCnt: Integer;
  Control: TControl;
  FileNames: TStringDynArray;
begin
  case Msg.Message of
    WM_DROPFILES:
      begin
        try
          if not Assigned(FOnDropFiles) then
            Exit;
          FileCnt := DragQueryFile(Msg.wParam, DWORD(-1), nil, 0);
          SetLength(FileNames, FileCnt);
          for i:=0 to FileCnt-1 do
            begin
              BufLen := DragQueryFile(Msg.wParam, i, nil, 0);
              SetLength(Buf, BufLen + 1);
              DragQueryFile(Msg.wParam, i, @Buf[1], BufLen + 1);
              FileNames[i] := Trim(Buf);
            end;
          Control := FindDragTarget(Msg.pt, False);
          OnDropFiles(Control, FileNames, FileCnt);
        finally
          DragFinish(Msg.wParam);
        end;
        Handled:= True;
      end;
  end;
end;

procedure TfrmMain.DropFiles(Sender: TObject; FileNames: TStringDynArray;
  Count: Integer);
// ファイルドロップ時
begin
  if (Sender is TCustomEdit) then
    (Sender as TCustomEdit).Text := FileNames[0];
end;

procedure TfrmMain.cbDstVersionChange(Sender: TObject);
// cbDstVersion 変更時
begin
  SetDestinationCharSetList;
  if cbDstCharset.Items.Count = 0 then
    DestinationCharset := NONE_CHARSET
  else
    begin
      // Unicode 系の文字コードをデフォルトで設定
      if (DestinationVersion >= fbv20) then
        DestinationCharset := CharSets[4].Name  // UTF8
      else
        DestinationCharset := CharSets[3].Name; // UNICODE_FSS
    end;
  chbStrictCheck.Checked := (DestinationVersion >= fbv20);
end;

// Access Methods
// -----------------------------------------------------------------------------
function TfrmMain.GetSourceClientDllName: string;
begin
  Result := TPath.Combine(ClientDir, Format(CLIENT_LIBNAME, [TZeosFB.VersionToString(Self.SourceVersion, True)]));
end;

function TfrmMain.GetClientDllName: string;
begin
  Result := TPath.Combine(ClientDir, Format(CLIENT_LIBNAME, [TZeosFB.VersionToString(Self.DestinationVersion, True)]));
end;

function TfrmMain.GetConvertFromIntToBigInt: Boolean;
begin
  FConvertFromIntToBigInt := chbIntToBigInt.Enabled and chbIntToBigInt.Checked;
  Result := FConvertFromIntToBigInt;
end;

function TfrmMain.GetDestinationCharset: string;
begin
  if cbDstCharset.ItemIndex < 0 then
    FDstConnectionRec.Charset := NONE_CHARSET
  else
    FDstConnectionRec.Charset := cbDstCharset.Items[cbDstCharset.ItemIndex];
  Result := FDstConnectionRec.Charset;
end;

function TfrmMain.GetDestinationDataBaseName: string;
begin
  FDstConnectionRec.DataBaseName := Trim(edDstDatabase.Text);
  Result := FDstConnectionRec.DataBaseName;
end;

function TfrmMain.GetDestinationHostName: string;
begin
  FDstConnectionRec.HostName := Trim(edDstHostName.Text);
  Result := FDstConnectionRec.HostName;
end;

function TfrmMain.GetDestinationPassword: string;
begin
  FDstConnectionRec.Password := Trim(edDstPassword.Text);
  Result := FDstConnectionRec.Password;
end;

function TfrmMain.GetDestinationPort: uInt16;
begin
  FDstConnectionRec.Port := StrToIntDef(edDstPort.Text, DEFAULT_GDS_PORT);
  Result := FDstConnectionRec.Port;
end;

function TfrmMain.GetDestinationServerType: TFBServerType;
begin
  if cbDstVersion.ItemIndex < 0 then
    Result := fbtUnknown
  else
    Result := ValueToVersionType(UInt32(cbDstVersion.Items.Objects[cbDstVersion.ItemIndex])).ServerType;
end;

function TfrmMain.GetDestinationUserName: string;
begin
  FDstConnectionRec.UserName := Trim(edDstUser.Text);
  Result := FDstConnectionRec.UserName;
end;

function TfrmMain.GetSourceVersion: TFBVersion;
begin
  if cbSrcVersion.ItemIndex < 0 then
    Result := fbvUnknown
  else
    Result := ValueToVersionType(UInt32(cbSrcVersion.Items.Objects[cbSrcVersion.ItemIndex])).Version;
  FSrcConnectionRec.Version := Result;
end;

function TfrmMain.GetDestinationVersion: TFBVersion;
begin
  if cbDstVersion.ItemIndex < 0 then
    Result := fbvUnknown
  else
    Result := ValueToVersionType(UInt32(cbDstVersion.Items.Objects[cbDstVersion.ItemIndex])).Version;
  FDstConnectionRec.Version := Result;
end;

function TfrmMain.GetEmbeddedDllName: string;
begin
  Result := TPath.Combine(EmbeddedDir, TZeosFB.VersionToString(Self.DestinationVersion)) + PathDelim + EMBEDDED_LIBNAME;
end;

function TfrmMain.GetSourceHostName: string;
begin
  FSrcConnectionRec.HostName := Trim(edSrcHostName.Text);
  Result := FSrcConnectionRec.HostName;
end;

function TfrmMain.GetIsEmptyDatabase: Boolean;
begin
  FIsEmptyDatabase := chbEmptyDatabase.Enabled and chbEmptyDatabase.Checked;
  Result := FIsEmptyDatabase;
end;

function TfrmMain.GetMetadataOnly: Boolean;
begin
  FMetadataOnly := chbMetadataOnly.Enabled and chbMetadataOnly.Checked;
  Result := FMetadataOnly;
end;

function TfrmMain.GetSourcePassword: string;
begin
  FSrcConnectionRec.Password := Trim(edSrcPassword.Text);
  Result := FSrcConnectionRec.Password;
end;

function TfrmMain.GetSourcePort: uInt16;
begin
  FSrcConnectionRec.Port := StrToIntDef(edSrcPort.Text, DEFAULT_GDS_PORT);
  Result := FSrcConnectionRec.Port;
end;

function TfrmMain.GetSourceDataBaseName: string;
begin
  FSrcConnectionRec.DataBaseName := Trim(edSrcDatabase.Text);
  Result := FSrcConnectionRec.DataBaseName;
end;

function TfrmMain.GetSourceCharset: string;
begin
  if cbSrcCharset.ItemIndex < 0 then
    FSrcConnectionRec.Charset := NONE_CHARSET
  else
    FSrcConnectionRec.Charset := cbSrcCharset.Items[cbSrcCharset.ItemIndex];
  Result := FSrcConnectionRec.Charset;
end;

function TfrmMain.GetStrictCheck: Boolean;
begin
  FStrictCheck := chbStrictCheck.Enabled and chbStrictCheck.Checked;
  Result := FStrictCheck;
end;

function TfrmMain.GetSourceUserName: string;
begin
  FSrcConnectionRec.UserName := Trim(edSrcUser.Text);
  Result := FSrcConnectionRec.UserName;
end;

function TfrmMain.GetVerboseMode: Boolean;
begin
  FVerboseMode := chbVerboseMode.Enabled and chbVerboseMode.Checked;
  Result := FVerboseMode;
end;

procedure TfrmMain.SetConvertFromIntToBigInt(const Value: Boolean);
begin
  FConvertFromIntToBigInt := Value and chbIntToBigInt.Enabled;
  chbIntToBigInt.Checked := FConvertFromIntToBigInt;
end;

procedure TfrmMain.SetDestinationCharset(const Value: string);
var
  Idx: Integer;
begin
  FDstConnectionRec.Charset := Trim(Value);
  Idx := cbDstCharset.Items.IndexOf(FDstConnectionRec.Charset);
  if Idx < 0 then
    FDstConnectionRec.Charset := NONE_CHARSET;
  Idx := cbDstCharset.Items.IndexOf(FDstConnectionRec.Charset);
  cbDstCharset.ItemIndex := Idx;
end;

procedure TfrmMain.SetDestinationDataBaseName(const Value: string);
begin
  FDstConnectionRec.DataBaseName := Trim(Value);
  edDstDatabase.Text := FDstConnectionRec.DataBaseName;
end;

procedure TfrmMain.SetDestinationHostName(const Value: string);
begin
  FDstConnectionRec.HostName := Trim(Value);
  edDstHostName.Text := FDstConnectionRec.HostName;
end;

procedure TfrmMain.SetDestinationPassword(const Value: string);
begin
  FDstConnectionRec.Password := Trim(Value);
  edDstPassword.Text := FDstConnectionRec.Password;
end;

procedure TfrmMain.SetDestinationPort(const Value: uInt16);
begin
  FDstConnectionRec.Port := Value;
  edDstPort.Text := IntToStr(FDstConnectionRec.Port);
end;

procedure TfrmMain.SetDestinationUserName(const Value: string);
begin
  FDstConnectionRec.UserName := Trim(Value);
  edDstUser.Text := FDstConnectionRec.UserName;
end;

procedure TfrmMain.SetSourceHostName(const Value: string);
begin
  FSrcConnectionRec.HostName := Trim(Value);
  edSrcHostName.Text := FSrcConnectionRec.HostName;
end;

procedure TfrmMain.SetIsEmptyDatabase(const Value: Boolean);
begin
  FIsEmptyDatabase := Value and chbEmptyDatabase.Enabled;
  chbEmptyDatabase.Checked := FIsEmptyDatabase;
end;

procedure TfrmMain.SetMetadataOnly(const Value: Boolean);
begin
  FMetadataOnly := Value and chbMetadataOnly.Enabled;
  chbMetadataOnly.Checked := FMetadataOnly;
end;

procedure TfrmMain.SetSourcePassword(const Value: string);
begin
  FSrcConnectionRec.Password := Trim(Value);
  edSrcPassword.Text := FSrcConnectionRec.Password;
end;

procedure TfrmMain.SetSourcePort(const Value: uInt16);
begin
  FSrcConnectionRec.Port := Value;
  edSrcPort.Text := IntToStr(FSrcConnectionRec.Port);
end;

procedure TfrmMain.SetSourceDataBaseName(const Value: string);
begin
  FSrcConnectionRec.DataBaseName := Trim(Value);
  edSrcDatabase.Text := FSrcConnectionRec.DataBaseName;
end;

procedure TfrmMain.SetSourceCharset(const Value: string);
var
  Idx: Integer;
begin
  FSrcConnectionRec.Charset := Trim(Value);
  Idx := cbSrcCharset.Items.IndexOf(FSrcConnectionRec.Charset);
  if Idx < 0 then
    FSrcConnectionRec.Charset := NONE_CHARSET;
  Idx := cbSrcCharset.Items.IndexOf(FSrcConnectionRec.Charset);
  cbSrcCharset.ItemIndex := Idx;
end;

procedure TfrmMain.SetStrictCheck(const Value: Boolean);
begin
  FStrictCheck := Value and chbStrictCheck.Enabled;
  chbStrictCheck.Checked := FStrictCheck;
end;

procedure TfrmMain.SetSourceUserName(const Value: string);
begin
  FSrcConnectionRec.UserName := Trim(Value);
  edSrcUser.Text := FSrcConnectionRec.UserName;
end;

procedure TfrmMain.SetVerboseMode(const Value: Boolean);
begin
  FVerboseMode := Value and chbVerboseMode.Enabled;
  chbVerboseMode.Checked := FVerboseMode;
end;

// Actions
// -----------------------------------------------------------------------------

procedure TfrmMain.AL_AutoDetectExecute(Sender: TObject);
// 文字コードの自動取得
begin
  AutoDetectCharset(True);
end;

procedure TfrmMain.AL_CancelExecute(Sender: TObject);
// アクション: 終了
begin
  ExitProgram;
end;

procedure TfrmMain.AL_ConnectionTest_DstExecute(Sender: TObject);
// アクション: 接続テスト (Dst)
begin
  ConnectionTest(cttDestination);
end;

procedure TfrmMain.AL_ConnectionTest_SrcExecute(Sender: TObject);
// アクション: 接続テスト (Src)
begin
  ConnectionTest(cttSource);
end;

procedure TfrmMain.AL_ConvertExecute(Sender: TObject);
// アクション: コンバート
begin
  ProcessConvert;
end;

procedure TfrmMain.AL_InformationExecute(Sender: TObject);
// アクション: 情報
begin
  ShowInformation;
end;

procedure TfrmMain.AL_OpenDstFileExecute(Sender: TObject);
// アクション: ファイルを開く (コンバート先)
begin
  OpenDestinationFile;
end;

procedure TfrmMain.AL_OpenSrcFileExecute(Sender: TObject);
// アクション: ファイルを開く (コンバート元)
begin
  OpenSourceFile;
end;

// Functions
// -----------------------------------------------------------------------------

function TfrmMain.GetEmbeddedVersion: TFBVersion;
// 選択されている Embedded Server のバージョンを得る
begin
  if cbDstVersion.Items.Count = 0 then
    Result := fbv10
  else
    Result := TFBVersion(Integer(cbDstVersion.Items.Objects[cbDstVersion.ItemIndex]));
end;

procedure TfrmMain.SetDestinationCharSetList;
// Destination CharSet の変更
var
  i: Integer;
begin
  cbDstCharset.Clear;
  // Embedded Server が一つもインストールされていなければ抜ける
  if DestinationVersion = fbvUnknown then
    Exit;
  // コピー先のバージョンで使える文字コードのみを列挙
  for i:=Low(CharSets) to High(CharSets) do
    begin
      if (CharSets[i].MinimumVersion > GetEmbeddedVersion) then
        Continue;
      cbDstCharset.Items.Add(CharSets[i].Name);
    end;
  cbDstCharset.Sorted := True;
end;

function TfrmMain.CheckError: Boolean;
// エラーチェック
var
  Msg: string;
  OpenFlg: Boolean;

  { IsLocalHost BEGIN }
  function IsLocalHost(s: string): Boolean;
  begin
    result := (s = '') or (LowerCase(s) = 'localhost') or (s = '127.0.0.1');
  end;
  { IsLocalHost END }
begin
  Result := False;
  if IsLocalHost(Self.SourceHostName) then
    begin
      // コピー元データベースファイルの存在チェック
      if not TFile.Exists(Self.SourceDataBaseName) then
        begin
          MessageDlg(ERR_MSG_FILENOTEXISTS, TMsgDlgType.mtError, [TMsgDlgBtn.mbOK], -1);
          edSrcDatabase.SetFocus;
          Exit;
        end;

      // ネットワークパスかどうかのチェック
      if TPathEx.IsRemotePath(Self.SourceDataBaseName) then
        begin
          MessageDlg(ERR_MSG_REMOTEPATHSPECIFIED, TMsgDlgType.mtError, [TMsgDlgBtn.mbOK], -1);
          edSrcDatabase.SetFocus;
          Exit;
        end;
    end;

  // コピー元データベースのオープンチェック
  Msg := '';
  try
    SetFirebirdProperty(FBExtract.Firebird, True);
    with FBExtract.Firebird do
      begin
        Connect;
        Disconnect;
      end;
    OpenFlg := True;
  except
    on E: Exception do
      begin
        OpenFlg := False;
        Msg := E.Message;
      end;
  end;

  if not OpenFlg then
    begin
      MessageDlg(ERR_MSG_DATABASEOPENERROR + sLineBreak + sLineBreak + Msg, TMsgDlgType.mtError, [TMsgDlgBtn.mbOK], -1);
      edSrcDatabase.SetFocus;
      Exit;
    end;

  if (not MetadataOnly) then
    begin
      if IsLocalHost(Self.DestinationHostName) then
        begin
          // コピー先データベースフォルダの存在チェック
          if (TPath.GetDirectoryName(Self.DestinationDataBaseName) = '') or
             (not TDirectory.Exists(TPath.GetDirectoryName(Self.DestinationDataBaseName))) then
            begin
              MessageDlg(ERR_MSG_DIRECTORYNOTEXISTS, TMsgDlgType.mtError, [TMsgDlgBtn.mbOK], -1);
              edDstDatabase.SetFocus;
              Exit;
            end;
        end;

      // コピー元データベースファイルとコピー先データベースファイルの同一性チェック
      if  SameFileName(Self.SourceHostName + Self.SourceDataBaseName, Self.DestinationHostName + Self.DestinationDataBaseName) then
        begin
          MessageDlg(ERR_MSG_SAMEFILE, TMsgDlgType.mtError, [TMsgDlgBtn.mbOK], -1);
          edDstDatabase.SetFocus;
          Exit;
        end;

      // コピー先データベースが存在する場合の上書確認ダイアログ
      if IsLocalHost(Self.DestinationHostName) then
        if TFile.Exists(Self.DestinationDataBaseName) then
          begin
            if MessageDlg(CON_MSG_FILEEXIST, TMsgDlgType.mtConfirmation, [TMsgDlgBtn.mbYes, TMsgDlgBtn.mbNo], -1) <> ID_YES then
              begin
                edDstDatabase.SetFocus;
                Exit;
              end;
          end;
    end;
  Result := True;
end;

procedure TfrmMain.SetWorkDir(aDir: string);
// ワークフォルダ内を設定
begin
  FWorkDir := TPath.Combine(aDir, 'Firebird Converter'); // "C:\Users\<UserName>\AppData\Roaming\Firebird Converter"
  FBlobDir := TPath.Combine(FWorkDir, 'BLOB_DATA');
  if TDirectory.Exists(FWorkDir) then
    CleanupWorkDir
  else
    ForceDirectories(FBlobDir);
end;

procedure TfrmMain.CleanupWorkDir;
// ワークフォルダ内をクリーンナップ
begin
  if (TDirectory.Exists(Self.WorkDir)) and (Self.WorkDir <> '') then
    TDirectoryEx.Delete(Self.WorkDir, True);
  ForceDirectories(FBlobDir);
end;

procedure TfrmMain.ShowInformation;
// DB の情報表示
var
  InfoStr: string;
begin
  InfoStr := '';
  SetFirebirdProperty(FBExtract.Firebird, True);
  if (Self.SourceDataBaseName <> '') and (not TPathEx.IsRemotePath(Self.SourceDataBaseName)) then
    begin
      FDatabaseInfo := FBExtract.Firebird.GetDatabaseInfo;
      if FDatabaseInfo.Success then
        begin
          InfoStr := Format(MSG_ODS_VERSION, [FDatabaseInfo.GetODSString]) + sLineBreak +
                     Format(MSG_PAGE_SIZE  , [FDatabaseInfo.PageSize    ]) + sLineBreak +
                     Format(MSG_DIALECT    , [FDatabaseInfo.Dialect     ]);
        end
      else
        InfoStr := FDatabaseInfo.Msg;
    end;
  ShowMessage(InfoStr);
end;

procedure TfrmMain.SetSourceVersion(const Value: TFBVersionType);
// コピー元のバージョンを設定
var
  i: Integer;
  dVersionType: TFBVersionType;
begin
  FSrcConnectionRec.Version := fbvUnknown;
  cbSrcVersion.ItemIndex := -1;
  if cbSrcVersion.Items.Count = 0 then
    Exit;
  if (Value.ServerType = fbtUnknown) or (Value.Version = fbvUnknown) then
    Exit;
  for i:=0 to cbSrcVersion.Items.Count-1 do
    begin
      dVersionType := ValueToVersionType(uInt32(cbSrcVersion.Items.Objects[i]));
      if (dVersionType.ServerType = Value.ServerType) and (dVersionType.Version = Value.Version) then
        begin
          FSrcConnectionRec.Version := dVersionType.Version;
          cbSrcVersion.ItemIndex := i;
          Break;
        end;
    end;
end;

procedure TfrmMain.SetDestinationVersion(const Value: TFBVersionType);
// コピー先のバージョンを設定
var
  i: Integer;
  dVersionType: TFBVersionType;
begin
  FDstConnectionRec.Version := fbvUnknown;
  cbDstVersion.ItemIndex := -1;
  if cbDstVersion.Items.Count = 0 then
    Exit;
  if (Value.ServerType = fbtUnknown) or (Value.Version = fbvUnknown) then
    Exit;
  for i:=0 to cbDstVersion.Items.Count-1 do
    begin
      dVersionType := ValueToVersionType(uInt32(cbDstVersion.Items.Objects[i]));
      if (dVersionType.ServerType = Value.ServerType) and (dVersionType.Version = Value.Version) then
        begin
          FDstConnectionRec.Version := dVersionType.Version;
          cbDstVersion.ItemIndex := i;
          Break;
        end;
    end;
end;

function MakeCommentStr(aText: string; aCommentLine: Boolean): string;
// コメント文字列の生成
begin
  if aCommentLine and (aText <> '')  then
    Result := Format('/* %s */', [aText])
  else
    Result := aText;
end;

procedure AddError(aText: string; aStayLastLine: Boolean = False);
// ビューア (メッセージウィンドウ) へ書き出し
begin
  frmLogViewer.AddError(aText, aStayLastLine);
end;

procedure DeleteError;
// ビューア (メッセージウィンドウ) の最下行を削除
begin
  frmLogViewer.DeleteError;
end;

procedure AddMessage(aText: string; aCommentLine: Boolean = True; aStayLastLine: Boolean = False);
// ビューア (メタデータ) へ書き出し
begin
  frmLogViewer.AddMessage(MakeCommentStr(aText, aCommentLine), aStayLastLine);
end;

procedure DeleteMessage;
// ビューア (メタデータ) の最下行を削除
begin
  frmLogViewer.DeleteMessage;
end;

procedure TfrmMain.ProcessConvert;
// コンバート処理
const
  RELAXED_ALIAS_CHECKING = 'RelaxedAliasChecking';
  ROOT_DIRECTORY         = 'RootDirectory';
  REPLACE_DUMMY_STRING   = '@@DIR@@';
var
  CanMultiThread: Boolean;
  l, PageSize: Integer;
  ConfigFile, DataList: TStringList;
  ConfigFileName, CreateDBString, Dmy, s, OrgCharset,
  SQLFileName, FBLibName, FileSpec: string;
  F: TextFile;
  FBCharSet: TFBCharset;
  FBVer: TFBVersion;
  Firebird: TZeosFB;
  i: TIndexRec;
  IdxDynArr: TIndexDynArray;
  StartTime: TDateTime;
  StrDynArr, TblDynArr: TStringDynArray;
  t: TTriggerRec;
  TrgDynArr: TTriggerDynArray;

  { WriteError BEGIN }
  procedure WriteError(aText: string);
  // ビューア (エラーログ) へ書き出し
  var
    Err: TStringList;
    i: Integer;
  begin
    Err := TStringList.Create;
    try
      Err.Text := aText;
      for i:=0 to Min(MaxErrorLine - 1, Err.Count - 1) do // 長くなる事があるので最大行で制限する
        AddError(Err[i]);
      AddError('...');
      AddError('');
    finally
      Err.Free;
    end;
  end;
  { WriteError END }

  { AddSQLLog BEGIN }
  procedure AddSQLLog(aText: string; aCommentLine: Boolean = False);
  // SQL ログ (メタデータ) へ書き出し
  begin
    // DB を更新するか？
    if IsEmptyDatabase or ((not VerboseMode) and (not MetadataOnly))then
      Exit;
    Writeln(F, MakeCommentStr(aText, aCommentLine));
  end;
  { AddSQLLog END }

  { ProcessSQL BEGIN }
  procedure ProcessSQL(aSQL: string; aNeedTerminater: Boolean = False);
  // SQL 文の生成
  const
    Terminater: array [Boolean] of string = ('', ALT_TERM_CHAR);
  begin
    Application.ProcessMessages;
    // ビューア (メタデータ) へ書き出し
    AddMessage(aSQL + Terminater[aNeedTerminater], False);
    // DB を更新するか？
    if MetadataOnly then
      Exit;
    // SQL をコピー先の DB に対して実行
    with Firebird.Query[0] do
      begin
        SQL.Text := aSQL;
        try
          ExecSQL;
        except
          on E: EZSQLException do
            begin
              WriteError(E.Message);
//            raise Exception.Create(E.Message);
            end;
        end;
      end;
  end;
  { ProcessSQL END }

  { ProcessSQLwithParam BEGIN }
  procedure ProcessSQLwithParam(aSQL: string);
  // SQL 文の生成 (パラメータあり)
  var
    i: Integer;
    FileName, ParamName: string;
    BS: TBytesStream;
  begin
    Application.ProcessMessages;
    // DB を更新するか？
    if MetadataOnly then
      Exit;
    // SQL をコピー先の DB に対して実行 (パラメータ付)
    with Firebird.Query[1] do
      begin
        SQL.Text := aSQL;
        try
          BS := TBytesStream.Create;
          try
            for i:=0 to Params.Count-1 do
              begin
                ParamName := Params[i].Name;
                if ParamName[1] <> 'P' then
                  Continue;
                case Ord(ParamName[2]) of
                  Ord('T'):
                    begin
                      // TEXT データが存在する
                      FileName := TPath.Combine(BlobDir, StringReplace(Params[i].Name, 'PT', '', [rfReplaceAll]) + '.text');
                      Params[i].AsString := FileToString(FileName);
                    end;
                  Ord('B'):
                    begin
                      // BLOB データが存在する
                      FileName := TPath.Combine(BlobDir, StringReplace(Params[i].Name, 'PB', '', [rfReplaceAll]) + '.blob');
                      BS.LoadFromFile(FileName);
                      // サイズ 0 の BLOB データ (NULL ではない) は登録できない
                      // http://qc.embarcadero.com/wc/qcmain.aspx?d=88212
                      if BS.Size > 0 then
                        begin
                          BS.Seek(0, soFromBeginning);
                          Params[i].AsBlob := BS.Bytes;
                        end;
                    end;
                end;
              end;
          finally
            BS.Free;
          end;
          ExecSQL
        except
          on E: EZSQLException do
            WriteError(E.Message);
        end;
      end;
  end;
  { ProcessSQLwithParam END }

  { ImportTable BEGIN }
  procedure ImportTable(aTableName: string; aTableIndex: Integer; aVerboseMode: Boolean; aMetaDataOnly: Boolean; aReconnect: Boolean = False);
  // テーブルデータインポート
  var
    ImportStartTime: TDateTime;
    InsertSQL: TInsertSQL;
    InsRec: TInsertRec;
    InsRecArr: TInsRecDynArray;
    RecCnt: Int64;
    SplitDataRec: TSplitDataRec;
    Stream: TMemoryStream;
    BulkRec: PFBBulkRec;
    SplitSize: Integer;
    AsyncList: array of IAsyncCall;
    AsyncIdx: Integer;

    { GetEmptyAsyncIndex BEGIN }
    function GetEmptyAsyncIndex: Integer;
    var
      i: Integer;
    begin
      Result := -1;
      for i:= Low(AsyncList) to High(AsyncList) do
        begin
          if (AsyncList[i] = nil) or (AsyncList[i].Finished) then
            begin
              if (AsyncList[i] <> nil) then
                begin
                  AsyncList[i].Forget;
                  AsyncList[i] := nil;
                end;
              Result := i;
            end;
        end;
    end;
    { GetEmptyAsyncIndex END }
  begin
    if not aMetaDataOnly then
      Firebird.Disconnect;
    SplitSize := PostDataSplitSize;
    if CanMultiThread then
      begin
        SplitSize := PostDataSplitSize div NumberOfImportThread;
        SetLength(AsyncList, NumberOfImportThread);
      end;
    SplitDataRec := FBExtract.PreBuildTableData(FBExtract, aTableName, SplitSize);
    if SplitDataRec.RecordCount > 0 then
      begin
        ImportStartTime := Now;
        RecCnt := 0;
        Stream := TMemoryStream.Create;
        try
          AddMessage(Format('Import: %s%s', [aTableName, StringOfChar(' ', 32 - Length(aTableName))]));
          if aVerboseMode then
            AddSQLLog(Format('[TABLE: %s]', [aTableName]), True);
          while SplitDataRec.Index < SplitDataRec.SplitCount do
            begin
              if aVerboseMode or aMetaDataOnly then
                begin
                  // ----------------------------------------------
                  // SQL 文を生成してからインポート (Verbose Mode)
                  // ----------------------------------------------
                  // メタデータ用 Insert 文のビルド
                  InsRecArr := FBExtract.BuildTableData(FBExtract, aTableName, aTableIndex, SplitDataRec);
                  InsertSQL.INSERT := QuotedString(aTableName);
                  for InsRec in InsRecArr do
                    begin
                      InsertSQL.FIELDS  := InsRec.Fields;
                      InsertSQL.VALUES  := InsRec.Values1;
                      if InsRec.Comment = '' then
                        InsertSQL.COMMENT := ''
                      else
                        InsertSQL.COMMENT := ' ' + MakeCommentStr(InsRec.Comment, True);
                      AddSQLLog(InsertSQL.Build(True));
                    end;
                  AddSQLLog('');
                  // コンバート用 Insert 文のビルド
                  InsertSQL.COMMENT := '';
                  for InsRec in InsRecArr do
                    begin
                      InsertSQL.FIELDS := InsRec.Fields;
                      InsertSQL.VALUES := InsRec.Values2;
                      ProcessSQLwithParam(InsertSQL.Build);
                    end;
                end
              else
                begin
                  // ----------------------------------------------
                  // SQL 文を生成せずにインポート (Fast Mode)
                  // ----------------------------------------------
                  if CanMultiThread then
                    begin
                      // マルチスレッド時のインポート処理
                      AsyncIdx := GetEmptyAsyncIndex;
                      if AsyncIdx <> -1 then
                        begin
                          New(BulkRec);
                          BulkRec^.TableName     := aTableName;
                          BulkRec^.SplitDataRec  := SplitDataRec;
                          BulkRec^.SrcConRec     := FBExtract.Firebird.ConnectonInfo;
                          BulkRec^.DstConRec     := Firebird.ConnectonInfo;
                          BulkRec^.StartTime     := ImportStartTime;
                          BulkRec^.IsMultiThread := True;
                          SplitDataRec.Index     := SplitDataRec.Index + 1;
                          AsyncList[AsyncIdx]    := AsyncCall(@ImportData, TObject(BulkRec));
                        end;
                      Sleep(50);
                      // スレッドプールに空きがなければ空きができるまで待つ (マルチスレッド時)
                      while AsyncMultiSync(AsyncList, False) = WAIT_TIMEOUT do
                        Application.ProcessMessages;
                    end
                  else
                    begin
                      // シングルスレッド時のインポート処理
                      New(BulkRec);
                      BulkRec^.TableName     := aTableName;
                      BulkRec^.SplitDataRec  := SplitDataRec;
                      BulkRec^.SrcConRec     := FBExtract.Firebird.ConnectonInfo;
                      BulkRec^.DstConRec     := Firebird.ConnectonInfo;
                      BulkRec^.IsMultiThread := False;
                      SplitDataRec.Index     := SplitDataRec.Index + 1;
                      ImportData(BulkRec);
                    end;
                end;

              // 処理件数のカウントアップ
              if SplitDataRec.Index < SplitDataRec.SplitCount then
                Inc(RecCnt, SplitDataRec.SplitSize)
              else
                Inc(RecCnt, SplitDataRec.LastSize);

              // 処理件数の表示 (非マルチスレッド時)
              if not CanMultiThread then
                begin
                  Dmy := FormatDateTime('(HH:NN:SS)', Now - ImportStartTime);
                  AddMessage(Format(MSG_PROCESSED_RECORD, [FormatFloat('#,##0', RecCnt), Dmy]));
                end;
            end;
          // すべてのスレッドの終了を待つ (マルチスレッド時)
          if CanMultiThread then
            while AsyncMultiSync(AsyncList, True) = WAIT_TIMEOUT do
              Application.ProcessMessages;
        finally
          Stream.Free;
        end;
        AddMessage('');
      end;
    if (not aMetaDataOnly) and aReconnect then
      Firebird.Connect;
  end;
  { ImportTable END }

  { ChangeCharsetToNone BEGIN }
  procedure ChangeCharsetToNone;
  begin
    FBExtract.Firebird.Disconnect;
    FBExtract.Firebird.ClientCharset := 'NONE';
    FBExtract.Firebird.Connect;
  end;
  { ChangeCharsetToNone END }

  { ChangeCharsetToOrg BEGIN }
  procedure ChangeCharsetToOrg;
  begin
    FBExtract.Firebird.Disconnect;
    FBExtract.Firebird.ClientCharset := OrgCharset;
    FBExtract.Firebird.Connect;
  end;
  { ChangeCharsetToOrg END }
begin
  // エラーチェック
  if not CheckError then
    Exit;

  FIsConvertRunning := True;
  try
    // コピー先データベースが存在すれば削除
    if (not MetadataOnly) and TFile.Exists(Self.DestinationDataBaseName) then
      TFile.Delete(Self.DestinationDataBaseName);

    // ワークフォルダ内をクリーンナップ
    CleanupWorkDir;

    // コピー元に使用する Firebird の準備
    SetFirebirdProperty(FBExtract.Firebird, True);
    FBExtract.Init;
    FBCharSet := FBExtract.GetCharsetInfoByName(DestinationCharset);
    FBExtract.TargetCharset := FBCharSet;
    FBExtract.BrobOutputDir := Self.BlobDir;
    FBExtract.IsExtractBlobData := True;
    FBExtract.ConvertFromIntToBigInt := Self.ConvertFromIntToBigInt;

    // Client / Server のバージョンチェック
    FBVer := TZeosFB.VersionToFBVersion(FBExtract.Firebird.GetServerInfo.Version);
    if FBVer <> Self.SourceVersion then
      if MessageDlg(Format(CON_MSG_VERSIONMISMATCH1, [Format('[%s]' ,[SOURCE_STR]), TZeosFB.VersionToString(Self.SourceVersion), TZeosFB.VersionToString(FBVer)]),
        TMsgDlgType.mtConfirmation, [TMsgDlgBtn.mbYes, TMsgDlgBtn.mbNo], -1) <> ID_YES then
          begin
            cbSrcVersion.SetFocus;
            Exit;
          end;
    if FBExtract.Firebird.ServerCharset = 'CP943C' then
      FBExtract.Firebird.ClientCharset := 'UTF8';
    FBExtract.Firebird.Connect;

    // コピー元 DB の情報を得る
    FDatabaseInfo := FBExtract.Firebird.GetDatabaseInfo;

    // DB を更新するか？
    if not MetadataOnly then
      Firebird := TZeosFB.Create(Self);
    try
      CanMultiThread := False;
      if not MetadataOnly then
        begin
          // コピー先に使用する Firebird / Embedded の準備 (1)
          Firebird.Version       := Self.DestinationVersion;
          Firebird.ServerType    := Self.DestinationServerType;
          case Self.DestinationServerType of
            fbtClientServer:
              begin
                Firebird.LibraryLocation := Self.ClientDllName;
                Firebird.HostName        := Self.DestinationHostName;
                Firebird.Port            := Self.DestinationPort;
                Firebird.UserName        := Self.DestinationUserName;
                Firebird.Password        := Self.DestinationPassword;
              end;
            fbtEmbedded:
              begin
                Firebird.LibraryLocation := Self.EmbeddedDllName;
                Firebird.HostName        := '';
                Firebird.Port            := 0;
                Firebird.UserName        := DEFAULT_ADMIN_USERNAME;
                Firebird.Password        := DEFAULT_ADMIN_PASSWORD;
              end;
          end;
          Firebird.ServerCharset := Self.DestinationCharset;
          Firebird.ClientCharset := Self.DestinationCharset;

          // Firebird Embedded Server の準備
          // http://www.firebirdsql.org/manual/ufb-cs-embedded.html
          case Self.DestinationServerType of
            fbtEmbedded:
              begin
                // firebird.conf をコピー
                ConfigFileName := TPath.Combine(TPath.GetDirectoryName(ParamStr(0)), FB_CONFIG_FILE);
                TFile.Copy(TPath.Combine(TPath.GetDirectoryName(Self.EmbeddedDllName), FB_CONFIG_FILE),
                           ConfigFileName, True);
                // firebird.conf を書き換え
                if TFile.Exists(ConfigFileName) then
                  begin
                    ConfigFile := TStringList.Create;
                    try
                      ConfigFile.LoadFromFile(ConfigFileName);
                      // RelaxedAliasChecking
                      ConfigFile.Text := TRegEx.Replace(ConfigFile.Text,
                                                        Format('^[ #\t]*?%s[ \t]*?=.*?$', [RELAXED_ALIAS_CHECKING]),
                                                        Format('%s = %d', [RELAXED_ALIAS_CHECKING, Integer(not StrictCheck)]),
                                                        [roIgnoreCase, roMultiLine]);
                      // RootDirectory
                      ConfigFile.Text := TRegEx.Replace(ConfigFile.Text,
                                                        Format('^[ #\t]%s[ \t]*?=.*?$', [ROOT_DIRECTORY]),
                                                        Format('%s = %s', [ROOT_DIRECTORY, REPLACE_DUMMY_STRING]), // XE4 の正規表現置換にはバグがあるため
                                                        [roIgnoreCase, roMultiLine]);                              // 一旦ダミー文字列で置換する
                      Dmy := TPath.GetDirectoryName(Self.EmbeddedDllName);
                      ConfigFile.Text := StringReplace(ConfigFile.Text, REPLACE_DUMMY_STRING, Dmy, [rfReplaceAll]);
                      // Save
                      ConfigFile.SaveToFile(ConfigFileName);
                    finally
                      ConfigFile.Free;
                    end;
                  end;
              end;
            fbtClientServer:
              begin
                // Client / Server のバージョンチェック
                FBVer := TZeosFB.VersionToFBVersion(Firebird.GetServerInfo.Version);
                if FBVer <> Self.DestinationVersion then
                  if MessageDlg(Format(CON_MSG_VERSIONMISMATCH1, [Format('[%s]' ,[DESTINATON_STR]), TZeosFB.VersionToString(Self.DestinationVersion), TZeosFB.VersionToString(FBVer)]),
                    TMsgDlgType.mtConfirmation, [TMsgDlgBtn.mbYes, TMsgDlgBtn.mbNo], -1) <> ID_YES then
                      begin
                        cbDstVersion.SetFocus;
                        Exit;
                      end;
              end;
          end;

          // DB の新規作成
          try
            if (Self.DestinationVersion > fbv21) and (DatabaseInfo.PageSize < 4096) then
              PageSize := 4096                   // FB 2.5 以降では 4096 以下のページサイズは利用不可
            else
              PageSize := DatabaseInfo.PageSize; // オリジナルのページサイズを使う
            if Self.DestinationServerType = fbtClientServer then
              begin
                // Host/port:DatabaseName
                FileSpec := TZeosFB.BuildFileSpec(Self.DestinationHostName, Self.DestinationPort, Self.DestinationDataBaseName);
                // fbclient.dll
                FBLibName := Self.ClientDllName;
              end
            else
              begin
                // DatabaseName
                FileSpec := DestinationDataBaseName;
                // fbembed.dll
                FBLibName := Self.EmbeddedDllName;
              end;
            CreateDBString := Firebird.CreateDatabase
                                (Self.DestinationVersion, FBLibName,
                                 Self.DestinationHostName, Self.DestinationPort,
                                 FileSpec, Self.DestinationCharset,
                                 Self.DestinationUserName, Self.DestinationPassword,
                                 PageSize);
          except
            on E: EZSQLException do
              begin
                MessageDlg(ERR_MSG_CANTCREATEDATABASE + sLineBreak + sLineBreak + E.Message, TMsgDlgType.mtError, [TMsgDlgBtn.mbOK], -1);
                Exit;
              end;
          end;

          // コピー先に使用する Firebird / Embedded の準備 (2)
          Firebird.DatabaseName  := Self.DestinationDataBaseName;
          Firebird.Query[0].ParamCheck := False;
          if Firebird.ServerCharset = 'CP943C' then
            Firebird.ClientCharset := 'UTF8';
          Firebird.Connect;

          // マルチスレッド動作が可能か？
          CanMultiThread := (not VerboseMode) and                                                        // バーボーズモードでない
                            ((Firebird.Version >= fbv25) or (Firebird.ServerType = fbtClientServer)) and // Embedded なら 2.5 以上、C/S ならすべて
                            UseMultiThread;                                                              // IniFile の USE_MULTI_THREAD が 1 (規定値)
        end;

      // ログビューアの起動
      with frmLogViewer do
        begin
          Width  := 640;
          Height := 600;
          InitLog;
          Show;
          Application.ProcessMessages;
        end;

      // -----------------------------------------------------------------------
      //  SQL 文の作成と実行
      // -----------------------------------------------------------------------

      OrgCharset := FBExtract.Firebird.ClientCharset;

      StartTime := Now;
      AddError(MSG_STARTCONVERSION);
      AddError('');

      // SQL ログの準備
      if (not IsEmptyDatabase) and (VerboseMode or MetadataOnly) then
        begin
          SQLFileName := TPath.Combine(WorkDir, FB_METADATA_SQL);
          StringToFile('', SQLFileName, CP_UTF8);
          AssignFile(F, SQLFileName, CP_UTF8);
          Append(F); // 追記
        end;

      // データベース
      AddMessage('Database');
      AddMessage(Format('SET NAMES %s', [Self.DestinationCharset, sLineBreak]), False);
      AddMessage(Format('SET SQL DIALECT %d', [DEFAULT_DIALECT, sLineBreak]), False);
      if not MetadataOnly then
        begin
          AddMessage(CreateDBString);
          AddMessage('');
        end;

      // ドメイン
      StrDynArr := FBExtract.EnumDomains;
      if Length(StrDynArr) > 0 then
        begin
          AddMessage('Domains');
          for s in StrDynArr do
            ProcessSQL(FBExtract.BuildDomain(s));
          AddMessage('');
        end;

      // テーブル (定義: CREATE TABLE)
      TblDynArr := FBExtract.EnumTables;
      if Length(TblDynArr) > 0 then
        begin
          AddMessage('Tables');
          for s in TblDynArr do
            begin
              ProcessSQL(FBExtract.BuildTable(s));
              AddMessage('');
            end;
        end;

      // ストアドプロシージャ (定義: CREATE PROCEDURE)
      StrDynArr := FBExtract.EnumProcedures;
      if Length(StrDynArr) > 0 then
        begin
          AddMessage('Stored Procedures (Definition)');
          AddMessage(SET_TERM_BEGIN, False);
          AddMessage('');
          for s in StrDynArr do
            begin
              ProcessSQL(FBExtract.BuildProcedure(s, True), True);
              AddMessage('');
            end;
          AddMessage(SET_TERM_END, False);
          AddMessage(COMMIT_WORK, False);
          AddMessage('');
        end;

      // ビュー
      StrDynArr := FBExtract.EnumViews;
      if Length(StrDynArr) > 0 then
        begin
          AddMessage('Views');
          for s in StrDynArr do
            begin
              ProcessSQL(FBExtract.BuildView(s));
              AddMessage('');
            end;
          AddMessage('');
        end;

      //  ジェネレータ (定義: CREATE GENERATOR)
      StrDynArr := FBExtract.EnumGenerators;
      if Length(StrDynArr) > 0 then
        begin
          AddMessage('Generators (Definition)');
          for s in StrDynArr do
            ProcessSQL(FBExtract.BuildGenerator(s));
          AddMessage('');
        end;

      // UDF
      StrDynArr := FBExtract.EnumFunctions;
      if Length(StrDynArr) > 0 then
        begin
          AddMessage('User-Defined Functions');
          ChangeCharsetToNone;
          for s in StrDynArr do
            begin
              ProcessSQL(FBExtract.BuildFunction(s));
              AddMessage('');
            end;
          ChangeCharsetToOrg;
        end;

      //  例外
      StrDynArr := FBExtract.EnumExceptions;
      if Length(StrDynArr) > 0 then
        begin
          AddMessage('Exceptionss');
          for s in StrDynArr do
            ProcessSQL(FBExtract.BuildException(s));
          AddMessage('');
        end;

      // BLOB フィルタ
      StrDynArr := FBExtract.EnumFilters;
      if Length(StrDynArr) > 0 then
        begin
          ChangeCharsetToNone;
          AddMessage('Blob Filters');
          for s in StrDynArr do
            ProcessSQL(FBExtract.BuildFilter(s));
          AddMessage('');
          ChangeCharsetToOrg;
        end;

      // ロール
      StrDynArr := FBExtract.EnumRoles;
      if Length(StrDynArr) > 0 then
        begin
          AddMessage('Roles');
          for s in StrDynArr do
            ProcessSQL(FBExtract.BuildRole(s));
          AddMessage('');
        end;

      // ストアドプロシージャ (実装: ALTER PROCEDURE)
      StrDynArr := FBExtract.EnumProcedures;
      if Length(StrDynArr) > 0 then
        begin
          AddMessage('Stored Procedures (Implementation)');
          AddMessage(SET_TERM_BEGIN, False);
          AddMessage('');
          for s in StrDynArr do
            begin
              ProcessSQL(FBExtract.BuildProcedure(s, False), True);
              AddMessage('');
            end;
          AddMessage(SET_TERM_END, False);
          AddMessage(COMMIT_WORK, False);
          AddMessage('');
        end;

      if not IsEmptyDatabase then
        begin
          AddMessage(Format('Import SQL File: %s', [TPath.GetFileName(SQLFileName)]));
          AddMessage(StringOfChar('-', 50));
          AddMessage('');
          // テーブルデータ
          if Length(TblDynArr) > 0 then
            begin
              AddMessage('[TABLE DATA IMPORT]');
              AddMessage('');
              AddSQLLog('[TABLES]', True);
              AddSQLLog('');
              // データの抽出
              for l := Low(TblDynArr) to High(TblDynArr) do
                ImportTable(TblDynArr[l], l, VerboseMode, MetadataOnly);
              AddMessage('');
              AddSQLLog('');
              if not MetadataOnly then
                Firebird.Connect;
            end;

          //  ジェネレータ (値: SET GENERATOR)
          StrDynArr := FBExtract.EnumGenerators;
          if Length(StrDynArr) > 0 then
            begin
              AddMessage('[GENERATOR DATA IMPORT]');
              AddMessage('');
              AddSQLLog('[GENERATORS]', True);
              AddSQLLog('');
              for s in StrDynArr do
                begin
                  AddMessage(Format('  Import: %s%s', [s, StringOfChar(' ', 32 - Length(s))]));
                  Dmy := FBExtract.BuildGeneratorValue(s);
                  AddSQLLog(Dmy);
                  ProcessSQLwithParam(Dmy);
                end;
              AddMessage('');
              AddSQLLog('');
            end;
          AddMessage(StringOfChar('-', 50));
          AddMessage('');
        end;

      // テーブル (外部キー: ALTER TABLE)
      if Length(TblDynArr) > 0 then
        begin
          AddMessage('External Keys (Table)');
          for s in TblDynArr do
            begin
              DataList := TStringList.Create;
              try
                DataList.Text := FBExtract.BuildForeignKey(s);
                for Dmy in DataList do
                  begin
                    if Dmy <> '' then
                      ProcessSQL(Dmy);
                  end;
              finally
                DataList.Free;
              end;
            end;
          AddMessage('');
        end;

      // インデックス
      if Length(TblDynArr) > 0 then
        begin
          AddMessage('Indicies');
          for s in TblDynArr do
            begin
              IdxDynArr := FBExtract.EnumIndicies(s);
              for i in IdxDynArr do
                ProcessSQL(FBExtract.BuildIndex(i));
            end;
          AddMessage('');
        end;

      // トリガ
      if Length(TblDynArr) > 0 then
        begin
          AddMessage('Triggers');
          AddMessage(SET_TERM_BEGIN, False);
          AddMessage('');
          for s in TblDynArr do
            begin
              TrgDynArr := FBExtract.EnumTriggers(s);
              for t in TrgDynArr do
                begin
                  ProcessSQL(FBExtract.BuildTrigger(t), True);
                  AddMessage('');
                end;
            end;
          AddMessage(SET_TERM_END, False);
          AddMessage(COMMIT_WORK, False);
          AddMessage('');
        end;

      // 権限
      StrDynArr := FBExtract.EnumAndBuildGrant_Object;
      if Length(StrDynArr) > 0 then
        begin
          AddMessage('Grants (Object)');
          for s in StrDynArr do
            ProcessSQL(s);
          AddMessage('');
        end;
      StrDynArr := FBExtract.EnumAndBuildGrant_Execute;
      if Length(StrDynArr) > 0 then
        begin
          AddMessage('Grants (Procedure)');
          for s in StrDynArr do
            ProcessSQL(s);
          AddMessage('');
        end;
      StrDynArr := FBExtract.EnumAndBuildGrant_Role;
      if Length(StrDynArr) > 0 then
        begin
          AddMessage('Grants (Role)');
          for s in StrDynArr do
            ProcessSQL(s);
          AddMessage('');
        end;

      AddError(MSG_ENDCONVERSION);
      Dmy := FormatDateTime('(HH:NN:SS)', Now - StartTime);
      AddError(Dmy);

      // -----------------------------------------------------------------------

      // DB の切断
      FBExtract.Firebird.DisConnect;
      if not MetadataOnly then
        begin
          Firebird.Commit;
          Firebird.DisConnect;
        end;

      // ログの保存
      if (not IsEmptyDatabase) and (VerboseMode or MetadataOnly) then
        CloseFile(F);
      frmLogViewer.SaveMessage(TPath.Combine(WorkDir, FB_METADATA_FILE));
      frmLogViewer.SaveError(TPath.Combine(WorkDir, FB_ERROR_FILE));

    finally
      if not MetadataOnly then
        Firebird.Free;
    end;

    // ログビューアが閉じられていたらダイアログで終了を出す
    if not frmLogViewer.Showing then
      MessageDlg(INF_MSG_DONE + sLineBreak + Dmy, TMsgDlgType.mtInformation, [TMsgDlgBtn.mbOK], -1)
    // ログビューアが最小化/最大化されていたら元に戻す
    else if frmLogViewer.WindowState <>  wsMinimized then
      frmLogViewer.WindowState := wsNormal;
  finally
    FIsConvertRunning := False;
  end;
end;

procedure TfrmMain.SetFirebirdProperty(aFirebird: TZeosFB; IsSetCharset: Boolean; aDisConnect: Boolean);
// プロパティのセット
begin
  if aDisConnect then
    aFirebird.Disconnect;
  aFirebird.Version         := Self.SourceVersion;
  aFirebird.LibraryLocation := Self.SourceClientDllName;
  aFirebird.HostName        := Self.SourceHostName;
  aFirebird.Port            := Self.SourcePort;
  aFirebird.DatabaseName    := Self.SourceDataBaseName;
  aFirebird.UserName        := Self.SourceUserName;
  aFirebird.Password        := Self.SourcePassword;
  if IsSetCharset then
    aFirebird.ServerCharset := Self.SourceCharset;
end;

procedure TfrmMain.AutoDetectCharset(ShowErrorMsg: Boolean);
// 文字コードの自動取得
var
  Idx: Integer;
  Msg: string;
begin
  SetFirebirdProperty(FBExtract.Firebird);
  if (not ShowErrorMsg) and (not FBExtract.Firebird.SimpleConnectionCheck(Msg)) then
    Exit;
  FDatabaseInfo := FBExtract.Firebird.GetDatabaseInfo;
  if FDatabaseInfo.Success then
    begin
      FBExtract.Init;
      Idx := cbSrcCharset.Items.IndexOf(FBExtract.DefaultCharset.Name);
      if Idx < 0 then
        cbSrcCharset.ItemIndex := 0
      else
        cbSrcCharset.ItemIndex := Idx;
    end
  else
    begin
      if ShowErrorMsg then
        ShowMessage(FDatabaseInfo.Msg);
    end;
end;

procedure TfrmMain.ConnectionTest(aType: TConnectionTestType);
// 接続テスト
var
  Firebird: TZeosFB;
  ServerInfo: TFBServerInfo;
  ServerVersion: TFBVersion;
  VersionType: TFBVersionType;
begin
  Firebird := TZeosFB.Create(Self);
  try
    case aType of
      cttSource:
        begin
          Firebird.Version         := Self.SourceVersion;
          Firebird.LibraryLocation := Self.SourceClientDllName;
          Firebird.HostName        := Self.SourceHostName;
          Firebird.Port            := Self.SourcePort;
          Firebird.UserName        := Self.SourceUserName;
          Firebird.Password        := Self.SourcePassword;
        end;
      cttDestination:
        begin
          Firebird.Version         := Self.DestinationVersion;
          Firebird.LibraryLocation := Self.ClientDllName;
          Firebird.HostName        := Self.DestinationHostName;
          Firebird.Port            := Self.DestinationPort;
          Firebird.UserName        := Self.DestinationUserName;
          Firebird.Password        := Self.DestinationPassword;
        end;
    end;
    ServerInfo := Firebird.GetServerInfo;
    if ServerInfo.Success then
      begin
        // 接続 OK
        ServerVersion := TZeosFB.VersionToFBVersion(ServerInfo.Version);
        if ServerVersion <> Firebird.Version then
          begin
            // サーバのバージョンとクライアントのバージョンが異なる
            if MessageDlg(Format(CON_MSG_VERSIONMISMATCH2, [MSG_CONNECTON_SUCCESS2, TZeosFB.VersionToString(Firebird.Version), TZeosFB.VersionToString(ServerVersion)]),
               TMsgDlgType.mtConfirmation, [TMsgDlgBtn.mbYes, TMsgDlgBtn.mbNo], -1) = ID_YES then
              begin
                VersionType.ServerType := fbtClientServer;
                VersionType.Version    := ServerVersion;
                case aType of
                  cttSource:
                    SetSourceVersion(VersionType);
                  cttDestination:
                    SetDestinationVersion(VersionType);
                end;
              end;
          end
        else
          ShowMessage(MSG_CONNECTON_SUCCESS);
      end
    else
      begin
        // 接続 NG
        ShowMessage(ServerInfo.Msg);
      end;
  finally
    Firebird.Free;
  end;
end;

procedure TfrmMain.ExitProgram;
// 終了処理
begin
  Close;
end;

procedure TfrmMain.OpenDestinationFile;
// コンバート先データベースファイルの指定
begin
  if not sdFile.Execute then
    Exit;
  DestinationDataBaseName := sdFile.FileName;
end;

procedure TfrmMain.OpenSourceFile;
// コンバート元データベースファイルの指定
begin
  if not odFile.Execute then
    Exit;
  SourceDataBaseName := odFile.FileName;
  Application.ProcessMessages;
  if TFile.Exists(SourceDataBaseName) then
    AutoDetectCharset(False);
end;

function TfrmMain.ValueToVersionType(aValue: uInt32): TFBVersionType;
// uInt32 から TFBVersionType へ変換
begin
  Result.ServerType := TFBServerType(HiWord(aValue));
  Result.Version    := TFBVersion(LoWord(aValue));
end;

function TfrmMain.VersionTypeToValue(aVersuionType: TFBVersionType): uInt32;
// TFBVersionType から uInt32 へ変換
begin
  Result := (uInt32(aVersuionType.ServerType) shl 16) + uInt32(aVersuionType.Version);
end;

procedure SyncAddMessage(aMsg: PMessageRec);
// メッセージの更新
var
  Dmy: string;
begin
  Dmy := Format(MSG_PROCESSED_RECORD_BLOCK,
    [FormatFloat('#,##0', aMsg^.StartBlock),
     FormatFloat('#,##0', aMsg^.EndBlock),
     FormatDateTime('(HH:NN:SS)', Now - aMsg^.StartTime)]);
  AddMessage(Dmy);
  Dispose(aMsg);
end;

procedure ImportData(aBulkRec: PFBBulkRec);
// インポート (Fast Mode)
var
  DstFB, SrcFB: TZeosFB;
  DstQuery, SrcQuery: TZQuery;
  SelectSQL: TSelectSQL;
  InsertSQL: TInsertSQL;
  dFields, dValues: string;
  i: Integer;
  Stream: TMemoryStream;
  MsgRec: PMessageRec;
  BulkRec: TFBBulkRec;
begin
  BulkRec := aBulkRec^;
  SrcFB := TZeosFB.Create(nil);
  DstFB := TZeosFB.Create(nil);
  Stream := TMemoryStream.Create;
  try
    SrcFB.ConnectonInfo := BulkRec.SrcConRec;
    SrcFB.Connect;
    SrcQuery := SrcFB.Query[0];

    DstFB.Connection.AutoCommit := False;
    DstFB.ConnectonInfo := BulkRec.DstConRec;
    DstFB.Connect;
    DstQuery := DstFB.Query[0];

    DstQuery.SQL.Text := '';
    // コピー元テーブルの準備
    SelectSQL.SELECT := Format('FIRST %d SKIP %d *',
                          [BulkRec.SplitDataRec.SplitSize, BulkRec.SplitDataRec.Index * BulkRec.SplitDataRec.SplitSize]);
    SelectSQL.FROM   := QuotedString(BulkRec.TableName);
    SrcQuery.SQL.Text := SelectSQL.Build;
    SrcQuery.Open;
    if not SrcQuery.IsEmpty then
      begin
        if DstQuery.SQL.Text = '' then
          begin
            // コピー先テーブルの準備
            InsertSQL.INSERT := QuotedString(BulkRec.TableName);
            dFields := '';
            dValues := '';
            for i:=0 to SrcQuery.FieldCount-1 do
              begin
                dFields := dFields + QuotedString(SrcQuery.Fields[i].FieldName);
                dValues := dValues + Format(':P%.4d', [i]);
                if i < SrcQuery.FieldCount-1 then
                  begin
                    dFields := dFields + SEPARATOR_STR;
                    dValues := dValues + SEPARATOR_STR;
                  end;
              end;
            InsertSQL.FIELDS := dFields;
            InsertSQL.VALUES := dValues;
            DstQuery.SQL.Text := InsertSQL.Build;
          end;
        // インポート
        DstFB.StartTransaction;
        while not SrcQuery.Eof do
          begin
            for i:=0 to SrcQuery.FieldCount-1 do
              begin
                if SrcQuery.Fields[i].IsNull then
                  begin
                    DstQuery.Params[i].Clear;
                    Continue;
                  end;    
                case SrcQuery.Fields[i].DataType of
                  ftBoolean:
                    DstQuery.Params[i].AsBoolean    := SrcQuery.Fields[i].AsBoolean;
                  ftByte, ftSmallInt, ftWord, ftInteger:
                    DstQuery.Params[i].AsInteger    := SrcQuery.Fields[i].AsInteger;
                  ftLongWord:
                    DstQuery.Params[i].AsLongWord   := SrcQuery.Fields[i].AsLongWord;
                  ftLargeint:
                    DstQuery.Params[i].AsLargeInt   := SrcQuery.Fields[i].AsLargeInt;
                  ftSingle:
                    DstQuery.Params[i].AsSingle     := SrcQuery.Fields[i].AsSingle;
                  ftExtended, ftFloat:
                    DstQuery.Params[i].AsFloat      := SrcQuery.Fields[i].AsFloat;
                  ftCurrency, ftBCD:
                    DstQuery.Params[i].AsCurrency   := SrcQuery.Fields[i].AsCurrency;
                  ftFMTBcd:
                    DstQuery.Params[i].AsFMTBCD     := SrcQuery.Fields[i].AsBCD;
                  ftString, ftMemo:
                    DstQuery.Params[i].AsAnsiString := SrcQuery.Fields[i].AsAnsiString;
                  ftWideString, ftWideMemo:
                      DstQuery.Params[i].AsString   := SrcQuery.Fields[i].AsString;
                  ftDate, ftTime, ftDateTime, ftTimeStamp:
                    DstQuery.Params[i].AsDateTime   := SrcQuery.Fields[i].AsDateTime;
                  ftBlob:
                    begin
                      Stream.Clear;
                      (SrcQuery.Fields[i] as TBlobField).SaveToStream(Stream);
                      // サイズ 0 の BLOB データ (NULL ではない) は登録できない
                      // http://qc.embarcadero.com/wc/qcmain.aspx?d=88212
                      if Stream.Size > 0  then
                        DstQuery.Params[i].LoadFromStream(Stream, ftBlob);
                    end
                else
                  raise Exception.Create(MSG_UNKNOWN_DATATYPE);
                end;
              end;
            DstQuery.ExecSQL;
            SrcQuery.Next;
          end;
      end;
    SrcQuery.Close;
    DstFB.Commit;
    DstFB.Disconnect;
    SrcFB.Disconnect;
  finally
    Stream.Free;
    DstFB.Free;
    SrcFB.Free;
    Dispose(aBulkRec);
  end;

  if not BulkRec.IsMultiThread then
    Exit;

  // 処理済メッセージの出力
  New(MsgRec);
  MsgRec^.StartBlock := BulkRec.SplitDataRec.Index * BulkRec.SplitDataRec.SplitSize + 1;
  MsgRec^.EndBlock   := MsgRec^.StartBlock - 1;
  if BulkRec.SplitDataRec.Index < (BulkRec.SplitDataRec.SplitCount - 1) then
    MsgRec^.EndBlock := MsgRec^.EndBlock + BulkRec.SplitDataRec.SplitSize
  else
    MsgRec^.EndBlock := MsgRec^.EndBlock + BulkRec.SplitDataRec.LastSize;
  MsgRec^.StartTime := BulkRec.StartTime;

  if GetCurrentThreadId <> MainThreadID then
    TThread.Synchronize(nil, procedure begin SyncAddMessage(MsgRec) end)
  else
    SyncAddMessage(MsgRec);
end;

end.
