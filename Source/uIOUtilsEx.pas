{*******************************************************}
{                                                       }
{            File / Directory / Path Utility            }
{                                                       }
{    Copyright (c) 2014-2015 Hideaki Tominaga (DEKO)    }
{                                                       }
{*******************************************************}
// -----------------------------------------------------------------------------
// Note:
//   - File / Directory / Path Utility Unit
// -----------------------------------------------------------------------------
unit uIOUtilsEx;
{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  SysUtils, Types, IOUtils, Diagnostics{$IFDEF MSWINDOWS}, ShlObj, Windows{$ENDIF};

type
  TDriveEx = record
  public
    class function DefineLocalDevice(const DeviceName: string;
      const TargetPath: string): Boolean; static;
    class function UndefLocalDevice(const DeviceName: string;
      const TargetPath: string = ''): Boolean; static;
    class function DefineRemoteDevice(const DeviceName: string;
      const TargetPath: string; const Username: string = '';
      const Password: string = ''; const Persistent: Boolean = False): Boolean; static;
    class function UndefRemoteDevice(const DeviceName: string;
      const TargetPath: string = ''): Boolean; static;
  end;

  TDirectoryEx = record
  public
    class procedure Copy(const SourceDirName, DestDirName: string;
      Persistent: Boolean = False); static;
    class procedure Delete(const Path: string; const Recursive: Boolean = True;
      Persistent: Boolean = False); static;
    class procedure Move(const SourceDirName, DestDirName: string;
      Persistent: Boolean = False); static;
  end;

  TFileEx = record
  public
    class procedure Copy(const SourceFileName, DestFileName: string;
        const Overwrite: Boolean = True); overload; static;
    class procedure CopyFiles(const SrcDir, SearchPattern, DstDir: string;
      const Recursive: Boolean = False); static;
    class procedure Delete(const FileName: string); static;
    class procedure DeleteFiles(const Dir, SearchPattern: string;
      const Recursive: Boolean = False); static;
    class procedure Move(SourceFileName, DestFileName: string); static;
    class procedure MoveFiles(const SrcDir, SearchPattern, DstDir: string;
      const Recursive: Boolean = False); static;
  end;

  TPathEx = record
  private
    const
      UNC_PREFIX          = '\\';
      EXTENDED_PREFIX     = UNC_PREFIX + '?\';
      EXTENDED_UNC_PREFIX = EXTENDED_PREFIX + 'UNC\';
  private
    class function PathDelimCount(const Path: string): Integer; static;
  public
    class function Combine(const Paths: array of string): string; static;
    class function GetExtendedPath(const Path: string): string; static;
    class function GetParentDir(const Path: string): string; static;
    class function GetParentPath(const Path: string): string; static;
    class function GetPathDepth(const Path: string): Integer; static;
    class function GetReducedPath(const Path: string): string; static;
    class function GetSubPath(const Path: string): string; static;
    class function IsBasePath(const BasePath, Path: string): Boolean; static;
    class function IsExtendedPath(const Path: string): Boolean; static;
    class function IsExtendedUNCPath(const Path: string): Boolean; static;
    class function IsRemotePath(const Path: string): Boolean; static;
    class function IsRootPath(const Path: string): Boolean; static;
    class function IsTrailingPathDelimiter(const Path: string): Boolean; static;
  end;

implementation

{ TDriveEx }

// 任意のローカルフォルダをドライブに割り当てる (SUBST 相当)
// -----------------------------------------------------------------------------
// <!> リモートフォルダを割り当てる事はできない <!>
// ※実はできるのだが [切断されたネットワークドライブ] として認識されてしまう。
// -----------------------------------------------------------------------------
//  DeviceName: ドライブ文字列 (Ex Z:)
//  TargetPath: 割り当てるパス
// -----------------------------------------------------------------------------
//  Result: 成功すれば True
// -----------------------------------------------------------------------------
class function TDriveEx.DefineLocalDevice(const DeviceName,
  TargetPath: string): Boolean;
{$IFDEF MSWINDOWS}
var
  dTargetPath: string;
{$ENDIF}
begin
  {$IFDEF MSWINDOWS}
  dTargetPath := ExcludeTrailingPathDelimiter(Trim(TargetPath));
  if TPath.IsUNCRooted(dTargetPath) then
    dTargetPath := TPathEx.GetExtendedPath(dTargetPath);
  Result := DefineDosDevice(0, PChar(DeviceName), PChar(dTargetPath));
  if Result then
    SHChangeNotify(SHCNE_DRIVEADD, SHCNF_FLUSH, nil, nil);
  {$ELSE}
  Result := False;
  {$ENDIF}
end;

// 任意のリモートフォルダをネットワークドライブに割り当てる (net use 相当)
// -----------------------------------------------------------------------------
// <!> ローカルフォルダを割り当てる事はできない <!>
// -----------------------------------------------------------------------------
//  DeviceName: ドライブ文字列 (Ex Z:)
//  TargetPath: 割り当てるパス
//  Username  : ユーザー名
//  Password  : パスワード
//  Persistent: True だと Windows 再起動時に再接続を試みる
// -----------------------------------------------------------------------------
//  Result: 成功すれば True
// -----------------------------------------------------------------------------
class function TDriveEx.DefineRemoteDevice(const DeviceName, TargetPath,
  Username, Password: string; const Persistent: Boolean): Boolean;
{$IFDEF MSWINDOWS}
var
  dTargetPath: string;
  NR: TNETRESOURCE;
  Flags: DWORD;
{$ENDIF}
begin
  {$IFDEF MSWINDOWS}
  dTargetPath := ExcludeTrailingPathDelimiter(TPathEx.GetReducedPath(Trim(TargetPath)));
  with NR do
    begin
      dwType := RESOURCETYPE_DISK;
      lpLocalName := PChar(DeviceName);
      lpRemoteName := PChar(dTargetPath);
      lpProvider := nil;
    end;
  if Persistent then
    Flags := CONNECT_UPDATE_PROFILE
  else
    Flags := 0;
  Result := (WnetAddConnection2(NR, PChar(Password), PChar(Username), Flags) = NO_ERROR);
  if Result then
    SHChangeNotify(SHCNE_DRIVEADD, SHCNF_FLUSH, nil, nil);
  {$ELSE}
  Result := False;
  {$ENDIF}
end;

// 割り当てたドライブを解除する (SUBST /D 相当)
// -----------------------------------------------------------------------------
//  DeviceName: ドライブ文字列 (Ex Z:)
//  TargetPath: 割り当てたパス
// -----------------------------------------------------------------------------
//  Result: 成功すれば True
// -----------------------------------------------------------------------------
class function TDriveEx.UndefLocalDevice(const DeviceName: string;
  const TargetPath: string): Boolean;
{$IFDEF MSWINDOWS}
var
  Flags: DWORD;
  dTargetPath: string;
  PTargetPath: PChar;
{$ENDIF}
begin
  {$IFDEF MSWINDOWS}
  dTargetPath := ExcludeTrailingPathDelimiter(Trim(TargetPath));
  Flags := DDD_REMOVE_DEFINITION;
  if TargetPath = '' then
    PTargetPath := nil
  else
    begin
      Flags := Flags or DDD_EXACT_MATCH_ON_REMOVE;
      if TPath.IsUNCRooted(dTargetPath) then
        dTargetPath := TPathEx.GetExtendedPath(dTargetPath);
      PTargetPath := PChar(dTargetPath);
    end;
  Result := DefineDosDevice(Flags, PChar(DeviceName), PTargetPath);
  if Result then
    SHChangeNotify(SHCNE_DRIVEREMOVED, SHCNF_FLUSH, nil, nil);
  {$ELSE}
  Result := False;
  {$ENDIF}
end;

// 割り当てたネットワークドライブを解除する (net use 相当)
// -----------------------------------------------------------------------------
//  DeviceName: ドライブ文字列 (Ex Z:)
//  TargetPath: 割り当てたパス (未使用)
// -----------------------------------------------------------------------------
//  Result: 成功すれば True
// -----------------------------------------------------------------------------
class function TDriveEx.UndefRemoteDevice(const DeviceName: string;
  const TargetPath: string): Boolean;
begin
  {$IFDEF MSWINDOWS}
  Result := (WNetCancelConnection2(PChar(DeviceName), 0, True) = NO_ERROR);
  if Result then
    SHChangeNotify(SHCNE_DRIVEREMOVED, SHCNF_FLUSH, nil, nil);
  {$ELSE}
  Result := False;
  {$ENDIF}
end;

{ TDirectoryEx }

// フォルダをコピーする (コピー先フォルダ自動作成)
// -----------------------------------------------------------------------------
// 単一ファイルのコピー -> TFileEx.Copy()
// ファイルのコピー     -> TFileEx.CopyFiles()
// -----------------------------------------------------------------------------
//  SourceDirName: コピー元ファイル
//  DestDirName  : コピー先ファイル
//  Persistent   : 可能な限りフォルダコピーを試みるか？ (予約)
// -----------------------------------------------------------------------------
//  Result: (なし)
// -----------------------------------------------------------------------------
class procedure TDirectoryEx.Copy(const SourceDirName, DestDirName: string;
  Persistent: Boolean);
begin
  // ソースフォルダが存在しなければ抜ける
  if not TDirectory.Exists(SourceDirName) then
    Exit;
  // フォルダを作成
  if not TDirectory.Exists(DestDirName) then
    SysUtils.ForceDirectories(DestDirName);
  // フォルダをコピー
  TDirectory.Copy(SourceDirName, DestDirName);
end;

// フォルダを削除する
// -----------------------------------------------------------------------------
// 単一ファイルの削除 -> TFileEx.Delete()
// ファイルの削除     -> TFileEx.DeleteFiles()
// -----------------------------------------------------------------------------
//  Path      : 削除フォルダ
//  Recursive : サブフォルダも削除するか？
//  Persistent: 可能な限りフォルダ削除を試みるか？ (予約)
// -----------------------------------------------------------------------------
//  Result: (なし)
// -----------------------------------------------------------------------------
class procedure TDirectoryEx.Delete(const Path: string; const Recursive: Boolean;
  Persistent: Boolean);
begin
  // ソースフォルダが存在しなければ抜ける
  if not TDirectory.Exists(Path) then
    Exit;
  // フォルダを削除
  TDirectory.Delete(Path, Recursive);
end;

// フォルダを移動する
// -----------------------------------------------------------------------------
// 単一ファイルの移動 -> TFileEx.Move()
// ファイルの移動     -> TFileEx.MoveFiles()
// -----------------------------------------------------------------------------
//  SourceDirName: 移動元ファイル
//  DestDirName  : 移動先ファイル
//  Persistent   : 可能な限りフォルダ移動を試みるか？
// -----------------------------------------------------------------------------
//  Result: (なし)
// -----------------------------------------------------------------------------
class procedure TDirectoryEx.Move(const SourceDirName, DestDirName: string;
  Persistent: Boolean);
begin
  // ソースフォルダが存在しなければ抜ける
  if not TDirectory.Exists(SourceDirName) then
    Exit;
  if TDirectory.Exists(DestDirName) and Persistent then
    begin
      // フォルダをコピー
      TDirectory.Copy(SourceDirName, DestDirName);
      // フォルダを削除
      TDirectory.Delete(SourceDirName, True);
    end
  else
    begin
      // フォルダを移動
      TDirectory.Move(SourceDirName, DestDirName);
    end;
end;

{ TFileEx }

// 単一ファイルをコピーする (コピー先フォルダ自動作成)
// -----------------------------------------------------------------------------
// ファイルのコピー -> TFileEx.CopyFiles()
// フォルダコピー   -> TDirectoryEx.Copy()
// -----------------------------------------------------------------------------
//  SourceFileName: コピー元ファイル
//  DestFileName  : コピー先ファイル
//  Overwrite     : 上書きするか？(デフォルト: True)
// -----------------------------------------------------------------------------
//  Result: (なし)
// -----------------------------------------------------------------------------
class procedure TFileEx.Copy(const SourceFileName, DestFileName: string;
  const Overwrite: Boolean);
begin
  // ソースファイルが存在しなければ抜ける
  if not TFile.Exists(SourceFileName) then
    Exit;
  // フォルダを作成
  if not TDirectory.Exists(TPath.GeTDirectoryName(DestFileName)) then
    SysUtils.ForceDirectories(TPath.GeTDirectoryName(DestFileName))
  // 属性変更
  else if Overwrite and TFile.Exists(DestFileName) then
    TFile.SetAttributes(DestFileName, [TFileAttribute.faNormal]);
  // ファイルコピー
  TFile.Copy(SourceFileName, DestFileName, Overwrite);
end;

// ファイルをコピーする (ワイルドカード対応 / コピー先フォルダ自動作成)
// -----------------------------------------------------------------------------
// 単一ファイルのコピー -> TFileEx.Copy()
// フォルダコピー       -> TDirectoryEx.Copy()
// -----------------------------------------------------------------------------
//  SrcDir       : コピー元フォルダ
//  SearchPattern: ファイルの検索パターン (*.* 等)
//  DstDir       : コピー先フォルダ
//  Recursive    : サブフォルダのファイルも対象か？
// -----------------------------------------------------------------------------
//  Result: (なし)
// -----------------------------------------------------------------------------
class procedure TFileEx.CopyFiles(const SrcDir, SearchPattern, DstDir: string;
  const Recursive: Boolean);
var
  OrgSrcDir, OrgDstDir: string;
  procedure _CopyFiles(const SrcDir, DstDir: string);
  var
    Directory, FileName: string;
    dDstDir, dDstFileName: string;
  begin
    // フォルダを作成
    if not TDirectory.Exists(DstDir) then
      SysUtils.ForceDirectories(DstDir);
    // ファイルをコピー
    for FileName in TDirectory.GetFiles(SrcDir, SearchPattern) do
      begin
        dDstFileName := SysUtils.IncludeTrailingPathDelimiter(DstDir) + TPath.GetFileName(FileName);
        if TFile.Exists(dDstFileName) then
          TFile.SetAttributes(dDstFileName, [TFileAttribute.faNormal]);
        TFile.Copy(FileName, dDstFileName , True);
      end;
    // サブディレクトリを走査
    if Recursive then
      for Directory in TDirectory.GetDirectories(SrcDir) do
        begin
          if SameFileName(Directory, dDstDir) then
            Continue;
          dDstDir := SysUtils.IncludeTrailingPathDelimiter(DstDir) + TPathEx.GetSubPath(Directory);
          _CopyFiles(Directory, dDstDir);
        end;
  end;
begin
  OrgSrcDir := SysUtils.ExcludeTrailingPathDelimiter(SrcDir);
  // ソースフォルダが存在しなければ抜ける
  if not TDirectory.Exists(OrgSrcDir) then
    Exit;
  if not TDirectory.Exists(OrgSrcDir) then
    Exit;
  OrgDstDir := SysUtils.ExcludeTrailingPathDelimiter(DstDir);
  _CopyFiles(OrgSrcDir, OrgDstDir);
end;

// 単一ファイルを削除する
// -----------------------------------------------------------------------------
// ファイルの削除 -> TFileEx.DeleteFiles()
// フォルダ削除   -> TDirectoryEx.Delete()
// -----------------------------------------------------------------------------
//  FileName: 削除するファイル
// -----------------------------------------------------------------------------
//  Result: (なし)
// -----------------------------------------------------------------------------
class procedure TFileEx.Delete(const FileName: string);
begin
  // ソースファイルが存在しなければ抜ける
  if not TFile.Exists(FileName) then
    Exit;
  // 属性変更
  TFile.SetAttributes(FileName, [TFileAttribute.faNormal]);
  // ファイルコピー
  TFile.Delete(FileName);
end;

// ファイルを削除する (ワイルドカード対応 / コピー元フォルダは削除しない)
// -----------------------------------------------------------------------------
// 単一ファイルの削除ー -> TFileEx.Delete()
// フォルダ削除         -> TDirectoryEx.Delete()
// -----------------------------------------------------------------------------
//  Dir          : 削除対象フォルダ
//  SearchPattern: ファイルの検索パターン (*.* 等)
//  Recursive    : サブフォルダのファイルも対象か？
// -----------------------------------------------------------------------------
//  Result: (なし)
// -----------------------------------------------------------------------------
class procedure TFileEx.DeleteFiles(const Dir, SearchPattern: string;
  const Recursive: Boolean);
var
  OrgDir: string;
  procedure _DeleteFiles(const Dir: string);
  var
    Directory, FileName: string;
  begin
    // ファイルを削除
    for FileName in TDirectory.GetFiles(Dir, SearchPattern) do
      begin
        TFile.SetAttributes(FileName, [TFileAttribute.faNormal]);
        TFile.Delete(FileName);
      end;
    // サブディレクトリを走査
    if Recursive then
      for Directory in TDirectory.GetDirectories(Dir) do
        _DeleteFiles(Directory);
  end;
begin
  OrgDir := SysUtils.ExcludeTrailingPathDelimiter(Dir);
  // ソースフォルダが存在しなければ抜ける
  if not TDirectory.Exists(OrgDir) then
    Exit;
  _DeleteFiles(OrgDir);
end;

// 単一ファイルを移動する (移動先フォルダ自動作成)
// -----------------------------------------------------------------------------
// ファイルの移動 -> TFileEx.MoveFiles()
// フォルダ移動   -> TDirectoryEx.Move()
// -----------------------------------------------------------------------------
//  SourceFileName: 移動ファイル
//  DestFileName  : 移動先ファイル
// -----------------------------------------------------------------------------
//  Result: (なし)
// -----------------------------------------------------------------------------
class procedure TFileEx.Move(SourceFileName, DestFileName: string);
begin
  // ソースファイルが存在しなければ抜ける
  if not TFile.Exists(SourceFileName) then
    Exit;
  // フォルダを作成
  if not TDirectory.Exists(TPath.GeTDirectoryName(DestFileName)) then
    SysUtils.ForceDirectories(TPath.GeTDirectoryName(DestFileName))
  else if TFile.Exists(DestFileName) then
    begin
      // 属性変更
      TFile.SetAttributes(DestFileName, [TFileAttribute.faNormal]);
      // 削除
      TFile.Delete(DestFileName);
    end;
  // ファイル移動
  TFile.Move(SourceFileName, DestFileName);
end;

// ファイルを移動する (ワイルドカード対応 / 移動先フォルダは削除しない / 移動先フォルダ自動作成)
// -----------------------------------------------------------------------------
// 単一ファイルの移動 -> TFileEx.Move()
// フォルダ移動       -> TDirectoryEx.Move()
// -----------------------------------------------------------------------------
//  SrcDir       : 移動元フォルダ
//  SearchPattern: ファイルの検索パターン (*.* 等)
//  DstDir       : 移動先フォルダ
//  Recursive    : サブフォルダのファイルも対象か？
// -----------------------------------------------------------------------------
//  Result: (なし)
// -----------------------------------------------------------------------------
class procedure TFileEx.MoveFiles(const SrcDir, SearchPattern, DstDir: string;
  const Recursive: Boolean);
var
  OrgSrcDir, OrgDstDir: string;
  procedure _MoveFiles(const SrcDir, DstDir: string);
  var
    Directory, FileName: string;
    dDstDir, dDstFileName: string;
  begin
    // フォルダを作成
    if not TDirectory.Exists(DstDir) then
      SysUtils.ForceDirectories(DstDir);
    // ファイルを移動
    for FileName in TDirectory.GetFiles(SrcDir, SearchPattern) do
      begin
        dDstFileName := SysUtils.IncludeTrailingPathDelimiter(DstDir) + TPath.GetFileName(FileName);
        if TFile.Exists(dDstFileName) then
          begin
            TFile.SetAttributes(dDstFileName, [TFileAttribute.faNormal]);
            TFile.Delete(dDstFileName);
          end;
        TFile.Move(FileName, dDstFileName );
      end;
    // サブディレクトリを走査
    if Recursive then
      for Directory in TDirectory.GetDirectories(SrcDir) do
        begin
          if SameFileName(Directory, dDstDir) then
            Continue;
          dDstDir := SysUtils.IncludeTrailingPathDelimiter(DstDir) + TPathEx.GetSubPath(Directory);
          _MoveFiles(Directory, dDstDir);
        end;
  end;
begin
  OrgSrcDir := SysUtils.ExcludeTrailingPathDelimiter(SrcDir);
  // ソースフォルダが存在しなければ抜ける
  if not TDirectory.Exists(OrgSrcDir) then
    Exit;
  OrgDstDir := SysUtils.ExcludeTrailingPathDelimiter(DstDir);
  _MoveFiles(OrgSrcDir, OrgDstDir);
end;

{ TPathEx }

// 複数のパスを結合する
// -----------------------------------------------------------------------------
//  Paths: パス
// -----------------------------------------------------------------------------
//  Result: 結合したパス
// -----------------------------------------------------------------------------
class function TPathEx.Combine(const Paths: array of string): string;
var
  i: Integer;
begin
  if High(Paths) = -1 then
    begin
      Result := '';
      Exit;
    end;
  Result := Paths[0];
  for i:=1 to High(Paths) do
    Result := TPath.Combine(Result, Paths[i]);
end;

// 拡張プレフィックス形式のパスを取得する
// -----------------------------------------------------------------------------
//  Path: パス
// -----------------------------------------------------------------------------
//  Result: \\?\ / \\?\UNC\ 形式のパス
// -----------------------------------------------------------------------------
class function TPathEx.GetExtendedPath(const Path: string): string;
var
  dPath: string;
begin
  dPath := Trim(Path);
  if TPath.IsPathRooted(dPath) then
    dPath := ExpandFileName(dPath);
  if Pos(EXTENDED_PREFIX, dPath) = 1 then
    Result := dPath
  else
    case Pos(UNC_PREFIX, dPath) of
      0: Result := EXTENDED_PREFIX + dPath;
      1: Result := EXTENDED_UNC_PREFIX + Copy(dPath, 3, Length(dPath));
    else
      Result := dPath;
    end;
end;

// 指定されたパスの一つ上の階層のパスを得る
// -----------------------------------------------------------------------------
//  Path: パス
// -----------------------------------------------------------------------------
//  Result: 一階層切り捨てたパス
// -----------------------------------------------------------------------------
class function TPathEx.GetParentDir(const Path: string): string;
var
  dPath: string;
  i, Idx: Integer;
begin
  // 空だったりルートパスだったら抜ける
  dPath := SysUtils.ExcludeTrailingPathDelimiter(Path);
  if TPath.IsPathRooted(dPath) then
    dPath := ExpandFileName(dPath);
  if (Trim(dPath) = '') or IsRootPath(dPath) then
    begin
      Result := dPath;
      Exit;
    end;
  // (カレント) サブパスを得る
  Idx := 0;
  for i:=Length(dPath) downto 1 do
    begin
      if dPath[i] = PathDelim then
        begin
          Idx := i;
          Break;
        end;
    end;
  Result := Copy(dPath, 1, Idx - 1);
end;

// 指定されたパスの一つ上の階層のパスを得る
// -----------------------------------------------------------------------------
//  Path: パス
// -----------------------------------------------------------------------------
//  Result: 一階層切り捨てたパス
// -----------------------------------------------------------------------------
class function TPathEx.GetParentPath(const Path: string): string;
begin
  Result := SysUtils.IncludeTrailingPathDelimiter(GetParentDir(Path));
end;

// パスの深さを返す
// -----------------------------------------------------------------------------
//  Path: パス
// -----------------------------------------------------------------------------
//  Result: 0 = 親階層, 1=子階層. 2=孫階層...
// -----------------------------------------------------------------------------
class function TPathEx.GetPathDepth(const Path: string): Integer;
var
  dPath: string;
begin
  dPath := SysUtils.ExcludeTrailingPathDelimiter(GetReducedPath(Path));
  if TPath.IsPathRooted(dPath) then
    dPath := ExpandFileName(dPath);
  Result := PathDelimCount(dPath);
  if TPath.IsUNCRooted(dPath) then
    Dec(Result, 3);
end;

// 拡張プレフィックス形式でないパスを取得する
// -----------------------------------------------------------------------------
//  Path: パス
// -----------------------------------------------------------------------------
//  Result: \\?\ / \\?\UNC\ 形式でないパス
// -----------------------------------------------------------------------------
class function TPathEx.GetReducedPath(const Path: string): string;
var
  dPath: string;
begin
  dPath := Trim(Path);
  if TPath.IsPathRooted(dPath) then
    dPath := ExpandFileName(dPath);
  if Pos(EXTENDED_UNC_PREFIX, dPath) = 1 then
    Result := UNC_PREFIX + StringReplace(dPath, EXTENDED_UNC_PREFIX, '', [rfIgnoreCase])
  else if Pos(EXTENDED_PREFIX, dPath) = 1 then
    Result := StringReplace(dPath, EXTENDED_PREFIX, '', [rfIgnoreCase])
  else
    Result := dPath;
end;

// (カレント) パスのサブパスを得る
// -----------------------------------------------------------------------------
//  Path: パス
// -----------------------------------------------------------------------------
//  Result: サブパス
// -----------------------------------------------------------------------------
class function TPathEx.GetSubPath(const Path: string): string;
var
  dPath: string;
  i, Idx: Integer;
begin
  // 空だったりルートパスだったら抜ける
  dPath := SysUtils.ExcludeTrailingPathDelimiter(Path);
  if TPath.IsPathRooted(dPath) then
    dPath := ExpandFileName(dPath);
  if (Trim(dPath) = '') or IsRootPath(dPath) then
    begin
      Result := '';
      Exit;
    end;
  // (カレント) サブパスを得る
  Idx := 0;
  for i:=Length(dPath) downto 1 do
    begin
      if dPath[i] = PathDelim then
        begin
          Idx := i;
          Break;
        end;
    end;
  Result := Copy(dPath, Idx + 1, Length(dPath));
end;

// 指定されたパスがベースパスであるかを調べる
// -----------------------------------------------------------------------------
//  BasePath: 基準となるパス (ベースパス)
//  Path    : パス
// -----------------------------------------------------------------------------
//  Result: 判定結果
// -----------------------------------------------------------------------------
class function TPathEx.IsBasePath(const BasePath, Path: string): Boolean;
var
  dBasePath, dPath: string;
begin
  dBasePath := BasePath;
  if TPath.IsPathRooted(dBasePath) then
    dBasePath := ExpandFileName(dBasePath);
  dPath     := Path;
  if TPath.IsPathRooted(dPath) then
    dPath := ExpandFileName(dPath);
  Result := SameFileName(SysUtils.ExcludeTrailingPathDelimiter(dBasePath),
                         SysUtils.ExcludeTrailingPathDelimiter(dPath));
end;

// \\?\ 形式の拡張プレフィックスパスかどうかを判定する
// -----------------------------------------------------------------------------
//  Path: パス
// -----------------------------------------------------------------------------
//  Result: 判定結果
// -----------------------------------------------------------------------------
class function TPathEx.IsExtendedPath(const Path: string): Boolean;
begin
  Result := (Pos(EXTENDED_PREFIX, Path) = 1);
end;

// \\?\UNC\ 形式の拡張プレフィックスパスかどうかを判定する
// -----------------------------------------------------------------------------
//  Path: パス
// -----------------------------------------------------------------------------
//  Result: 判定結果
// -----------------------------------------------------------------------------
class function TPathEx.IsExtendedUNCPath(const Path: string): Boolean;
begin
  Result := (Pos(EXTENDED_UNC_PREFIX, UpperCase(Path)) = 1);
end;

// 指定されたパスがネットワーク上を指しているかを調べる
// -----------------------------------------------------------------------------
//  Path: パス
// -----------------------------------------------------------------------------
//  Result: 判定結果
// -----------------------------------------------------------------------------
class function TPathEx.IsRemotePath(const Path: string): Boolean;
var
  dDrive, dPath: string;
begin
  Result := False;
  dPath := TPathEx.GetReducedPath(Path);
  if Length(dPath) = 0 then
    Exit;
  if TPath.IsUNCPath(dPath) then
    begin
      Result := True;
      Exit;
    end;
  dDrive := ExtractFileDrive(dPath);
  if Length(dDrive) = 0 then
    begin
      Result := False;
      Exit;
    end;
  if (Length(dDrive) > 1) and (dDrive[2] <> TPath.VolumeSeparatorChar) then
    begin
      Result := True;
      Exit;
    end;
  {$IFDEF MSWINDOWS}
  Result := (GetDriveType(PChar(dDrive)) = DRIVE_REMOTE);
  {$ELSE}
  Result := False;
  {$ENDIF}
end;

// 指定されたパスがパスのルートであるかを調べる
// -----------------------------------------------------------------------------
//  Path: パス
// -----------------------------------------------------------------------------
//  Result: 判定結果
// -----------------------------------------------------------------------------
class function TPathEx.IsRootPath(const Path: string): Boolean;
var
  dPathRoot: string;
  dPath: string;
begin
  // Windows      : x:
  // Windows (UNC): \\hostname\sharename
  // POSIX        : /
  dPath := GetReducedPath(Path);
  if TPath.IsPathRooted(dPath) then
    dPath := ExpandFileName(dPath);
  dPath := SysUtils.ExcludeTrailingPathDelimiter(dPath);
  dPathRoot := ExtractFileDrive(dPath);
  Result := SameFileName(dPathRoot, dPath);
end;

// 指定されたパスがパス区切り文字で終わっているかを調べる
// -----------------------------------------------------------------------------
//  Path : パス
// -----------------------------------------------------------------------------
//  Result: 判定結果
// -----------------------------------------------------------------------------
class function TPathEx.IsTrailingPathDelimiter(
  const Path: string): Boolean;
begin
  Result := IsPathDelimiter(Path, Length(Path));
end;

// パスにパス区切り文字がいくつ含まれるか調べる
// -----------------------------------------------------------------------------
//  Path: パス
// -----------------------------------------------------------------------------
//  Result: パス区切り文字の数
// -----------------------------------------------------------------------------
class function TPathEx.PathDelimCount(const Path: string): Integer;
var
  s: string;
begin
  Result := 0;
  for s in Path do
    if s = PathDelim then
      Inc(Result);
end;

end.
