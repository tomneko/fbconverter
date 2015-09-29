{*******************************************************}
{                                                       }
{              Firebird Database Converter              }
{                                                       }
{*******************************************************}

// -----------------------------------------------------------------------------
// Note:
//   - Firebird Database Converter Project File
// -----------------------------------------------------------------------------
program fbconverter;

uses
  {$IFDEF FullDebugMode}
  FastMM4,
  {$ENDIF}
  Vcl.Forms,
  WinAPI.Windows,
  frmuMain in 'frmuMain.pas' {frmMain},
  uZeosFBUtils in 'uZeosFBUtils.pas',
  Vcl.Themes,
  Vcl.Styles,
  uFBExtract in 'uFBExtract.pas',
  frmuLogViewer in 'frmuLogViewer.pas' {frmLogViewer},
  uSQLBuilder in 'uSQLBuilder.pas',
  uFBConvConsts in 'uFBConvConsts.pas',
  uIOUtilsEx in 'uIOUtilsEx.pas',
  uFBRawAccess in 'uFBRawAccess.pas';

{$R *.res}

const
  MUTEX_APPLI_ID = 'Global\FIREBIRD_CONVERTER';
var
  hMutex:THandle;
begin
  // デバッグ実行時にはアプリケーション終了時にメモリリークをレポート
  {$WARN SYMBOL_PLATFORM OFF}
  if DebugHook <> 0 then
    ReportMemoryLeaksOnShutdown := True;
  {$WARN SYMBOL_PLATFORM ON}

  // 二重起動防止ロジック
  hMutex := OpenMutex(MUTEX_ALL_ACCESS, False, MUTEX_APPLI_ID);
  if (hMutex <> 0) then
    begin
      CloseHandle(hMutex);
      Exit;
    end;
  hMutex := CreateMutex(nil, False, MUTEX_APPLI_ID);
  try
    Application.Initialize;
    Application.MainFormOnTaskbar := True;
    Application.CreateForm(TfrmMain, frmMain);
    Application.CreateForm(TfrmLogViewer, frmLogViewer);
    Application.Run;
  finally
    ReleaseMutex(hMutex);
  end;
end.
