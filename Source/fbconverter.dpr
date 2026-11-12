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
  System.SysUtils,
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

  CommandLineMode := ParamCount > 0;
  if CommandLineMode and (ParamStr(1) <> '--convert') then
    begin
      if (ParamCount <> 1) or (ParamStr(1) <> '--help') then
        ExitCode := 2;
      WriteCommandLine(
        'Usage: fbconverter --convert --source FILE --destination FILE ' +
        '--source-version VERSION --destination-version VERSION' + sLineBreak +
        'Options: --source-host HOST | --source-local, --source-port PORT, --source-user USER,' + sLineBreak +
        '         --source-charset CHARSET, --destination-charset CHARSET, --work-dir DIRECTORY, --single-thread' + sLineBreak +
        'Set FBCONVERTER_SOURCE_PASSWORD to supply the source password.', ExitCode <> 0);
      Exit;
    end;

  // 二重起動防止ロジック
  hMutex := CreateMutex(nil, False, MUTEX_APPLI_ID);
  if (hMutex = 0) or (GetLastError = ERROR_ALREADY_EXISTS) then
    begin
      if hMutex <> 0 then
        CloseHandle(hMutex);
      if CommandLineMode then
        begin
          WriteCommandLine('Could not start a conversion. Check for another fbconverter instance.', True);
          ExitCode := 1;
        end;
      Exit;
    end;
  try
    Application.Initialize;
    Application.MainFormOnTaskbar := True;
    if CommandLineMode then
      try
        Application.CreateForm(TfrmMain, frmMain);
        Application.CreateForm(TfrmLogViewer, frmLogViewer);
        frmMain.RunCommandLine;
      except
        on E: EArgumentException do
          begin
            WriteCommandLine(E.Message, True);
            ExitCode := 2;
          end;
        on E: Exception do
          begin
            WriteCommandLine(E.Message, True);
            ExitCode := 1;
          end;
      end
    else
      begin
        Application.CreateForm(TfrmMain, frmMain);
        Application.CreateForm(TfrmLogViewer, frmLogViewer);
        Application.Run;
      end;
  finally
    CloseHandle(hMutex);
  end;
end.
