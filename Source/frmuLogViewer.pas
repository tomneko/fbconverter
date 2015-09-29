{*******************************************************}
{                                                       }
{              Firebird Database Converter              }
{                                                       }
{*******************************************************}

// -----------------------------------------------------------------------------
// Note:
//   - Log Viewer Form
// -----------------------------------------------------------------------------
unit frmuLogViewer;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls;

type
  TfrmLogViewer = class(TForm)
    mmMessage: TMemo;
    mmError: TMemo;
    Splitter1: TSplitter;
    procedure mm_KeyPress(Sender: TObject; var Key: Char);
  private
    { Private 宣言 }
  public
    { Public 宣言 }
    procedure AddError(s: string; StayLastLine: Boolean = False);
    procedure AddMessage(s: string; StayLastLine: Boolean = False);
    procedure DeleteError;
    procedure DeleteMessage;
    procedure SaveMessage(const FileName: string);
    procedure SaveError(const FileName: string);
    procedure InitLog;
  end;

var
  frmLogViewer: TfrmLogViewer;

implementation

{$R *.dfm}

procedure TfrmLogViewer.AddError(s: string; StayLastLine: Boolean);
// エラーログを追加
begin
  if StayLastLine and (mmError.Lines.Count > 0) then
    mmError.Lines[mmError.Lines.Count-1] := s
  else
    mmError.Lines.Add(s);
  mmError.Perform(EM_SCROLLCARET, 0, 0);
  Application.ProcessMessages;
end;

procedure TfrmLogViewer.DeleteError;
// メッセージログの最下行を削除
begin
  if mmError.Lines.Count > 0 then
    mmError.Lines.Delete(mmError.Lines.Count-1);
end;

procedure TfrmLogViewer.AddMessage(s: string; StayLastLine: Boolean);
// メッセージログを追加
begin
  if StayLastLine and (mmMessage.Lines.Count > 0) then
    mmMessage.Lines[mmMessage.Lines.Count-1] := s
  else
    mmMessage.Lines.Add(s);
  mmError.Perform(EM_SCROLLCARET, 0, 0);
  Application.ProcessMessages;
end;

procedure TfrmLogViewer.DeleteMessage;
// メッセージログの最下行を削除
begin
  if mmMessage.Lines.Count > 0 then
    mmMessage.Lines.Delete(mmMessage.Lines.Count-1);
end;

procedure TfrmLogViewer.InitLog;
// ログのクリア
begin
  mmMessage.Lines.Clear;
  mmError.Lines.Clear;
end;

procedure TfrmLogViewer.mm_KeyPress(Sender: TObject; var Key: Char);
// 全選択
begin
  if Key = ^A then
    (Sender as TMemo).SelectAll;
end;

procedure TfrmLogViewer.SaveError(const FileName: string);
// エラーログを保存
begin
  mmError.Lines.SaveToFile(FileName, TEncoding.UTF8);
end;

procedure TfrmLogViewer.SaveMessage(const FileName: string);
// メッセージログを保存
begin
  mmMessage.Lines.SaveToFile(FileName, TEncoding.UTF8);
end;

end.
