![fbconverter — Firebird Database Converter](assets/fbconverter-banner.png)

[English](README.en.md) | Japanese

# fbconverter

fbconverter は、Firebird のデータベースを別バージョンへ移行する Windows ツールです。変換元からスキーマとデータを読み取り、指定したバージョンのデータベースを新規作成します。GUI、CLI、ローカルの stdio MCP から利用できます。

## ダウンロード

[GitHub Releases の最新版](https://github.com/tomneko/fbconverter/releases/latest) から WIN32 または WIN64 の ZIP を取得します。ZIP をフォルダ構成ごと展開し、`fbconverter.exe` を実行してください。

| 配布版 | 同梱の Embedded 変換先 |
| --- | --- |
| WIN32 | 1.5、2.0、2.1、2.5、3.0、4.0、5.0 |
| WIN64 | 2.1、2.5、3.0、4.0、5.0 |

変換元には Firebird 1.0、1.5、2.0、2.1、2.5、3.0、4.0、5.0 系のサーバーを指定できます。1.0 系の動作検証には Firebird 1.0.3 を使用しました。

公式の Firebird 2.1.7 サーバー版には、`EUCJ_0208` の不具合を直す [CORE-2095](https://github.com/FirebirdSQL/firebird/issues/2529) の修正が入っていません。Firebird 2.1 への移行でこの不具合に遭遇した場合は、WIN32・WIN64 両 ZIP に同梱の **修正版 Firebird 2.1.7 Embedded** を変換先に選んでください。変換元の Firebird 2.1 サーバーでこの不具合が発生する場合は、変換元サーバー側にも修正が必要です。

## CLI

PowerShell から 2.5 のサーバーに接続し、5.0 Embedded のデータベースを作成する例です。

```powershell
.\fbconverter.exe --convert --source C:\data\source.fdb --destination C:\data\migrated.fdb `
  --source-version 2.5 --destination-version 5.0 --source-host localhost `
  --source-charset SJIS_0208 --destination-charset UTF8 | Out-Host
$LASTEXITCODE
```

移行先のファイルが既にあると終了コード `2` を返し、既存ファイルを保持します。パスワードは環境変数 `FBCONVERTER_SOURCE_PASSWORD` で渡せます。[CLI の引数とログ](dev_docs/ARCHITECTURE.md#コマンドラインでの変換)

## MCP

同梱の `fbconverter_mcp.py` を stdio サーバーとして登録すると、MCP クライアントから `convert_database` を呼べます。Python 3.11 以降と `uv` が必要です。スクリプトと同じフォルダの `fbconverter.exe` を使います。[設定例とツール引数](dev_docs/ARCHITECTURE.md#mcp-から実行する)

## 開発

Delphi 13.1 で Win32／Win64 Release をビルドできます。依存は Zeos 8.0-patches r8528（Firebird 1.0.3 対応パッチ適用）と AsyncCalls です。[ビルド手順と検証記録](dev_docs/ARCHITECTURE.md) · [日本語 Wiki](https://github.com/tomneko/fbconverter/wiki/Japanese)

fbconverter は [MIT ライセンス](LICENSE) です。依存ライブラリのライセンス、fbconverter と修正版 DLL の対応ソース、ビルド手順は配布 ZIP に含まれます。
