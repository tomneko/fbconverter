## fbconverter v2.0.0

Firebird 1.0.3 などの旧版から、Firebird 3.0・4.0・5.0 Embedded へ直接移行できるようになりました。Windows 向けの WIN32／WIN64 ZIP を添付しています。

### 主な変更

- **Firebird 3.0・4.0・5.0 に対応**。両版に 3.0.14、4.0.7、5.0.4 の Embedded ランタイムを同梱しました。3.0 → 4.0／5.0 と 4.0 → 5.0 で、`BOOLEAN`、`INT128`、38桁の `NUMERIC`／`DECIMAL`、`DECFLOAT`、タイムゾーン付き時刻・日時、バイナリ列などの列定義と値を確認しました。
- **Firebird 1.0.3 を変換元として検証**。バージョン判定と Zeos の列メタデータ取得を修正しました。1.5・2.0・2.1・2.5 を含む移行で、日本語と NULL・空・非空の BLOB を照合しました。
- **Firebird 2.1.7 Embedded の EUC を修正**。公式 2.1.7 に CORE-2095 をバックポートした `fbintl.dll` を両版に同梱しました。Firebird 2.1 を変換元として EUC で接続する場合は、サーバー側にも対応する修正が必要です。
- **CLI と MCP を追加**。CLI の `--convert` と stdio MCP の `convert_database` で移行を実行できます。移行先のファイルが既にあれば終了コード `2` を返し、既存ファイルを保持します。
- Delphi 13.1 で Win32／Win64 Release をビルドし、Zeos 8.0-patches と AsyncCalls を使用しています。

### ダウンロードと使い方

- `fbconverter-2.0.0_WIN32.zip` — 32ビット版。Embedded 変換先は 1.5、2.0、2.1、2.5、3.0、4.0、5.0。
- `fbconverter-2.0.0_WIN64.zip` — 64ビット版。Embedded 変換先は 2.1、2.5、3.0、4.0、5.0。

ZIP を展開し、`fbconverter.exe` を実行してください。MCP には Python 3.11 以降と `uv` が必要です。対応ソースとビルド手順は各 ZIP の `Sources/source-2.0.0.tar.gz` に含まれます。[CLI と MCP の使い方](https://github.com/tomneko/fbconverter/blob/v2.0.0/dev_docs/ARCHITECTURE.md#コマンドラインでの変換)

### 配布物の検証

両 ZIP を別々に展開し、2.1 EUC 関連12件と追加データ型を含む移行6件でデータが一致することを確認しました。同梱 MCP からも両版で 1.0.3 → 5.0 を実行し、出力先の衝突時にファイルが変わらないことを確認しました。

| ZIP | SHA-256 |
| --- | --- |
| WIN32 | `ca18f671b6ce2709e69d0a763284a50e9f183e8c03dfa1d2d3d88469d212e02c` |
| WIN64 | `8b04609624ca20c484104b6f387c57e7af3b183f91f1015f3f4b373b3b02e043` |
