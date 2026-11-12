# fbconverter のアーキテクチャ

対象はコミット `7bbd479` 時点のソースと同梱配布物。fbconverter は Delphi XE4 で開発された Windows VCL アプリケーション。変換元の Firebird に接続し、システムカタログから生成した DDL で変換先データベースを新規作成し、行データとジェネレータ値を移す。これにより、変換先の Firebird エンジンが自身の ODS（On-Disk Structure）でデータベースを作成する。

## ソースの構成

| ファイル | 責務 |
| --- | --- |
| [fbconverter.dpr](../Source/fbconverter.dpr) / [fbconverter.dproj](../Source/fbconverter.dproj) | VCL の起動、名前付きミューテックスによる多重起動の抑止、フォーム生成、Win32/Win64 のビルド設定 |
| [frmuMain.pas](../Source/frmuMain.pas) / [frmuMain.dfm](../Source/frmuMain.dfm) | 入出力の設定と検証、Firebird DLL の選択、変換順序、行データの取り込み、ログ出力 |
| [uFBExtract.pas](../Source/uFBExtract.pas) | `TFBExtractor` が `RDB$` システムカタログを参照して DDL を生成。文字セット情報、行の分割、Verbose 用の INSERT と BLOB ファイル生成も担当 |
| [uZeosFBUtils.pas](../Source/uZeosFBUtils.pas) | `TZeosFB` が Zeos の接続・クエリ・トランザクションをラップし、DB 作成と DB／サーバー情報取得を担当 |
| [uFBRawAccess.pas](../Source/uFBRawAccess.pas) | Firebird のネイティブ API 型と定数、接続パラメータブロック、接続先文字列の組み立て |
| [uSQLBuilder.pas](../Source/uSQLBuilder.pas) | SELECT／INSERT／DELETE／UPDATE の組み立て用レコード |
| [frmuLogViewer.pas](../Source/frmuLogViewer.pas) / [frmuLogViewer.dfm](../Source/frmuLogViewer.dfm) | メタデータとメッセージの表示・ファイル保存 |
| [uIOUtilsEx.pas](../Source/uIOUtilsEx.pas) / [uFBConvConsts.pas](../Source/uFBConvConsts.pas) | パス・ディレクトリ操作、メッセージ定数 |

## 接続と対応バージョン

- `TFBVersion` の列挙値は 1.0、1.5、2.0、2.1、2.5、3.0、4.0、5.0（および不明）。画面の候補は `FormCreate` が起動時に検出した DLL によって決まる。列挙値だけでは実際に変換できる組み合わせは判断できない。
- 変換元は `Client/fbclient<ドットなしの版>.dll` を選択する。変換先は同じクライアント DLL を用いる接続、または Embedded によるローカル接続を選択する。DLL の検出とロードは実行ファイルのディレクトリが基準。2.5 までは `Embedded/<版>/fbembed.dll`、3.0 以降は下記の公式 ZIP の構成を使う。
- DB 情報の取得では `TZeosFB.GetDatabaseInfo` がネイティブ API を呼び、ODS、ページサイズ、SQL dialect などを取得する。文字セットの自動検出には `TFBExtractor.Init` を使う。
- Embedded 変換では、選択したエンジンの `firebird.conf` を実行ファイルの隣にコピーし、`RootDirectory` と `RelaxedAliasChecking` を設定する。画面の Strict check は後者の有効・無効を反転させる。

## 修正版 ZeosLib を使う理由

[日本語 Wiki の開発経緯](https://github.com/tomneko/fbconverter/wiki/Japanese) では、Zeos の標準メタデータ取得で変換用 DDL に必要な情報が揃わなかったため、`uFBExtract.pas` を実装したと説明している。現行ソースでは `TFBExtractor` が Zeos のクエリを使って `RDB$` カタログを読み、DDL を組み立てる。

Wiki によると、当時の Zeos 7.1.3 以降はメタデータ取得の SQL で `TRIM()` を使い、同関数を組み込みで持たない Firebird 1.0／1.5 でエラーになった。一方、7.1.2 には BLOB サブタイプ `-1` のエラー、マルチバイト環境でのシステムカタログ値の切り詰め、`SJIS_0208` を EUC として扱う問題があった。このため、XE4 時代は 7.1.2 をベースにした修正版を同梱していた。Florence 向けの依存関係は下記のとおり。Zeos 8 での Firebird 1.0／1.5 の動作は未検証。

## 文字セット変換時の注意点

- **`NONE` のデータは実際の書き込み時の文字セットを確認する。** Auto detect は `TFBExtractor.Init` が `RDB$DATABASE` と `RDB$CHARACTER_SETS` から取得した DB の既定文字セットを画面に反映する。格納済みのバイト列から文字コードを推定する処理はない。Wiki の `SJIS_0208` は変換元画面の初期値（INI の `DEFAULT_CHARSET` で変更可能）としても使われるが、Auto detect はカタログに `NONE` があればその値を選ぶ。別の文字セットで書き込んだデータは変換元で明示的に指定する。
- **変換先 DB の既定値と個別の列定義を区別する。** 画面の変換先文字セットは選択肢がある場合に 2.0 以上で `UTF8`、それ以前で `UNICODE_FSS` を初期選択する。`BuildFieldString` と `BuildArgumentString` は変換元 DB の既定文字セットと異なる文字セットを持つ文字型に `CHARACTER SET` を明記するため、変換先の既定文字セットを指定しても全列の定義が一律に置き換わるわけではない。
- **生成された `CHAR`／`VARCHAR` の長さを確認する。** Wiki は 1 文字あたりの最大バイト数（BPC）と 32765 バイトの上限を説明している。実装では列・引数の文字セット ID と変換先文字セット名を `CharSets` 定数表で引く。変換先の BPC が大きく、元の文字数 × 変換先 BPC が 32765 を超えると、生成する文字数を `元の文字数 div 変換先 BPC` にする。この式は `32765 div 変換先 BPC` より小さな定義を生成し得るため、長い列では生成 DDL と実データの最大長を照合する。
- **接続時の文字セット設定も変換に影響する。** `TZeosFB` は Firebird の接続パラメータ `lc_ctype`、Zeos の `ClientCodepage`、Delphi 側の `ControlsCodePage`（UTF-16）を個別に設定する。変換元または変換先の接続文字セットが `CP943C` の場合、`ProcessConvert` は該当接続の `ClientCharset` に `UTF8` を設定する。Zeos の更新や設定変更では、日本語の行データとマルチバイトの識別子の両方が確認対象になる。

## 変換の実行順序

入口は `TfrmMain.AL_ConvertExecute` → `ProcessConvert`。

1. `CheckError` で変換元に接続できることを確認する。ローカルパスについては存在確認などを行い、変換先の上書きを確認する。
2. 作業ディレクトリを作り直し、変換元の接続・文字セット情報と DB 情報を取得する。Metadata only 以外では変換先 DB を新規作成する。新規 DB の SQL dialect は既定値の 3。ページサイズは変換元を基準にし、変換先が 2.5 以上の場合は最低 4096 バイトにする。
3. ドメイン、テーブル、プロシージャの仮定義、ビュー、ジェネレータ、外部関数、例外、BLOB フィルタ、ロール、プロシージャ本体の順に生成する。
4. テーブルの行を取り込み、ジェネレータの現在値を設定する。
5. 外部キー、インデックス、テーブルのトリガ、権限を生成・適用し、接続を閉じてログを保存する。

外部キーとテーブルのトリガの適用は行データの取り込み後。DDL の `EZSQLException` はログに記録して処理を続ける。Fast Mode も分割単位でコミットするため、失敗時には作成途中の DB が残り得る。また、`ProcessConvert` 冒頭の既存ファイル削除はローカルホスト判定を再利用せず、指定パスに対して `TFile.Exists` を実行する。

## データの取り込みと出力

| モード | 実装上の動作 |
| --- | --- |
| 通常（Fast Mode） | `ImportTable` → `ImportData`。分割ごとに変換元・変換先へ接続し、`FIRST … SKIP …` で取得した行を型別に INSERT パラメータへ設定してコミットする。NULL はパラメータをクリアし、BLOB はストリームを介して渡す。 |
| Verbose mode | `TFBExtractor.BuildTableData` でログ出力用と実行用の INSERT を生成する。実行用では作業ディレクトリ内のテキスト・BLOB ファイルから値をパラメータに設定する。 |
| Metadata only | 変換先への DB 作成・SQL 実行を省略する。テーブルの行も読み取り、`metadata.sql` に INSERT を出力する。 |
| Create empty database | 行データの取り込みとジェネレータ値の設定を省略し、スキーマと権限を作成する。 |

並列取り込みは Fast Mode で `USE_MULTI_THREAD` が有効、かつ変換先がクライアント／サーバー接続または Embedded 2.5 以上の場合に使用する。`AsyncCalls` のワーカーごとに接続を持ち、テーブルごとに全ワーカーの終了を待つ。取り込み時の分割クエリに `ORDER BY` は指定されていない。

作業場所は `WORK_DIR` の下の `Firebird Converter/`。起動時（既存の場合）と変換開始時にこのサブディレクトリを削除して `BLOB_DATA/` を作り直す。正常にログ保存まで進むと `metadata.ddl` と `metadata.err` を出力し、Verbose mode または Metadata only では `metadata.sql` も作成する（Create empty database を除く）。`metadata.ddl` には接続パス・ユーザー名・パスワードを含む CREATE DATABASE 文がコメントとして出力され得る。SQL と BLOB ファイルには変換元のデータが含まれ得る。

## ビルドと配布

- `Source/fbconverter.dproj` は Win32／Win64 と Debug／Release／FullDebug の構成を持つ。プロジェクトの既定は Win32 Release。
- Florence 向けのプロジェクトは公式 Zeos `8.0-patches` の r8528 と AsyncCalls の `dd26254` をソースからビルドする。Zeos のプロトコルは `firebird`、接続と DB 作成の API は `FirebirdAPI=legacy` を指定する。[同梱の修正版 ZeosLib 7.1.2a](../Library/ZeosLib-7.1.2a.zip) は旧構成の参照用。FullDebug では FastMM4 も参照する。
- 設定ファイルは実行ファイルと同名の UTF-8 INI（通常は `fbconverter.ini`）。`[Globals]` の主な項目は `DEFAULT_CHARSET`（SJIS_0208）、`MAX_ERROR_LINE`（5）、`POST_DATA_SPLIT_SIZE`（1000、最小 100）、`USE_MULTI_THREAD`（有効）、`NUMBER_OF_IMPORT_THREAD`（4、2 から分割サイズの半分まで）、`WORK_DIR`（既定は `GetHomePath`）。`DEFAULT_CHARSET` は起動時の変換元文字セットにも使われる。
- `Release/` には既存の Win32／Win64 配布 ZIP がある。1.0.5.19 の Win32 ZIP では実行ファイル、INI、Client 1.0〜2.5、Embedded 1.5〜2.5 の DLL を確認した。同版の Win64 ZIP は Windows の ZIP API でエントリが 0 件となり、`tar` でも破損と報告された。
- 実行検証した変換の組み合わせは下記の単一スレッド試験を参照。

### Delphi 13.1 向け Win32 Release

リポジトリのルートで依存ソースを取得する。

```bat
svn export -r 8528 https://svn.code.sf.net/p/zeoslib/code-0/branches/8.0-patches Library\zeos-8.0-patches-r8528
git apply --directory=Library/zeos-8.0-patches-r8528 dev_docs/zeos-r8528-firebird-1.0.3.patch
git clone https://github.com/ahausladen/AsyncCalls.git Library\asynccalls-dd26254
git -C Library\asynccalls-dd26254 checkout --detach dd26254
```

RAD Studio Command Prompt からビルドする。既存のバージョン番号を維持するため、下のコマンドでは自動インクリメントを無効にしている。

```bat
msbuild Source\fbconverter.dproj /t:Build /p:Config=Release /p:Platform=Win32 /p:VerInfo_AutoIncVersion=false
```

出力先は `Source\Win32\Release\fbconverter.exe`。公式 Zeos r8528 には Florence の `ftLargeUint` 対応がある。検証環境の IDE はバージョン情報に Delphi 13 Update 1（13.1）と表示され、`bds.exe` と同じインストール先の `dcc32.exe` はともに `37.0.58802.5665`。修正後の `.dproj` を使い、上記の MSBuild コマンドで Win32 Release の実行ファイル生成を確認した。

## コマンドラインでの変換

`--convert` は変換元の Firebird に接続し、指定したバージョンの Embedded DB をローカルに新規作成する。変換元と変換先に必要な DLL は、実行ファイルに隣接する `Client/` と `Embedded/` に配置する。引数なしの起動では従来の画面を開く。

```powershell
.\fbconverter.exe --convert --source C:\data\source.fdb --destination C:\data\converted.fdb `
  --source-version 1.5 --destination-version 2.5 --source-host localhost `
  --source-charset SJIS_0208 --destination-charset UTF8 --single-thread | Out-Host
$LASTEXITCODE
```

PowerShell ではパイプを付けて実行すると、GUI サブシステムの実行ファイルでも終了を待って結果を取得できる。引数一覧は `--help` で表示する。接続先の既定値は `localhost:3050`。必要に応じて `--source-port`、`--source-user` と環境変数 `FBCONVERTER_SOURCE_PASSWORD` を指定する。空のホスト名でローカル接続するには `--source-local` を使う。

終了コードは成功が `0`、変換エラーが `1`、引数エラーや変換先ファイルが既に存在する場合が `2`。確認ダイアログが必要な状況ではエラーを出力して終了する。ログは一時ディレクトリまたは `--work-dir` の下に毎回新しいディレクトリを作って保存する。出力された `Logs:` のパスから `metadata.ddl` と `metadata.err` を参照できる。

2026-11-11 の単一スレッド試験では、Firebird 1.5.6 サーバー（ループバック）から Firebird 2.5.3 Embedded へ `SJIS_0208` → `UTF8` で変換した。変換元と変換先で日本語 2 件と NULL 文字列が一致し、バイナリ BLOB の NULL、0 バイトの非 NULL、4 バイト（`00 7F 80 FF`）も区別して一致した。

同じ構成で固定のテストデータ 3,257 行を単一スレッドと並列で変換した。並列は `USE_MULTI_THREAD=1`、`NUMBER_OF_IMPORT_THREAD=4`、`POST_DATA_SPLIT_SIZE=1000` とし、250 行ずつの 13 ブロックと末尾 7 行の計 14 ブロックの完了をログで確認した。変換元と両方の変換先を ID 順に読み戻した結果は、ID 1〜3,257 が欠落・重複なく一致した。日本語文字列、NULL 文字列、および BLOB の NULL 1,086 件、空 1,086 件、非空 1,085 件も全件一致した。

Win64 Release も Delphi 13.1 の MSBuild でビルドした。上記コマンドの `/p:Platform` を `Win64` に変更する。検証には 1.0.4.18 の Win64 ZIP に含まれる x64 の Firebird 2.5.4 クライアントと Embedded DLL を使用した。この ZIP に Firebird 1.5 の x64 クライアントがないため、検証用ディレクトリの `Client/fbclient15.dll` に 2.5.4 クライアントのコピーを置き、変換元は Win32 試験と同じ Firebird 1.5.6 サーバーに接続した。出力エンジンは Win64 が 2.5.4 Embedded、Win32 が 2.5.3 Embedded。Win64 の単一・並列とも終了コード 0 で、並列は 4 ワーカー設定で 14 ブロックの完了を確認した。変換元と Win32／Win64 の 4 出力を ID 順に読み戻し、全 3,257 行の ID、文字列、BLOB を比較した結果はすべて一致した。

### MCP から実行する

`mcp/fbconverter_mcp.py` はローカルの stdio MCP サーバー。公式 MCP Python SDK を使用し、`convert_database` から既存の `fbconverter.exe --convert` を呼ぶ。Python 3.11 以降と `uv` が必要。初回起動時に `uv` が依存パッケージを取得する。

MCP クライアントで stdio サーバーを登録する。以下は `mcpServers` 形式の設定例。`FBCONVERTER_EXE` は Client／Embedded ランタイムを配置した Win32 または Win64 版を指定する。設定を省略すると、Python スクリプトと同じフォルダの実行ファイルを探す。

```json
{
  "mcpServers": {
    "fbconverter": {
      "command": "uv",
      "args": ["run", "--script", "D:/data/fbconverter/mcp/fbconverter_mcp.py"],
      "env": {
        "FBCONVERTER_EXE": "D:/data/fbconverter/Source/Win32/Release/fbconverter.exe"
      }
    }
  }
}
```

`convert_database` の呼び出し例。

```json
{
  "source": "C:/data/source.fdb",
  "destination": "C:/data/migrated.fdb",
  "source_version": "2.5",
  "destination_version": "5.0",
  "source_host": "localhost",
  "source_charset": "SJIS_0208",
  "destination_charset": "UTF8",
  "single_thread": true
}
```

`source` は接続先 Firebird サーバーから見たパス。ホストの既定値は `localhost` で、直接ローカル接続する場合は `source_local: true` を指定する。`source_port`、`source_user`、`work_dir` も指定できる。相対の出力先や作業フォルダは MCP サーバーのカレントディレクトリが基準。パスワードが必要なときは MCP サーバーの環境変数 `FBCONVERTER_SOURCE_PASSWORD` で渡す。

ツールは CLI の完了を待ち、`destination`、`exit_code`、`log_dir`、`stdout`、`stderr` を返す。CLI の終了コードが 0 以外なら MCP のツールエラーとして返す。移行先が既にある場合は終了コード 2 で、既存ファイルは保持される。CLI を起動できなかった場合の `exit_code` は `null`。


公式 MCP クライアントから、Win32 ではプロトコル 2026-07-28、Win64 では 2025-11-25 でツールを呼び出した。Firebird 1.0.3 → 5.0 Embedded の日本語と BLOB を読み戻し、同じ出力先で再実行すると終了コード 2 が返り、出力ファイルの SHA-256 が変わらないことも確認した。

### 日本語文字セットの交差試験

Win32 Release の単一スレッド CLI を使用した。変換元はループバックで起動した Firebird サーバー、変換先は配布 ZIP 同梱の Embedded。テスト DB は日本語 2 件と NULL 文字列、NULL・空・非空のバイナリ BLOB を含む 3 行。ID 順に読み戻して文字列を UTF-8 バイト列で比較し、BLOB の NULL 判定と内容も比較した。`NONE` は Shift_JIS のバイト列を格納したものとし、変換元は `SJIS_0208` を指定、変換先の読み戻しも Shift_JIS として行った。`CP943C` の読み戻しは UTF-8 接続を使用した。

候補は `SJIS_0208`、`EUCJ_0208`、`NONE`、`UNICODE_FSS`、`UTF8`、`CP943C`。`UTF8` は 2.0 以降、`CP943C` は 2.1 以降を対象とした。Firebird 1.0 の Embedded は配布されておらず、現行 CLI は変換先に Embedded を使う。この初回の交差試験には 1.0 を含めていない。後述の 1.0.3 検証では、インストール済みサーバーを変換元に使用した。 [Firebird の文字セット一覧](https://www.firebirdsql.org/refdocs/langrefupd25-charsets.html)

変換元は 1.5.6、2.0.7、2.1.7、2.5.9。変換先は 1.5.6、2.0.7、2.1.6、2.5.3 Embedded。2.1 の `EUCJ_0208` はテスト DB の初期化で SQL エラー `-802` が出たため、変換元から除外した。2.0 のクライアント DLL は同梱の VC7.1 ランタイムを検証フォルダへ置くと読み込めた。

| 変換元 | 実行 | 一致 | 失敗 |
| --- | ---: | ---: | ---: |
| 1.5 | 84 | 64 | 20 |
| 2.0 | 105 | 80 | 25 |
| 2.1 | 105 | 80 | 25 |
| 2.5 | 126 | 90 | 36 |
| 合計 | 420 | 314 | 106 |

失敗は 2.1.6 Embedded の `EUCJ_0208` 宛ての文字セット未定義エラーが 20 件、6 秒の試験タイムアウトが 86 件。タイムアウトのうち 60 件は 1.5／2.0／2.1 の `NONE` 宛て、18 件は 2.5 の `UNICODE_FSS` 宛て。2.5 の `NONE` 宛ては試した 20 件がすべて一致した。タイムアウトは制限時間内に終了しなかったことを示し、無限待ちかどうかは未判定。ケース別の結果は検証フォルダの `charset-matrix/matrix-results.csv` に保存した。

#### タイムアウトと EUC エラーの再調査

- 1.5.6／2.0.7／2.1.6 Embedded の `NONE` 宛ては、30 秒でも完了せず CPU を使い続けた。DB ファイルは作成されたが、スキーマの処理には進まなかった。Zeos 8 の旧 API 経路には、DB の既定値が `NONE` で接続コードページが異なると再接続する分岐がある。`TZeosFB.CreateDatabase` で作成接続の `ClientCodepage` を `NONE` に明示した。
- 2.5.3 Embedded の `UNICODE_FSS` 宛ては、30 秒試験で INSERT がテーブル未定義エラーになり、エラー出力後も終了しなかった。Zeos の SQL 作成経路が文字セット名の解析でアンダースコアを認めず、`UNICODE_FSS` を `UNICODE` として扱う箇所を確認した。2.5 以降の Embedded でこの文字セットを作る場合は、Zeos 経由で `isc_create_database` を使い、文字セット、ページサイズ、dialect を DB パラメータブロックに指定する。作成後の DB を `isql` で開き、`UNICODE_FSS`、4096 バイト、dialect 3 を確認した。
- 2.1.6 Embedded を `isql` で直接試すと、`EUCJ_0208` はシステムカタログに存在するが、同文字セットを指定した接続は `-924`。接続を `NONE` または `UTF8` に変更しても、既定の EUC を使う `VARCHAR` の作成は `-802`。2.1.7 サーバーの `isql` でもテスト DB 初期化時に `-802` が出た。Firebird 2.1 の文字セット処理まで切り分けられたが、エンジン内部の原因は未特定。
  - [CORE-1524](https://github.com/FirebirdSQL/firebird/issues/1938) は Windows の Firebird 2.1 Beta 2 で `EUCJ_0208` の DB 作成が同じエラー文と ISC コード `335544325` を返した報告。チケットは再現不可で閉じられ、非英語版 Windows の VC8 CRT 依存が候補に挙がっている。今回の 2.1.6 Embedded の `isql` は `MSVCR80.dll`、同梱の `fbintl.dll` と ICU 3.0 を実際にロードしていた。同一 DB で `SJIS_0208` 列の作成は成功し、`EUCJ_0208` 列の作成は `-802`。CRT の単純な未ロードでは説明できない。
  - [CORE-2095](https://github.com/FirebirdSQL/firebird/issues/2529) は 2.1.0／2.1.1 などで記録された EUC → Unicode 変換の不具合。2.5 Beta 1 向けの [修正差分](https://github.com/FirebirdSQL/firebird/commit/795e42647e7e100914aa834554b86a3a29794e2b) は出力バッファ長の見積もりと不完全なマルチバイト入力の判定を修正する。今回の接続・DDL エラーは下記の A/B 試験で見積もりの不具合に絞り込んだ。
- 2.5 EUC → 2.1 `NONE` のケースでは、生成された列定義に `CHARACTER SET EUCJ_0208` が残る。この DDL は既定文字セットが `NONE` の DB でも `isql` で `-802` になる。CLI は最初の SQL エラーとその時点のログを保存し、変換元接続を明示的に切断するようにした。このケースは最初の `CREATE TABLE` エラーを表示して終了コード 1 を返した。

最終 Win32 Release で対象の 80 通りを再実行した。旧版 `NONE` 宛ては 60 通り中 59 通り、2.5 `UNICODE_FSS` 宛ては 20 通りすべてで日本語・NULL・空と非空の BLOB が一致した。残りは上記の 2.1 EUC 列作成エラーで、80 通りにタイムアウトはなかった。Win64 Release もビルドでき、作成経路の修正を含む Win64 版で 1.5.6 `SJIS_0208` → 2.5.4 Embedded `UNICODE_FSS` の 3 行を読み戻して照合した。

#### Firebird 2.1 の EUC 初期化と CORE-2095

2.1.7 の `intlnames.h` は `EUCJ_0208` を `CS_euc_j` と `JIS230_init` に登録する。2.5.3 の `ld.cpp` にも同じ対応がある。2.1.6 の実行時トレースでは文字セットと照合順序の lookup は両方とも成功した。 [2.1.7 の登録表](https://github.com/FirebirdSQL/firebird/blob/R2_1_7/src/jrd/intlnames.h#L96-L105) · [2.5.3 の登録表](https://github.com/FirebirdSQL/firebird/blob/R2_5_3/src/intl/ld.cpp#L271-L287)

その後、`CharSetContainer::lookupCollation` が `TextType` を生成する。コンストラクタは SQL のパターン文字を正規化し、`canonical` が `convertLength` で見積もった長さを UTF-16 変換の出力容量に渡す。2.1.7 の `CVJIS_eucj_to_unicode` は長さ照会で入力バイト数をそのまま返すため、ASCII 1 バイトに対して UTF-16 の 2 バイトを出力できず `CS_TRUNCATION_ERROR` になる。DDL では `-802` が返る。接続時は `INTL_defined_type` が初期化例外を捕捉して false を返し、`jrd.cpp` が `-924` の文字セット未定義エラーとして報告する。 [TextType の初期化](https://github.com/FirebirdSQL/firebird/blob/R2_1_7/src/jrd/TextType.cpp#L105-L111) · [UTF-16 への変換](https://github.com/FirebirdSQL/firebird/blob/R2_1_7/src/jrd/TextType.cpp#L329-L349) · [接続時のエラー処理](https://github.com/FirebirdSQL/firebird/blob/R2_1_7/src/jrd/jrd.cpp#L5393-L5410)

最小再現は既定文字セット `NONE` の DB で `CREATE TABLE EUC_TEST (V VARCHAR(20) CHARACTER SET EUCJ_0208)` を実行し、別の `isql` から `SET NAMES EUCJ_0208` で接続する。通常の 2.1.6 Embedded では両方失敗し、通常の 2.5.3 Embedded では両方成功した。さらに、隔離した 2.1.6 のテスト用 DLL で元の `fbintl.dll` に処理を委譲し、EUC の長さ照会の戻り値だけを CORE-2095 と同じ `sizeof(USHORT) * src_len` に切り替えた。切替前は `input=1, capacity=1, error=1` で両方失敗し、切替後は `input=1, capacity=2, result=2, error=0` で両方成功した。検証用 DLL は配布物に含めていない。


#### CORE-2095 の Win32 バックポート

Firebird の公式 `R2_1_7` ソースに [バックポート差分](CORE-2095-firebird-2.1.patch) を適用した。`CVJIS_eucj_to_unicode` の出力長の見積もりを UTF-16 のバイト数にし、入力末尾に EUC の先行バイトだけが残った場合は次のバイトを参照する前に `CS_BAD_INPUT` を返す。 [公式ソースタグ](https://github.com/FirebirdSQL/firebird/releases/tag/R2_1_7)

VS 2022 Build Tools の MSVC 14.44（x86、最適化有効、静的 CRT）で、公式 `intl.vcproj` の C++ ソース一式をコンパイルした。ビルド用の `autoconfig.h` で UCRT と衝突する `snprintf`／`vsnprintf` のマクロを外し、旧 ICU をロードする Win32 のコードで VC8 CRT の DLL 名を指定した。公式 2.1.7 の `icuuc30.dll` のエクスポートからインポートライブラリを生成してリンクした。生成物は PE x86 で、公式 DLL と同じ4個の `LD_*` エクスポートを確認した。ビルド手順とコンパイルログは検証フォルダ `charset-matrix/backport-21/` の `build.cmd`、`compile.rsp`、`compile-modern.log` にある。`fbintl.dll` の SHA-256 は `D35F7C36A2C91903328CFD8A0D5F571593DCFB27007187AE7576F4EB59DF3548`。

同じ 2.1.7 Embedded で `isql` を実行した。純正 DLL では EUC の `VARCHAR` を作成すると `-802`、`SET NAMES EUCJ_0208` で接続すると `-924` を再現した。検証用 DLL では列作成と日本語データ3行の挿入に成功し、別プロセスから EUC で接続して3行を確認した。試験 SQL とログは `isql-69de0d95/` にある。

Win32 Release の単一スレッド CLI で、次の42件を再試験した。2.1 の変換元は修正済み DLL を入れた 2.1.7 サーバーを使用した。2.1 の変換先は公式 2.1.7 Embedded と従来の 2.1.6 Embedded のコピーで検証し、それぞれ42件すべて一致した。

| ケース | 2.1.7 Embedded を使用 | 2.1.6 Embedded を使用 |
| --- | ---: | ---: |
| 各変換元 → 2.1 `EUCJ_0208`（2.1 EUC 同士を除く） | 20/20 | 20/20 |
| 2.1 `EUCJ_0208` → 各変換先（2.1 EUC 同士を含む） | 21/21 | 21/21 |
| 2.5 `EUCJ_0208` → 2.1 `NONE`（明示的な EUC 列あり） | 1/1 | 1/1 |

各ケースで ID 順に3行を読み戻し、日本語の UTF-8 バイト列、NULL 文字列、BLOB の NULL と空の区別、非空 BLOB の全バイトを照合した。結果 CSV とケース別のログは `matrix-7d8c8d58/` と `matrix-216-091c2390/` にある。今回の検証範囲は Win32 の単一スレッド。配布 DLL と配布 ZIP は変更していない。

### Firebird 3.0・4.0・5.0 を変換先にする

fbconverter と同じアーキテクチャ（Win32 は x86、Win64 は x64）の公式 Windows ZIP を、実行ファイルに隣接する `Embedded` 以下に展開する。ZIP 内のフォルダ構成を保持し、表の DLL と `firebird.conf` が指定位置に来るように配置する。ICU DLL、`intl`、`plugins` などの同梱ファイルも同じルートに置く。 [Firebird 3.0](https://www.firebirdsql.org/en/firebird-3-0/) · [Firebird 4.0](https://www.firebirdsql.org/en/firebird-4-0/) · [Firebird 5.0](https://firebirdsql.org/en/firebird-5-0/)

| 変換先 | クライアント DLL | エンジンプラグイン | 新規 DB の ODS |
| --- | --- | --- | --- |
| 3.0 | `Embedded/3.0/fbclient.dll` | `Embedded/3.0/plugins/engine12.dll` | 12.0 |
| 4.0 | `Embedded/4.0/fbclient.dll` | `Embedded/4.0/plugins/engine13.dll` | 13.0 |
| 5.0 | `Embedded/5.0/fbclient.dll` | `Embedded/5.0/plugins/engine13.dll` | 13.1 |

配置後に GUI を起動し、変換先で該当バージョンの Embedded を選ぶ。CLI では `--destination-version 3.0`、`4.0` または `5.0` を指定する。サーバーをインストールせずに新しい DB を作成できる。

```powershell
.\fbconverter.exe --convert --source C:\data\source.fdb --destination C:\data\migrated.fdb `
  --source-version 2.5 --destination-version 5.0 --source-host localhost `
  --source-charset UTF8 --destination-charset UTF8 --single-thread | Out-Host
```

3.0 以降の Embedded は、作成・通常接続・並列ワーカー接続の DPB で `Providers` に対応するエンジンを指定する。エンジンの設定は公式 ZIP 内の `firebird.conf` を使う。DB の作成は Zeos の `isc_create_database` 経路を使い、既定文字セット、ページサイズ、SQL dialect を DPB に渡す。既定文字セットを `CP943C` にした場合も、作成・データ転送の接続文字セットは UTF-8 を使う。 [Firebird の Providers 設定](https://www.firebirdsql.org/docs/html/en/refdocs/fbconf/firebird-configuration-reference.html)

変換先のエンジンが受理する DDL を使う必要がある。たとえば従来の外部関数（UDF）は 4.0・5.0 の既定設定で無効になっているため、移行元で使っている場合は関数と呼び出し元の移行が別途必要になる。 [Firebird 4.0 の互換性情報](https://www.firebirdsql.org/file/documentation/release_notes/html/en/4_0/rlsnotes40.html)

#### 3.0・4.0・5.0 の検証結果

Delphi 13.1 の MSBuild で Win32／Win64 Release をビルドし、公式 3.0.14／4.0.7／5.0.4 の Embedded を使用した。変換元は 1.5.6 の `SJIS_0208`、2.0.7 の `EUCJ_0208`、CORE-2095 を修正した 2.1.7 の `EUCJ_0208`、2.5.9 の `UTF8`。Win64 の 1.5／2.0 サーバーには 2.5.4 クライアントを用いた。これら4種類の変換元から、新しい3系列それぞれの `SJIS_0208`、`EUCJ_0208`、`NONE`、`UNICODE_FSS`、`UTF8`、`CP943C` へ単一スレッドで変換し、両アーキテクチャの計144件すべてで、日本語と NULL 文字列、NULL・空・非空の BLOB が一致した。

1.5.6 の `SJIS_0208` から新しい3系列の `UTF8` へ、3,257行を単一スレッドと並列で移した。Win32／Win64 の全12出力を ID 順に読み戻し、変換元と全行が一致した。並列は4ワーカー設定で各14ブロックの完了を確認した。各 ZIP の `gstat -h` で ODS が上表のとおりであることも確認した。従来の 2.5 Embedded 宛ては `NONE` と `UNICODE_FSS` の2ケースを Win32／Win64 で再試験し、全件一致した。

検証ログは `Source/Win32/Release/fb-smoke/modern/` の `test-b740a4cd/results.csv` と `parallel-d3459257/` にある。

### Firebird 1.0.3 を変換元にする

インストール済みの Firebird 1.0.3.972 に `127.0.0.1:3050` で接続した。`isql` でサーバーの版と ODS 10.0 を確認した。Win32 は同じ版の `gds32.dll` を `Client/fbclient10.dll` として配置した。Win64 は 2.5.4 の x64 クライアントを同名で配置し、同じ 1.0.3 サーバーに接続した。 [Firebird 1.0 の Windows クライアントの配置](https://www.firebirdsql.org/manual/qsg10-disk-locations.html)

1.0.3 のサービス応答は `WI-V6.2.972 Firebird 1.0.3`。fbconverter は後半の製品バージョンを読んで 1.0 と判定する。Zeos r8528 に [1.0.3 対応パッチ](zeos-r8528-firebird-1.0.3.patch) を適用し、製品バージョンのマイナー番号にリリース番号が含まれるケースを処理した。また、1.0 にない `COALESCE` を列メタデータの照会から外した。列の NULL フラグが NULL のときは取得済みのドメインのフラグを使用する。

変換元は `SJIS_0208`、`EUCJ_0208`、`NONE`、`UNICODE_FSS` の各3行。`NONE` には Shift_JIS のバイト列を格納した。変換先で使える文字セットを全て組み合わせ、日本語2行、NULL 文字列、NULL・空・4バイトの BLOB を読み戻した。

| 変換先 Embedded | Win32 | Win64 |
| --- | ---: | ---: |
| 1.5.6 | 16/16 | 対象外 |
| 2.0.7 | 20/20 | 対象外 |
| 2.1（Win32 は 2.1.6、Win64 は 2.1.7） | 24/24 | 24/24 |
| 2.5（Win32 は 2.5.3、Win64 は 2.5.4） | 24/24 | 24/24 |
| 3.0.14 | 24/24 | 24/24 |
| 4.0.7 | 24/24 | 24/24 |
| 5.0.4 | 24/24 | 24/24 |
| 合計 | 156/156 | 120/120 |

1.5・2.0 の公式 Windows Embedded は Win32 のみ。2.1 の表は両アーキテクチャとも CORE-2095 を修正した検証用 `fbintl.dll` の結果。Win64 2.1.7 の純正 DLL では `EUCJ_0208` 宛て4件が `-924` で失敗し、検証用 DLL では同じ24件がすべて一致した。x64 DLL は同じ 2.1.7 ソースとパッチを MSVC 14.44 x64 でビルドし、公式 x64 ICU 3.0 にリンクした。配布 ZIP とインストール済みサービスの設定は変更していない。 [1.5 の配布一覧](https://firebirdsql.org/en/firebird-1-5/) · [2.0 の配布一覧](https://firebirdsql.org/en/firebird-2-0/) · [2.1 の配布一覧](https://www.firebirdsql.org/en/firebird-2-1/)

ケース別ログは `Source/Win32/Release/fb-smoke/fb103-4d8eb051/` 以下の `matrix-e2e63fdf/results.csv` と `matrix-x64-core2095-e7e46545/results.csv` にある。

1.0.3 の別の試験 DB から、2.5 と 5.0 へ 3,257 行を単一スレッドと並列で変換した。Win32／Win64 の全8出力で、ID、日本語、NULL 文字列、NULL・空・非空の BLOB が一致した。並列のログには `SAMPLES` の14ブロックと別テーブルの1ブロックがあり、別テーブルはドメインの `NOT NULL` 制約と値も一致した。ログは `parallel-6e8e958f/` にある。

### Firebird 3.0 → 4.0・5.0 と 4.0 → 5.0

公式 3.0.14／4.0.7 のサーバーに TCP で接続し、4.0.7／5.0.4 Embedded に移行した。Win32／Win64 の合計6件で、`isql` を使って ID 順の全行と列の型コード、サブタイプ、精度、スケール、長さ、文字セット ID を比較した。日本語と NULL 文字列、NULL・空・非空の BLOB も全件一致した。試験ログは `Source/Win32/Release/fb-smoke/modern/typed-20261112-082123/cases-2d90e589/` にある。

- 3.0 の試験では `BOOLEAN` の TRUE、FALSE、NULL を照合した。
- 4.0 の試験では `INT128` の上下限、`NUMERIC(38,9)`、`DECIMAL(38,0)`、`DECFLOAT(16/34)`、`TIME WITH TIME ZONE`、`TIMESTAMP WITH TIME ZONE`、`BINARY`、`VARBINARY`、`FLOAT(24/53)` を照合した。`DECFLOAT` は指数表記、符号付きゼロ、NaN、正負の無限大を含む。時刻は地域名と数値オフセットを試し、夏時間終了時に同じ現地時刻となる2つの UTC 時刻を区別した。バイナリ列は高位ビットを含むバイト列を比較し、`VARBINARY` の空と NULL も区別した。 [Firebird 4.0 の追加型](https://www.firebirdsql.org/file/documentation/release_notes/html/en/4_0/rlsnotes40.html)

旧来の Zeos 接続を使う経路では、`INT128` 系の数値、`DECFLOAT`、タイムゾーン付き型をソース SQL で文字列に変換し、変換先の SQL で元の型に戻す。タイムゾーン付き型は UTC の値と元のタイムゾーンを両方渡す。列定義はシステムカタログから再構成し、`NUMERIC` と `DECIMAL`、`BINARY`／`VARBINARY` と通常の `OCTETS` 列のサブタイプを保つ。

4.0 → 5.0 は1,005行でも並列試験を行った。Win32／Win64 とも4ワーカー設定で5ブロックの完了を確認し、全行と列定義が一致した。ログは同じ検証フォルダの `cases-2b57fef6/` にある。

## 2.0.0 の配布 ZIP

`Release/fbconverter-2.0.0_WIN32.zip` と `Release/fbconverter-2.0.0_WIN64.zip` は、Delphi 13.1 でビルドした製品バージョン 2.0.0（ファイルバージョン 2.0.0.0）を収録する。両版に Firebird 2.1.7 Embedded と同じアーキテクチャの CORE-2095 修正版 `intl/fbintl.dll`、3.0.14・4.0.7・5.0.4 Embedded、stdio 用 `fbconverter_mcp.py` を同梱した。従来の Embedded は Win32 の 1.5・2.0 と両版の 2.5 を収録する。MCP は展開先で同じフォルダの `fbconverter.exe` を参照する。対応ソース、ライセンス、ビルド手順は両版の `Sources/source-2.0.0.tar.gz` にある。

両 ZIP を個別に展開し、実行ファイルと DLL のアーキテクチャおよび 2.1.7 の修正版 DLL のハッシュを確認した。展開先の CLI で 2.1 EUC 関連12件、3.0 → 4.0・5.0 と 4.0 → 5.0 の追加型付き移行6件が一致した。同梱の MCP からも両版で 1.0.3 → 5.0 を実行し、既存の出力先への再実行が終了コード 2 になり、ファイルのハッシュが変わらないことを確認した。

## 2.0.1 の配布 ZIP と終了表示の検証

`Release/fbconverter-2.0.1_WIN32.zip` と `Release/fbconverter-2.0.1_WIN64.zip` は、Delphi 13.1 でビルドした製品バージョン 2.0.1（ファイルバージョン 2.0.1.0）を収録する。2.0.0 の配布 ZIP をベースに、実行ファイルと対応ソースのアーカイブ `Sources/source-2.0.1.tar.gz` を更新した。両 ZIP を展開して `fbconverter.exe --help` の起動を確認し、公開した ZIP の SHA-256 が手元の ZIP と一致することを確認した。 [v2.0.1 リリース](https://github.com/tomneko/fbconverter/releases/tag/v2.0.1)

GUI で変換中に SQL エラーが記録された場合は、終了時に `Conversion completed with errors. See the error log.` を表示する。公開版で Firebird 3.0 → 1.5 の変換を再試行し、`VARCHAR(255)` 3列の複合 UNIQUE インデックス `IDX01_INIDATA` の作成がキー長制限で失敗したあと、GUI のログ末尾にこのメッセージが出ることを確認した。
