## fbconverter v2.0.1

変換中に SQL エラーが発生した場合、GUI のログと終了ダイアログにエラーがあったことを表示するよう修正しました。たとえば、移行先の Firebird でインデックスを作成できなかった場合は、エラーログの確認を案内します。

対応する Firebird バージョン、CLI・MCP の使い方、Firebird 2.1.7 Embedded の EUC 修正については [v2.0.0 のリリースノート](https://github.com/tomneko/fbconverter/releases/tag/v2.0.0) を参照してください。

Windows の環境に合わせて WIN32 または WIN64 の ZIP を展開し、`fbconverter.exe` を実行してください。対応ソースとビルド手順は各 ZIP の `Sources/source-2.0.1.tar.gz` に収録しています。
