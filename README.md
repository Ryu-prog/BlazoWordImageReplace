# WordImageReplace
Wordドキュメント（.docx）内のヘッダー画像を、ブラウザ上だけで安全かつ高速に差し替えるためのWebアプリケーションです。

## 🚀 デモサイト

**[こちらからすぐに試せます](https://ryu-prog.github.io/BlazoWordImageReplace/)**

> ※サーバーへのファイルアップロードは一切行われません。ブラウザ内で処理が完結します。

# 概要
このツールは Blazor WebAssembly を利用しており、Wordファイルの内部構造をブラウザ内で直接操作します。サーバーにファイルをアップロードすることなくクライアントサイドのみで処理が完結するため、機密性の高い文書でも安心して利用できるのが特徴です。

# 主な機能
- インストール不要: Webブラウザがあれば、Windows/Macを問わずどこでも動作します。
- プライバシー保護: すべての処理はブラウザ内のメモリで実行されます。ファイルデータが外部サーバーに送信されることはありません。
- 画像一括置換: 表面（Default）と裏面（Even）のヘッダー画像を個別に、または同時に置換可能です。
- リアルタイムプレビュー: 選択した差し替え用画像をその場で確認できます。
- 自動サイズ計算: SixLabors.ImageSharp により画像の解像度を解析し、Word内での表示サイズ（EMU）を適切に調整します。

## 自動公開 (CI/CD) の仕組み

このプロジェクトは GitHub Actions を使用して、`master` ブランチへのプッシュ時に自動的に GitHub Pages へデプロイされます。Blazor WASM 特有の課題を以下の手順で解決しています。

### 1. ワークフロー構成 ([deploy.yml](.github/workflows/deploy.yml))
- **.NET 9.0 ビルド**: 最新の SDK を使用して Release ビルドを作成します。
- **ハッシュ付きファイルの互換性維持**: 
  ビルド時に生成される `blazor.webassembly.*.js` を、ハッシュなしの `blazor.webassembly.js` としてコピーすることで、`index.html` からの参照切れを防止しています。
- **Base Path の自動変換**: 
  ローカルデバッグ時は `<base href="/" />` を維持し、デプロイ時のみ `sed` コマンドで `/BlazoWordImageReplace/` へ動的に書き換えます。
- **リロード対策 (404.html)**: 
  GitHub Pages の仕様による SPA リロード時の 404 エラーを防ぐため、`index.html` を `404.html` として複製しています。
- **Jekyll の無効化**: 
  `.nojekyll` ファイルを生成し、`_framework` などのアンダースコア付きフォルダが無視されるのを防いでいます。



### 2. 公開手順
1. `Settings` > `Actions` > `General` にて、`Workflow permissions` を **Read and write permissions** に変更。
2. `Settings` > `Pages` にて、`Build and deployment` の Source を **Deploy from a branch**、Branch を **gh-pages** に設定。
3. コードを `master` ブランチにプッシュすると、自動的にビルドとデプロイが実行されます。

## 技術スタック
- **Framework**: Blazor WebAssembly (.NET 9.0)
- **Word操作**: [Open XML SDK](https://github.com/dotnet/Open-XML-SDK)
- **画像解析**: [SixLabors.ImageSharp](https://github.com/SixLabors/ImageSharp) (WASM対応の完全マネージドライブラリ)


# 動作環境
- モダンなWebブラウザ（Chrome, Edge, Firefox, Safari等）

- インターネット接続（初期読み込み時のみ）

# 使い方
1. Wordファイルの選択: 編集対象の .docx ファイルを選択します。

1. 画像の選択: 差し替えたい画像（.png, .jpg）を選択します。選択後、画面右側にプレビューが表示されます。

1. 対象の選択: 表面のみ、あるいは裏面のみなど、差し替えたい箇所にチェックを入れます。

1. 実行: 「画像差し替え実行」をクリックすると、加工されたWordファイルが自動的にダウンロードされます。s

# 開発者向け情報
核となるロジックは OpenXmlWordHeaderReplacer.cs に集約されています。

- メモリ効率: 物理パスを使用せず byte[] と Stream ですべての処理を行うよう最適化されています。

- クロスプラットフォーム: Windows APIに依存する System.Drawing を排除し、WASM環境に完全対応させています。

ライセンス
MIT License

## 依存ライブラリ
このプロジェクトでは [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) を使用しています。
ローカルでビルドする場合は、以下のコマンドを実行してパッケージを追加してください。

```bash
dotnet add package DocumentFormat.OpenXml