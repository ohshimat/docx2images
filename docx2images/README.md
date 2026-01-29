# docx2images

Microsoft Word `.docx` ファイルから画像を抽出し、個別の画像ファイルとして保存します。

## 説明

このPythonスクリプトは、`.docx` Wordファイルに埋め込まれたすべての画像を自動的に抽出し、カスタム命名形式で個別の画像ファイルとして保存します。オプションで画像のDPI（解像度）を変更することもできます。

## 機能

- ✅ .docxファイルから埋め込まれた画像をすべて抽出
- ✅ 自動画像形式検出（JPEG、PNG、GIF、BMP、WEBP、ICO、SVG）
- ✅ カスタムファイル名プレフィックス（ユーザー指定の英数字文字列）
- ✅ 連番付け（figure1、figure2など）
- ✅ DPI（解像度）変換機能（オプション）
- ✅ 包括的なエラーハンドリング

## 必要要件

- Python 3.6以上
- Pillow（DPI変換機能を使用する場合）

## インストール

1. このリポジトリをクローン：
```bash
git clone https://github.com/ohshimat/docx2images.git
cd docx2images
```

2. 依存関係をインストール（DPI変換機能を使用する場合）：
```bash
pip install -r requirements.txt
```

3. スクリプトを実行可能にする（オプション）：
```bash
chmod +x extract_images.py
```

## 使い方

### 基本的な使い方

.docxファイルから画像を抽出：
```bash
python extract_images.py --file example.docx
```

これにより、現在のディレクトリに `user-figure1.jpg`、`user-figure2.png` などとして画像が保存されます。

### ユーザープレフィックスの指定

カスタムプレフィックスで画像を抽出：
```bash
python extract_images.py --file example.docx --user johndoe
```

これにより、`johndoe-figure1.jpg`、`johndoe-figure2.png` などとして保存されます。

### 出力ディレクトリの指定

特定のディレクトリに画像を保存：
```bash
python extract_images.py --file example.docx --user alice --output ./images
```

これにより、`images` ディレクトリが作成され、そこにファイルが保存されます。

### DPI（解像度）の変更

画像を指定したDPIに変換：
```bash
python extract_images.py --file example.docx --dpi 350
```

これにより、すべての画像が350dpiに変換されて保存されます。

### コマンドラインオプション

```
--file, -f    .docxファイルへのパス（必須）
--user, -u    ファイル名のユーザー指定英数字プレフィックス（デフォルト：user）
--output, -o  抽出された画像の出力ディレクトリ（デフォルト：カレントディレクトリ）
--dpi, -d     画像のDPI（解像度）を指定（例：300、350）。指定しない場合は元の解像度のまま
```

### ヘルプ

ヘルプ情報を表示：
```bash
python extract_images.py --help
```

## 使用例

```bash
# デフォルト設定で画像を抽出
python extract_images.py --file document.docx

# カスタムユーザープレフィックス
python extract_images.py --file report.docx --user projectA

# 特定のディレクトリに保存
python extract_images.py --file presentation.docx --user team1 --output ./extracted_images

# 350dpiに変換して保存
python extract_images.py --file document.docx --dpi 350

# すべてのオプションを組み合わせ
python extract_images.py --file report.docx --user project --dpi 300 --output ./output

# 短縮形式の引数
python extract_images.py -f data.docx -u researcher -d 350 -o ./output
```

## ファイル命名形式

画像は次の形式で保存されます：`<user>-figure<X>.<ext>`

各部分の説明：
- `user` はユーザー指定の英数字文字列（デフォルト：「user」）
- `X` は1から始まる連番
- `ext` は元の画像形式（jpg、png、gif、bmp、webp、ico、svg）

## エラーハンドリング

スクリプトは以下のエラーに対応します：
- 無効なファイルパス
- .docx以外のファイル
- 存在しないまたはアクセス不可能な入力ファイル
- 無効なユーザープレフィックス（英数字である必要があります）
- 破損した.docxファイル

## 技術詳細

`.docx` ファイルは本質的にXMLファイルとメディアを含むZIPアーカイブです。このスクリプトは：
1. .docxファイルをZIPアーカイブとして開く
2. `word/media/` ディレクトリ内のすべてのファイルを検索
3. ファイルシグネチャ（マジックバイト）を調べて画像形式を判定
4. 適切な拡張子で各画像を抽出して保存

## サポートされている画像形式

- JPEG (.jpg)
- PNG (.png)
- GIF (.gif)
- BMP (.bmp)
- WebP (.webp)
- ICO (.ico)
- SVG (.svg)

## ライセンス

このプロジェクトはオープンソースで、MITライセンスの下で利用可能です。

## 貢献

貢献を歓迎します！プルリクエストをお気軽に提出してください。

## 作成者

ohshimat/docx2imagesリポジトリ用に作成されました。
