#!/usr/bin/env python3
"""
.docxファイルから画像を抽出して個別の画像ファイルとして保存します。

このスクリプトは、Microsoft Word .docxファイルに埋め込まれたすべての画像を抽出し、
カスタム命名形式で保存します：<user>-figure<X>.<ext>
オプションで画像のDPI（解像度）を変更できます。
"""

import argparse
import os
import sys
import zipfile
import re
from pathlib import Path
from io import BytesIO

try:
    from PIL import Image
    PILLOW_AVAILABLE = True
except ImportError:
    PILLOW_AVAILABLE = False


def is_valid_prefix(text):
    """テキストが有効なプレフィックスかどうかを確認します（英字、数字、アンダースコアが使用可能）。"""
    return bool(re.match(r'^[a-zA-Z0-9_]+$', text))


def get_image_extension(image_data):
    """
    ファイルシグネチャ（マジックバイト）に基づいて画像ファイル拡張子を判定します。
    
    Args:
        image_data: バイナリ画像データ
        
    Returns:
        ファイル拡張子（例：'jpg'、'png'、'gif'）
    """
    # 一般的な画像形式のマジックバイトをチェック
    if image_data.startswith(b'\xff\xd8\xff'):
        return 'jpg'
    elif image_data.startswith(b'\x89PNG\r\n\x1a\n'):
        return 'png'
    elif image_data.startswith(b'GIF87a') or image_data.startswith(b'GIF89a'):
        return 'gif'
    elif image_data.startswith(b'BM'):
        return 'bmp'
    elif image_data.startswith(b'RIFF') and b'WEBP' in image_data[:12]:
        return 'webp'
    elif image_data.startswith(b'\x00\x00\x01\x00'):
        return 'ico'
    elif image_data.startswith(b'<svg') or image_data.startswith(b'<?xml'):
        return 'svg'
    else:
        # 不明な場合はjpegをデフォルトとする
        return 'jpg'


def convert_image_dpi(image_data, dpi, extension):
    """
    画像のDPI（解像度）を変更します。
    
    Args:
        image_data: バイナリ画像データ
        dpi: 設定するDPI値
        extension: 画像ファイル拡張子
        
    Returns:
        DPIが設定された画像のバイナリデータ
        
    Raises:
        ImportError: Pillowがインストールされていない場合
    """
    if not PILLOW_AVAILABLE:
        raise ImportError("DPI変換にはPillowライブラリが必要です。pip install Pillowでインストールしてください。")
    
    # SVGとICOは除外（DPI設定が適用できない形式）
    if extension in ['svg', 'ico']:
        return image_data
    
    try:
        # 画像を開く
        img = Image.open(BytesIO(image_data))
        
        # 出力バッファを作成
        output = BytesIO()
        
        # 画像形式に応じてDPIを設定して保存
        if extension == 'jpg':
            img.save(output, format='JPEG', dpi=(dpi, dpi), quality=95)
        elif extension == 'png':
            img.save(output, format='PNG', dpi=(dpi, dpi))
        elif extension == 'bmp':
            img.save(output, format='BMP', dpi=(dpi, dpi))
        elif extension == 'webp':
            img.save(output, format='WEBP', dpi=(dpi, dpi), quality=95)
        elif extension == 'gif':
            img.save(output, format='GIF', dpi=(dpi, dpi))
        else:
            # その他の形式はそのまま返す
            return image_data
        
        return output.getvalue()
    except Exception as e:
        # DPI変換に失敗した場合は元の画像データを返す
        print(f"警告: DPI変換に失敗しました: {str(e)}", file=sys.stderr)
        return image_data


def extract_images_from_docx(docx_path, output_dir, user_prefix='user', dpi=None):
    """
    .docxファイルからすべての画像を抽出し、カスタム命名で保存します。
    
    Args:
        docx_path: .docxファイルへのパス
        output_dir: 画像が保存されるディレクトリ
        user_prefix: 画像ファイル名のユーザー指定プレフィックス（デフォルト：'user'）
        dpi: 画像のDPI値（Noneの場合は変換なし）
        
    Returns:
        保存された画像ファイル名のリスト
        
    Raises:
        FileNotFoundError: docxファイルが存在しない場合
        ValueError: ファイルが有効な.docxファイルでない場合
        zipfile.BadZipFile: ファイルが有効なZIPアーカイブでない場合
    """
    # 入力ファイルの存在を確認
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"ファイルが見つかりません: {docx_path}")
    
    # .docxファイルかどうかをチェック
    if not docx_path.lower().endswith('.docx'):
        raise ValueError(f"ファイルは.docxファイルである必要があります: {docx_path}")
    
    # ユーザープレフィックスを検証
    if not is_valid_prefix(user_prefix):
        raise ValueError(f"ユーザープレフィックスは英数字である必要があります: {user_prefix}")
    
    # DPI値の検証
    if dpi is not None:
        if dpi <= 0:
            raise ValueError(f"DPI値は正の数である必要があります: {dpi}")
        if not PILLOW_AVAILABLE:
            raise ImportError("DPI変換にはPillowライブラリが必要です。pip install Pillowでインストールしてください。")
    
    # 出力ディレクトリが存在しない場合は作成
    os.makedirs(output_dir, exist_ok=True)
    
    saved_images = []
    image_counter = 1
    
    try:
        # .docxファイルをZIPアーカイブとして開く
        with zipfile.ZipFile(docx_path, 'r') as docx_zip:
            # アーカイブ内のすべてのファイルをリスト
            file_list = docx_zip.namelist()
            
            # word/media/ディレクトリ内のすべての画像ファイルを検索
            image_files = [f for f in file_list if f.startswith('word/media/')]
            
            if not image_files:
                print(f"{docx_path}に画像が見つかりません")
                return saved_images
            
            # 各画像を抽出して保存
            for image_file in sorted(image_files):
                # 画像データを読み込む
                image_data = docx_zip.read(image_file)
                
                # ファイル拡張子を判定
                extension = get_image_extension(image_data)
                
                # DPI変換が指定されている場合は適用
                if dpi is not None:
                    image_data = convert_image_dpi(image_data, dpi, extension)
                
                # 出力ファイル名を作成
                output_filename = f"{user_prefix}-figure{image_counter}.{extension}"
                output_path = os.path.join(output_dir, output_filename)
                
                # 画像を保存
                with open(output_path, 'wb') as img_file:
                    img_file.write(image_data)
                
                saved_images.append(output_filename)
                dpi_info = f" ({dpi}dpi)" if dpi else ""
                print(f"保存しました: {output_filename}{dpi_info}")
                image_counter += 1
                
    except zipfile.BadZipFile:
        raise ValueError(f"無効または破損した.docxファイル: {docx_path}")
    except (FileNotFoundError, ValueError, ImportError):
        # これらの特定の例外は変更せずに再スロー
        raise
    except Exception as e:
        # 予期しない例外の場合、コンテキストをラップするが元の例外を保持
        raise Exception(f"画像抽出エラー: {str(e)}") from e
    
    return saved_images


def main():
    """コマンドライン引数を処理し、画像抽出を実行するメイン関数。"""
    parser = argparse.ArgumentParser(
        description='.docxファイルから画像を抽出して個別のファイルとして保存します。',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  %(prog)s --file example.docx --user johndoe
  %(prog)s --file document.docx --user alice --output ./images
  %(prog)s --file report.docx --dpi 350
  %(prog)s --file presentation.docx --user project --dpi 300 --output ./output
        """
    )
    
    parser.add_argument(
        '--file', '-f',
        required=True,
        help='.docxファイルへのパス'
    )
    
    parser.add_argument(
        '--user', '-u',
        default='user',
        help='画像ファイル名のユーザー指定英数字プレフィックス（デフォルト：user）'
    )
    
    parser.add_argument(
        '--output', '-o',
        default='.',
        help='抽出された画像の出力ディレクトリ（デフォルト：カレントディレクトリ）'
    )
    
    parser.add_argument(
        '--dpi', '-d',
        type=int,
        default=None,
        help='画像のDPI（解像度）を指定（例：300、350）。指定しない場合は元の解像度のまま'
    )
    
    args = parser.parse_args()
    
    try:
        # 画像を抽出
        saved_images = extract_images_from_docx(
            args.file,
            args.output,
            args.user,
            args.dpi
        )
        
        # サマリーを出力
        if saved_images:
            dpi_info = f"（{args.dpi}dpiに変換）" if args.dpi else ""
            print(f"\n{args.file}から{len(saved_images)}個の画像を正常に抽出しました{dpi_info}")
            print(f"画像の保存先: {os.path.abspath(args.output)}")
        else:
            print(f"\n{args.file}に画像が見つかりません")
        
        return 0
        
    except FileNotFoundError as e:
        print(f"エラー: {e}", file=sys.stderr)
        return 1
    except ValueError as e:
        print(f"エラー: {e}", file=sys.stderr)
        return 1
    except ImportError as e:
        print(f"エラー: {e}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"エラー: {e}", file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
