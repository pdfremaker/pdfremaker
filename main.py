"""
replitで変更したデータをGitHubに反映させるときは次のコードをShellにコピペ

git add .
git commit -m "update"
git push

↑git commit -m ""の""の中に更新内容を書く
別にupdateのままでもおけ丸水産
"""

# Import the necessary modules
from flask import Flask, request, render_template_string, jsonify, render_template, send_file
import os
import fitz  # PyMuPDF
from werkzeug.utils import secure_filename
import html
import firebase_admin
from firebase_admin import credentials, firestore
import json
from weasyprint import HTML
from weasyprint.urls import path2url  # ローカルファイルのパスをURLに変換するために必要
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from typing import Optional, Dict, Any

# プログラム起動のあいさつ
print("(;^ω^)起動中...")
print("DEBUG: fitz module path:", getattr(fitz, "__file__", None))

# まず関数を定義
app_root = os.path.dirname(os.path.abspath(__file__))

# ←ここに追加
FONT_FILE_MAP = {
    "MS Mincho": "fonts/MSMincho.ttf",
    "MS Gothic": "fonts/MSGothic.ttf",
    "Noto Sans JP": "fonts/NotoSansJP-Regular.ttf",
    "IPAexGothic": "ipaexg.ttf",
    "明朝体, serif": "fonts/Mincho.ttf",  # ←追加
    "Verdana, sans-serif": "fonts/Verdana.ttf"  # 必要なら
}


def get_font_path(app_root, font_family_name="IPAexGothic"):
    font_file = FONT_FILE_MAP.get(font_family_name, "ipaexg.ttf")
    font_path = os.path.join(app_root, font_file)
    font_path = os.path.abspath(font_path)
    if not os.path.exists(font_path):
        print(f"⚠️ フォントファイルが存在しません: {font_path}")
        font_path = os.path.join(app_root, "ipaexg.ttf")
    return font_path

# 関数を呼び出す
app_root = os.path.dirname(os.path.abspath(__file__))
font_path = get_font_path(app_root, "IPAexGothic")
font_url = path2url(font_path)

# Firebaseを初期化
cred_json = os.getenv("GOOGLE_APPLICATION_CREDENTIALS_JSON")
if cred_json:
    cred_dict = json.loads(cred_json)
    cred = credentials.Certificate(cred_dict)
else:
    cred = credentials.Certificate("serAccoCaMnFv.json")

firebase_admin.initialize_app(cred)
db = firestore.client()


def get_document(collection_name, doc_id):
    doc_ref = db.collection(collection_name).document(doc_id)
    docf = doc_ref.get()
    return docf.to_dict() if docf.exists else None


app = Flask(__name__)
app_root = os.path.dirname(os.path.abspath(__file__))
JAPANESE_FONT_NAME = 'HeiseiKakuGo-W5'
pdfmetrics.registerFont(UnicodeCIDFont(JAPANESE_FONT_NAME))
student_font = JAPANESE_FONT_NAME
student_font_size = 12
student_line_height = 4
UPLOAD_FOLDER = os.path.join(app.root_path, "uploads")
OUTPUT_FOLDER = os.path.join(app.root_path, "output")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# PDF再構築のHTML
HTML_FORM = """
<!doctype html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PDF Extractor & Rebuilder</title>
    <style>
        body { font-family: 'Helvetica Neue', Arial, sans-serif; margin: 40px; background-color: #f4f4f9; color: #333; }
        .container { max-width: 600px; margin: auto; text-align: center; }
        h1 { color: #5a5a5a; }
        p { color: #666; }
        form { background: white; padding: 2em; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-top: 1em; }
        input[type=file], input[type=text] { display: block; margin: 0 auto 1em auto; border: 1px solid #ccc; padding: 10px; border-radius: 5px; width: 95%; box-sizing: border-box;}
        input[type=button], input[type=submit] {
            background-color: #007bff; color: white; padding: 10px 20px;
            border: none; border-radius: 5px; cursor: pointer; font-size: 16px; width: 100%;
            transition: background-color 0.3s ease;
        }
        input[type=button]:hover, input[type=submit]:hover { background-color: #0056b3; }
        #student-info {
            background: #fff; padding: 20px; border-radius: 12px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.1); margin-top: 20px; text-align: left;
        }

        /* 1. ボタンリンクのスタイル */
        .button-link {
            text-decoration: none;
            background-color:#228b22;
            color: white;
            padding: 6px 6px;
            border-radius: 5px;
            display: inline-block;
            transition: background-color 0.3s;
            margin-right: 10px; /* 生徒ID入力欄との余白 */
        }
        .button-link:hover {
            background-color: #333333;
        }

        /* 2. ボタンを右揃えにする親要素のスタイル */
        .align-right-container {
            text-align: right;
            width: 100%;
            margin-bottom: 10px;
        }
        .centered-heading {
            text-align: center;
        }
    </style>
</head>

<body>
    <div class="container">
        <h1>PDFアップローダー</h1>
        <p>生徒IDを入力して設定を反映し、PDFをアップロードしてください。</p>

        <form method=post enctype=multipart/form-data>

            <div class="align-right-container">
                <a href="https://03b5d2fe-5aad-4c9d-969d-32d1d7c1af6e-00-k58ex9qtsoqz.pike.replit.dev/" class="button-link">
                    <b>生徒設定編集画面へ</b>
                </a>
            </div>
            <h2 class="centered-heading">1. 反映する生徒設定情報の選択</h2>

            <input type="text" id="student-id" name="student_id" placeholder="生徒IDを入力してください">
            <input type="button" id="fetch-button" value="データ取得 & 確認 & 反映">
            <div id="student-info">設定情報をここに表示します</div>

            <h2 style="margin-top: 2em;">2. PDFファイルのアップロード</h2>
            <input type=file name=file accept=".pdf" required>
            <input type=submit value="アップロードして処理">
        </form>
    </div>

    <script>
    document.getElementById("fetch-button").addEventListener("click", async function() {
        const id = document.getElementById("student-id").value;
        if (!id) {
            alert("生徒IDを入力してください。");
            return;
        }
        const res = await fetch(`/get_message?id=${encodeURIComponent(id)}`);
        const data = await res.json();
        const div = document.getElementById("student-info");

        if (data.error) {
            div.innerHTML = `<p style="color:red;">${data.error}</p>`;
        } else {
            div.innerHTML = `
                <h3>現在の生徒設定情報（加算値）</h3>
                <p><strong>ID:</strong> ${data.id}</p>
                <p><strong>フォント（上書き）:</strong> ${data.fontSelect}</p>
                <p><strong>文字サイズ（追加）:</strong> +${data.fontSize}</p>
                <p><strong>行間（追加）:</strong> +${data.lineHeight}</p>
                <p style="color:green; font-weight:bold;">設定が確認できました。ファイルをアップロードしてください。</p>`;
        }
    });
    document.getElementById("fetch-button").addEventListener("click", async function() {
        const id = document.getElementById("student-id").value;
        if (!id) {
            alert("生徒IDを入力してください。");
            return;
        }
        const res = await fetch(`/get_message?id=${encodeURIComponent(id)}`);
        const data = await res.json();
        const div = document.getElementById("student-info");

        if (data.error) {
            div.innerHTML = `<p style="color:red;">${data.error}</p>`;
        } else {
            div.innerHTML = `
                <h3>現在の生徒設定情報（加算値）</h3>
                <p><strong>ID:</strong> ${data.id}</p>
                <p><strong>フォント（上書き）:</strong> ${data.fontSelect}</p>
                <p><strong>文字サイズ（追加）:</strong> +${data.fontSize}</p>
                <p><strong>行間（追加）:</strong> +${data.lineHeight}</p>
                <p style="color:green; font-weight:bold;">設定が確認できました。ファイルをアップロードしてください。</p>`;
        }
    });

    </script>
</body>
</html>
"""


# Firestoreの情報変える生徒側へ行く
@app.route('/return')
def return_page():
    return HTML_FORM


# Firestoreの情報変える生徒側へ行く
@app.route('/edit')
def another_page():
    return render_template('stuSets.html')
if not os.path.exists(font_path):
    app.logger.warning(f"フォントファイルが見つかりません: {font_path}")


# Firestoreのメッセージ取得
@app.route("/get_message", methods=["GET"])
def get_message_api():
    doc_id = request.args.get("id", "").strip()
    if not doc_id:
        return jsonify({"error": "IDが指定されていません"}), 400
    data = get_document("messages", doc_id)
    if not data:
        return jsonify({"error": f"ID '{doc_id}' は存在しません"}), 404
    return jsonify({
        k: data.get(k, "N/A")
        for k in ["fontSelect", "fontSize", "lineHeight"]
    } | {"id": doc_id})


@app.route("/", methods=["GET", "POST"])
def upload_pdf():
    if request.method == "POST":
        if "file" not in request.files or not request.files["file"].filename:
            return "ファイルが選択されていません。"

        uploaded_file = request.files["file"]
        filename = uploaded_file.filename or ""  # None対策

        student_id = request.form.get("student_id", "").strip()
        firebase_settings = None  # 未定義エラー防止の初期化

        if student_id:
            firebase_settings = get_document("messages", student_id)
            if firebase_settings:
                app.logger.info(f"ID '{student_id}' の設定を適用します: {firebase_settings}")
            else:
                app.logger.info(f"ID '{student_id}' は見つかりませんでした。デフォルト設定で処理します。")

        if filename.lower().endswith(".pdf"):
            filename = secure_filename(filename)
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            uploaded_file.save(filepath)
            result_html = process_pdf(filepath, firebase_settings)
            return result_html
        else:
            return "PDFファイルをアップロードしてください。"

    return render_template_string(HTML_FORM)


@app.route('/outputs/<path:filepath>')
def serve_output_file(filepath):
    """
    filepath: URLパス部分（例: mypdf/mypage.pdf や basename/mypage.png）
    出力フォルダ（OUTPUT_FOLDER）配下のファイルのみを返す。安全チェックを厳密に行う。
    """

    safe_path = os.path.normpath(filepath)  # 正規化（これで .. などは取り除かれる）
    full_path = os.path.join(OUTPUT_FOLDER, safe_path)  # 実際のファイルパスを構成

    # 重要: 絶対パスにしてOUTPUT_FOLDERの下にあることを確認（ディレクトリ脱出防止）
    full_path = os.path.abspath(full_path)
    output_folder_abs = os.path.abspath(OUTPUT_FOLDER)
    if not full_path.startswith(output_folder_abs + os.path.sep
                                ) and full_path != output_folder_abs:
        return "不正なパスです", 400

    # ファイルが存在するか確認
    if not os.path.isfile(full_path):
        return "ファイルが見つかりません。", 404

    # 安全に送信（ダウンロードとして返す）
    try:
        return send_file(full_path, as_attachment=True)
    except Exception as e:
        app.logger.exception("ファイル送信でエラー")
        return f"ファイル送信中にエラーが発生しました: {e}", 500

# --- 変更後 ---
# --- 改良後 ---
# --- 変更後 ---
def create_pdf_with_weasyprint(neo_content, output_pdf_path, app_root, firebase_settings=None):
    # Firestore設定の反映
    font_family_name = firebase_settings.get("fontSelect", "IPAexGothic") if firebase_settings else "IPAexGothic"
    font_size = float(firebase_settings.get("fontSize", 1.0)) if firebase_settings else 1.0
    line_height = float(firebase_settings.get("lineHeight", 1.0)) if firebase_settings else 1.0

    # --- フォントファイルのマッピング ---
    font_file = FONT_FILE_MAP.get(font_family_name, "ipaexg.ttf")
    font_path = os.path.join(app_root, font_file)
    font_path = os.path.abspath(font_path)  # ← ここを追加！

    if not os.path.exists(font_path):
        print(f"⚠️ フォントファイルが存在しません: {font_path}")
        font_path = os.path.join(app_root, "ipaexg.ttf")

    # --- WeasyPrintが認識できるURL形式に変換 ---
    font_url = path2url(font_path)
    font_family_css = "StudentCustomFont"  # 固定名に統一

    print(f"🟢 Using font: {font_family_name}")
    print(f"🟢 Font file path: {font_path}")
    print(f"🟢 Font URL: {font_url}")

    # --- CSS生成 ---
    css_string = f"""
    @font-face {{
        font-family: '{font_family_css}';
        src: url('{font_url}');
        font-weight: normal;
        font-style: normal;
    }}
    body {{
        font-family: '{font_family_css}', 'IPAexGothic', sans-serif;
        font-size: {12 * font_size}pt;
        line-height: {1.6 * line_height};
        word-wrap: break-word;
    }}
    """

    html_string = f"""
    <!DOCTYPE html>
    <html lang="ja">
        <head>
            <meta charset="utf-8">
            <style>{css_string}</style>
        </head>
        <body>{neo_content}</body>
    </html>
    """

    try:
        HTML(string=html_string, base_url=app_root).write_pdf(output_pdf_path)
        return (True, None)
    except Exception as e:
        return (False, f"WeasyPrintエラー: {e}")


def create_styled_html(text_content, app_root):
    """
    NEO形式のテキストをHTMLに変換。
    - [行間]
    - [フォント:フォント名:サイズ:weight]
    - [画像:ファイルパス:x:y:width:height]
    """
    lines = text_content.strip().split('\n')
    styled_html = ""
    for line in lines:
        # --- 行間タグの処理 ---
        if line.startswith('[行間]'):
            try:
                line_spacing = float(line.split(']')[0].split('[行間]')[1])
                styled_html += f'<div style="margin-top: {line_spacing}px;"></div>'
            except (ValueError, IndexError):
                continue

        # --- フォントタグの処理 ---
        elif line.startswith('[フォント:'):
            try:
                parts = line.split(']', 3)
                if len(parts) != 4:
                    continue

                meta1, meta2, meta3, text_content_line = parts
                text_content_line = text_content_line.strip()

                font_name_display = meta1.split(':')[1]
                font_size = float(meta2.split(':')[1])
                weight = meta3.split(':')[1]

                style_str = f"font-size: {font_size}px; margin: 0; padding: 0;"

                # フォントファミリーの決定
                style_str += f" font-family: {font_name_display};"


                if weight == 'bold':
                    style_str += " font-weight: bold;"

                styled_html += f'<p style="{style_str}">{html.escape(text_content_line)}</p>'
            except (ValueError, IndexError):
                continue

        # --- 画像タグの処理 ---
        elif line.startswith('[画像:'):
            try:
                parts = line.strip('[]').split(':')[
                    1:]  # [画像:ファイルパス:x:y:width:height]
                if len(parts) != 5:
                    continue

                rel_path = parts[0]
                width = float(parts[3])
                height = float(parts[4])

                abs_path = os.path.join(app_root, rel_path)
                if not os.path.exists(abs_path):
                    styled_html += f'<p style="color:red;">[画像読み込みエラー: {rel_path}]</p>'
                    continue

                image_url = path2url(abs_path)
                styled_html += f'<p><img src="{image_url}" width="{width}" height="{height}"></p>'

            except Exception as e:
                styled_html += f'<p style="color:red;">[画像処理エラー: {e}]</p>'

        # --- 通常テキスト（タグなし） ---
        else:
            styled_html += f'<p>{html.escape(line)}</p>'

    return styled_html


def process_pdf(pdf_path: str, firebase_settings: Optional[Dict[str, Any]] = None) -> str:
    """
    PDFを分解・再構成し、解析結果をHTMLで返す関数。
    firebase_settings に基づきフォントやサイズを適用。
    """
    try:
        # fitz.open() に type: ignore を追加してPyright警告を無効化
        doc = fitz.open(pdf_path)  # type: ignore
    except Exception as e:
        return f"PDFファイルを開けませんでした: {e}"

    basename = os.path.splitext(os.path.basename(pdf_path))[0]
    dir_name = os.path.join(OUTPUT_FOLDER, basename)
    os.makedirs(dir_name, exist_ok=True)

    # --- Firebase設定を取得 ---
    fs_font_override = firebase_settings.get("fontSelect") if firebase_settings else None
    fs_size_add = float(firebase_settings.get("fontSize", 0)) if firebase_settings else 0.0
    fs_line_height_add = float(firebase_settings.get("lineHeight", 0)) if firebase_settings else 0.0 # noqa #加える行間のデータ

    # --- 出力ファイルパス ---
    output_file_OG = os.path.join(dir_name, f"{basename}_OG.txt")
    output_file_NEO = os.path.join(dir_name, f"{basename}_NEO.txt")
    output_file_SORTED = os.path.join(dir_name, f"{basename}_SORTED.txt")

    # --- OGテキスト抽出 ---
    with open(output_file_OG, "w", encoding="utf-8") as f:
        for page in doc:  # type: ignore
            f.write(page.get_text("text"))

    neo_content_lines = []
    sorted_content_lines = []
    image_urls = []
    page_heights = [page.rect.height for page in doc]  # type: ignore

    for page_num, page in enumerate(doc):  # type: ignore
        sorted_content_lines.append(f"\n--- Page {page_num + 1} ---\n")

        # --- テキストブロック＆画像抽出 ---
        text_blocks = page.get_text("dict")["blocks"]
        images = page.get_images(full=True)
        page_elements = []

        # --- テキスト抽出 ---
        for block in text_blocks:
            if block["type"] == 0:  # テキストのみ
                text_content = ""
                for line in block["lines"]:
                    for span in line["spans"]:
                        text_content += span["text"]
                if text_content.strip():
                    page_elements.append({
                        "type": "text",
                        "bbox": block["bbox"],
                        "content": text_content.strip()
                    })

        # --- 画像抽出 ---
        for i, img in enumerate(images):
            try:
                xref = img[0]
                pix = fitz.Pixmap(doc, xref)  # type: ignore
                img_filename = f"image_p{page_num+1}_idx{i}.png"
                img_path_full = os.path.join(dir_name, img_filename)

                if pix.n >= 5:
                    pix = fitz.Pixmap(fitz.csRGB, pix)  # type: ignore
                pix.save(img_path_full)

                relative_path = os.path.join(basename, img_filename).replace("\\", "/")
                image_urls.append(relative_path)

                bbox_info = page.get_image_info(xref)
                bbox = bbox_info[0]["bbox"] if bbox_info else [0, 0, 0, 0]

                page_elements.append({
                    "type": "image",
                    "bbox": bbox,
                    "content": img_path_full
                })
            except Exception as e:
                print(f"⚠️ ページ{page_num+1}の画像抽出に失敗: {e}")

        # --- ソート（上→下、左→右） ---
        page_elements.sort(key=lambda x: (x["bbox"][1], x["bbox"][0]))

        # --- 行間計算＋NEO構築 ---
        previous_y = None
        for element in page_elements:
            current_y = element["bbox"][1]
            if previous_y is not None:
                spacing = current_y - previous_y
                if spacing > 0:
                    neo_content_lines.append(f"[行間]{spacing:.2f}\n")

            if element["type"] == "text":
                text_content = element["content"]
                font_name = fs_font_override or "IPAexGothic, sans-serif"
                final_font_size = 12.0 + fs_size_add
                weight = "bold" if "bold" in text_content.lower() else "normal"

                neo_line = (
                    f"[フォント:{font_name}]"
                    f"[サイズ:{final_font_size:.2f}]"
                    f"[ウェイト:{weight}]{text_content}\n"
                )
                neo_content_lines.append(neo_line)
                sorted_content_lines.append(f"テキスト: {text_content}\n\n")
                previous_y = element["bbox"][3]

            # 画像抽出
            elif element["type"] == "image":
                bbox = element["bbox"]
                # ReportLabや座標再配置向けに Y 座標を計算
                img_y_reportlab = page_heights[page_num] - bbox[3]
                neo_line = (
                    f"[画像:{element['content']}:"
                    f"{bbox[0]:.2f}:{img_y_reportlab:.2f}:"
                    f"{bbox[2]-bbox[0]:.2f}:{bbox[3]-bbox[1]:.2f}]\n"
                )
                neo_content_lines.append(neo_line)
                sorted_content_lines.append(f"[画像] {element['content']} | BBOX: {bbox}\n\n")
                previous_y = bbox[3]

    # --- テキストファイル出力 ---
    neo_content = "".join(neo_content_lines)
    sorted_content = "".join(sorted_content_lines)

    with open(output_file_NEO, "w", encoding="utf-8") as f:
        f.write(neo_content)
    with open(output_file_SORTED, "w", encoding="utf-8") as f:
        f.write(sorted_content)

    # --- OG読み込み ---
    try:
        with open(output_file_OG, "r", encoding="utf-8") as f:
            og_content = f.read()
    except FileNotFoundError:
        og_content = "OGテキストファイルの読み込みに失敗しました。"

    # --- PDF再構成 ---
    recreated_pdf_filename = f"{basename}_recreated.pdf"
    recreated_pdf_path = os.path.join(dir_name, recreated_pdf_filename)
    pdf_created_successfully, _ = create_pdf_with_weasyprint(
        neo_content, recreated_pdf_path, app_root, firebase_settings
    )
    recreated_pdf_url = (
        os.path.join(basename, recreated_pdf_filename).replace("\\", "/")
        if pdf_created_successfully else ""
    )

    # --- HTML生成 ---
    styled_neo_html = create_styled_html(neo_content, app_root)
    og_safe = html.escape(og_content)
    neo_safe = html.escape(neo_content)
    sorted_safe = html.escape(sorted_content)

    image_gallery_html = "".join(
        f'<a href="/outputs/{html.escape(url)}" target="_blank">'
        f'<img src="/outputs/{html.escape(url)}" alt="image"></a>'
        for url in image_urls
    ) or "<p>画像は抽出されませんでした。</p>"

    download_html = ""
    if pdf_created_successfully:
        download_html = (
            f'<div class="download-section">'
            f'<h3>再構成されたPDF</h3>'
            f'<a href="/outputs/{html.escape(recreated_pdf_url)}" '
            f'class="action-link" download>ダウンロード</a></div>'
        )

    doc.close()

    # --- 最終HTML ---
    result_html = f"""
    <!doctype html>
    <html lang="ja">
    <head>
        <meta charset="UTF-8">
        <title>処理結果</title>
        <style>
            body {{ font-family: 'Helvetica Neue', Arial, sans-serif; line-height: 1.6; background-color: #f4f4f9; }}
            .container {{ max-width: 960px; margin: 2em auto; padding: 2em; background: white; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }}
            h2 {{ border-bottom: 2px solid #007bff; color: #333; }}
            .content-box {{ border: 1px solid #ddd; padding: 1em; margin-top: 1em; background: #fdfdfd; max-height: 400px; overflow-y: auto; font-family: monospace; }}
            .styled-content-box {{ border: 1px solid #ddd; padding: 1em; margin-top: 1em; max-height: 400px; overflow-y: auto; }}
            details {{ border: 1px solid #ccc; border-radius: 5px; padding: 0.5em; background: #f9f9f9; }}
            summary {{ font-weight: bold; cursor: pointer; color: #0056b3; }}
            .info, .download-section {{ background: #eef; padding: 1em; border-radius: 8px; margin-bottom: 1.5em; }}
            .image-gallery {{ display: flex; flex-wrap: wrap; gap: 15px; }}
            .image-gallery img {{ border: 2px solid #ddd; border-radius: 5px; max-width: 150px; transition: transform 0.2s; }}
            .image-gallery img:hover {{ transform: scale(1.05); border-color: #007bff; }}
            .action-link {{ display: inline-block; margin-top: 1em; background-color: #28a745; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; font-weight: bold; }}
            .back-link {{ background-color: #6c757d; }}
            .action-link:hover {{ background-color: #218838; }}
            .back-link:hover {{ background-color: #5a6268; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h2>処理完了！</h2>
            <div class="info">
                <p><strong>処理対象ファイル:</strong> {html.escape(os.path.basename(pdf_path))}</p>
                <p><strong>保存先フォルダ:</strong> {html.escape(os.path.abspath(dir_name))}</p>
            </div>
            {download_html}
            <details><summary>スタイル付きNEOプレビュー</summary><div class="styled-content-box">{styled_neo_html}</div></details>
            <details><summary>NEOテキスト</summary><div class="content-box">{neo_safe}</div></details>
            <details><summary>OGテキスト</summary><div class="content-box">{og_safe}</div></details>
            <details><summary>時系列ソート</summary><div class="content-box">{sorted_safe}</div></details>
            <details open><summary>抽出画像 ({len(image_urls)}枚)</summary><div class="image-gallery">{image_gallery_html}</div></details>
            <a href="/" class="action-link back-link">別のファイルを処理する</a>
        </div>
    </body>
    </html>
    """
    return result_html


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=3000)