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
import html
import json
from typing import Optional, Dict, Any
from werkzeug.utils import secure_filename
import pymupdf as fitz  # PyMuPDF
import firebase_admin
from firebase_admin import credentials, firestore
from weasyprint import HTML
from weasyprint.urls import path2url  # ローカルファイルのパスをURLに変換するために必要
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont


# プログラム起動のあいさつ
print("(;^ω^)起動中...")
print(f"DEBUG: fitz module path: {fitz.__file__}")
print(f"DEBUG: fitz.open available: {hasattr(fitz, 'open')}")

# まず関数を定義
app_root = os.path.dirname(os.path.abspath(__file__))

# フォント情報の設定
FONT_FILE_MAP = {
    "明朝体, serif": "ipaexm.ttf",
    "IPAex明朝": "ipaexm.ttf",
    "IPAexゴシック": "ipaexg.ttf",
    "ゴシック体, sans-serif": "ipaexg.ttf",
    "Verdana, sans-serif": "ipaexg.ttf",  # 代替としてIPAexゴシック
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


def create_styled_html(text_content, app_root):
    lines = text_content.strip().split("\n")
    html_out = ""
    for line in lines:
        if line.startswith("[行間]"):
            try:
                sp = float(line.replace("[行間]", "").strip())
                html_out += f'<div style="margin-top:{sp}px;"></div>'
            except:
                continue
        elif line.startswith("[フォント:"):
            try:
                parts = line.split("]")
                font_name = parts[0].split(":")[1]
                size = float(parts[1].split(":")[1])
                weight = parts[2].split(":")[1]
                text = parts[3]
                style = f"font-size:{size}px; font-weight:{weight};"
                html_out += f'<p style="{style}">{html.escape(text)}</p>'
            except:
                continue
        elif line.startswith("[画像:"):
            try:
                parts = line.strip("[]").split(":")[1:]
                if len(parts) != 5:
                    continue
                path, x, y, w, h = parts
                abs_path = os.path.join(app_root, path)
                if os.path.exists(abs_path):
                    url = path2url(abs_path)
                    html_out += f'<p><img src="{url}" width="{w}" height="{h}"></p>'
                else:
                    html_out += f'<p style="color:red;">[画像なし: {path}]</p>'
            except:
                continue
        else:
            html_out += f"<p>{html.escape(line)}</p>"
    return html_out


def create_pdf_with_weasyprint(neo_content, output_pdf_path, app_root):
    try:
        html_body = create_styled_html(neo_content, app_root)
        font_file_path = os.path.join(app_root, "ipaexg.ttf")
        if not os.path.exists(font_file_path):
            return (False, "フォントファイル 'ipaexg.ttf' が見つかりません。")

        font_url = path2url(font_file_path)
        css = f"""
        @font-face {{
            font-family: 'IPAexGothic';
            src: url('{font_url}');
        }}
        body {{
            font-family: 'IPAexGothic', sans-serif;
            font-size: 12pt;
            line-height: 1.6;
        }}
        """
        HTML(string=f"<style>{css}</style>{html_body}", base_url=app_root).write_pdf(output_pdf_path)
        return (True, None)
    except Exception as e:
        return (False, f"WeasyPrintエラー: {e}")


def process_pdf(pdf_path, firebase_settings=None):
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        return f"PDFを開けません: {e}"

    basename = os.path.splitext(os.path.basename(pdf_path))[0]
    dir_name = os.path.join(OUTPUT_FOLDER, basename)
    os.makedirs(dir_name, exist_ok=True)

    # Firebase設定
    fs_font_override = firebase_settings.get('fontSelect') if firebase_settings else None
    fs_size_add = float(firebase_settings.get('fontSize', 0)) if firebase_settings else 0.0

    # 出力ファイルパス
    output_file_OG = os.path.join(dir_name, f"{basename}_OG.txt")
    output_file_NEO = os.path.join(dir_name, f"{basename}_NEO.txt")
    output_file_SORTED = os.path.join(dir_name, f"{basename}_SORTED.txt")

    # OGテキスト保存
    with open(output_file_OG, "w", encoding="utf-8") as f:
        for p in doc:
            f.write(p.get_text("text"))

    neo, sorted_txt, imgs = [], [], []
    page_heights = [p.rect.height for p in doc]

    # --- ページごとの抽出 ---
    for i, page in enumerate(doc):
        sorted_txt.append(f"\n--- Page {i+1} ---\n")
        elements = []

        # テキスト抽出
        for blk in page.get_text("dict")["blocks"]:
            if blk["type"] == 0:
                text = "".join(span["text"] for ln in blk["lines"] for span in ln["spans"]).strip()
                if text:
                    elements.append({"type": "text", "bbox": blk["bbox"], "content": text})

        # 画像抽出
        for j, img in enumerate(page.get_images(full=True)):
            try:
                xref = img[0]
                pix = fitz.Pixmap(doc, xref)
                if pix.n >= 5:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                name = f"image_p{i+1}_{j}.png"
                full = os.path.join(dir_name, name)
                pix.save(full)
                rel = os.path.join(basename, name).replace("\\", "/")
                imgs.append(rel)
                bbox = page.get_image_info(xref)[0]["bbox"]
                elements.append({"type": "image", "bbox": bbox, "content": full})
            except Exception as e:
                print("画像抽出失敗:", e)

        # 要素を座標順でソート
        elements.sort(key=lambda x: (x["bbox"][1], x["bbox"][0]))

        prev_y = None
        for el in elements:
            y = el["bbox"][1]
            if prev_y is not None:
                gap = y - prev_y
                if gap > 0:
                    neo.append(f"[行間]{gap:.2f}\n")

            if el["type"] == "text":
                text = el["content"]
                font = fs_font_override or "IPAexGothic, sans-serif"
                size = 12.0 + fs_size_add
                neo.append(f"[フォント:{font}][サイズ:{size:.2f}][ウェイト:normal]{text}\n")
                sorted_txt.append(f"テキスト: {text}\n")
                prev_y = el["bbox"][3]
            elif el["type"] == "image":
                bbox = el["bbox"]
                neo.append(f"[画像:{el['content']}:{bbox[0]:.2f}:{bbox[1]:.2f}:{bbox[2]-bbox[0]:.2f}:{bbox[3]-bbox[1]:.2f}]\n")
                sorted_txt.append(f"[画像] {el['content']} | BBOX: {bbox}\n\n")
                prev_y = bbox[3]

    neo_content = "".join(neo)
    sorted_content = "".join(sorted_txt)

    # ファイル保存
    with open(output_file_NEO, "w", encoding="utf-8") as f:
        f.write(neo_content)
    with open(output_file_SORTED, "w", encoding="utf-8") as f:
        f.write(sorted_content)

    # PDF生成
    recreated_pdf_filename = f"{basename}_recreated.pdf"
    recreated_pdf_path = os.path.join(dir_name, recreated_pdf_filename)
    pdf_ok, _ = create_pdf_with_weasyprint(neo_content, recreated_pdf_path, app_root)
    recreated_pdf_url = os.path.join(basename, recreated_pdf_filename).replace("\\", "/") if pdf_ok else ""

    # HTML生成用
    styled_neo_html = create_styled_html(neo_content, app_root)
    og = open(output_file_OG, encoding="utf-8").read()

    image_gallery_html = "".join(
        f'<a href="/outputs/{html.escape(url)}" target="_blank">'
        f'<img src="/outputs/{html.escape(url)}" alt="image"></a>'
        for url in imgs
    ) or "<p>画像は抽出されませんでした。</p>"

    download_html = (
        f'<div class="download-section"><h3>再構成されたPDF</h3>'
        f'<a href="/outputs/{html.escape(recreated_pdf_url)}" class="action-link" download>ダウンロード</a></div>'
        if pdf_ok else ""
    )

    # --- 見た目統合版HTML出力 ---
    result_html = f"""
    <!doctype html>
    <html lang="ja">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>処理結果</title>
            <style>
                body {{ font-family: 'Helvetica Neue', Arial, sans-serif; line-height: 1.6; margin: 0; background-color: #f4f4f9; }}
                .container {{ max-width: 960px; margin: 2em auto; padding: 2em; background: white; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }}
                h2 {{ border-bottom: 2px solid #007bff; padding-bottom: 10px; color: #333; }}
                .content-box {{ border: 1px solid #ddd; background-color: #fdfdfd; padding: 1em; margin-top: 1em; white-space: pre-wrap; word-wrap: break-word; max-height: 400px; overflow-y: auto; font-family: 'Courier New', monospace; font-size: 14px; }}
                .styled-content-box {{ border: 1px solid #ddd; background-color: #fdfdfd; padding: 1em; margin-top: 1em; max-height: 400px; overflow-y: auto; }}
                details {{ border: 1px solid #ccc; border-radius: 5px; padding: 0.5em; margin-bottom: 1em; background-color: #f9f9f9; }}
                summary {{ font-weight: bold; cursor: pointer; padding: 0.5em; font-size: 1.1em; color: #0056b3; }}
                .info, .download-section {{ background: #eef; padding: 1em; border-radius: 8px; margin-bottom: 1.5em; }}
                .image-gallery {{ display: flex; flex-wrap: wrap; gap: 15px; padding: 1em; }}
                .image-gallery img {{ border: 2px solid #ddd; border-radius: 5px; padding: 5px; max-width: 150px; height: auto; cursor: pointer; transition: transform 0.2s; }}
                .image-gallery img:hover {{ transform: scale(1.05); border-color: #007bff; }}
                .action-link {{ display: inline-block; margin-top: 1em; background-color: #28a745; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; font-weight: bold; }}
                .action-link:hover {{ background-color: #218838; }}
                .back-link {{ background-color: #6c757d; }}
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
                <details><summary>スタイル付き NEOテキスト</summary><div class="styled-content-box">{styled_neo_html}</div></details>
                <details><summary>NEOテキスト (タグ付き)</summary><div class="content-box">{html.escape(neo_content)}</div></details>
                <details><summary>OGテキスト</summary><div class="content-box">{html.escape(og)}</div></details>
                <details><summary>時系列ソート</summary><div class="content-box">{html.escape(sorted_content)}</div></details>
                <details open><summary>抽出画像 ({len(imgs)}枚)</summary><div class="image-gallery">{image_gallery_html}</div></details>
                <a href="/" class="action-link back-link">別のファイルを処理する</a>
            </div>
        </body>
    </html>
    """
    return result_html


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 3000))
    app.run(debug=False, host="0.0.0.0", port=port)