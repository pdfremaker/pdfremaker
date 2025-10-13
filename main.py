"""
replitで変更したデータをGitHubに反映させるときは次のコードをShellにコピペ

git add .
git commit -m "update"
git push

↑git commit -m ""の""の中に更新内容を書く
別にupdateのままでもおけ丸水産
"""

# Import the necessary modules
from flask import Flask, request, render_template_string, jsonify, send_file
import html
import json
import os
import tempfile
import re
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
    # 教科書体に近い読みやすさ（明朝系）
    "Noto Serif JP": "fonts/NotoSerifJP-Regular.ttf",
    "明朝体, serif": "fonts/NotoSerifJP-Regular.ttf",
    "IPAex明朝": "fonts/ipaexg.ttf",  # fallback（代用：明朝もIPAexに）

    # 読みやすいゴシック（標準・ベース）
    "Noto Sans JP": "fonts/NotoSansJP-Regular.ttf",
    "ゴシック体, sans-serif": "fonts/NotoSansJP-Regular.ttf",
    "IPAexゴシック": "fonts/ipaexg.ttf",  # fallback

    # 優しい丸ゴシック（読み障がい支援向け）
    "Kosugi Maru": "fonts/KosugiMaru-Regular.ttf",

    # 英字・軽量フォントの代替
    "Verdana, sans-serif": "fonts/NotoSansJP-Regular.ttf",
    "Arial, sans-serif": "fonts/NotoSansJP-Regular.ttf"
}


def get_font_path(app_root, font_family_name="IPAexGothic"):
    font_file = FONT_FILE_MAP.get(font_family_name, "ipaexg.ttf")
    if not os.path.isabs(font_file):
        font_path = os.path.join(app_root, font_file)
    else:
        font_path = font_file

    font_path = os.path.abspath(font_path)
    if not os.path.exists(font_path):
        # ここでフォールバック先を fonts フォルダに限定
        fallback_path = os.path.join(app_root, "fonts", "ipaexg.ttf")
        if os.path.exists(fallback_path):
            print(f"✅ フォントファイルが見つかりました: {fallback_path}")
            return fallback_path
        else:
            print(f"⚠️ フォントファイルが存在しません: {fallback_path}")
            return None
    return font_path


# 関数を呼び出す
font_path = get_font_path(app_root, "IPAexGothic")
font_url = path2url(font_path)

# Firebaseを初期化
service_key_json = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS_JSON")
if service_key_json:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".json", mode="w") as temp_file:
        temp_file.write(service_key_json)
        temp_file_path = temp_file.name
    cred = credentials.Certificate(temp_file_path)
    firebase_admin.initialize_app(cred)
else:
    cred = credentials.Certificate("serviceAccountKey.json")
    firebase_admin.initialize_app(cred)

db = firestore.client()
config_ref = db.collection("messages")


def get_firestore_config(user_id="default_user"):
    doc = config_ref.document(user_id).get()
    if doc.exists:
        return doc.to_dict()
    else:
        default_config = {
            "fontSize": 16,
            "lineHeight": 1.6,
            "fontSelect": "Kosugi Maru"
        }
        config_ref.document(user_id).set(default_config)
        return default_config


def get_document(collection_name, doc_id):
    if db is None:
        print("⚠️ Firestoreが初期化されていません")
        return None
    try:
        doc_ref = db.collection(collection_name).document(doc_id)
        docf = doc_ref.get()
        return docf.to_dict() if docf.exists else None
    except Exception as e:
        print(f"⚠️ Firestoreアクセス失敗: {e}")
        return None


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
                <a href="/edit" class="button-link">
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
            const id = document.getElementById("student-id").value.trim();
            if (!id) {
                alert("生徒IDを入力してください。");
                return;
            }

            try {
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
            } catch (e) {
                alert("通信エラー: " + e.message);
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


HTML_EDIT = """
<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Firestore 設定ページ</title>
  <style>
    body { font-family: 'Helvetica Neue', Arial, sans-serif; margin: 40px; background-color: #f4f4f9; color: #333; }
    .container { max-width: 600px; margin: auto; text-align: center; }
    h1 { color: #5a5a5a; }
    p { color: #666; }
    form { background: white; padding: 2em; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-top: 1em; }
    input, select, button { margin: 10px; font-size: 16px; }
    input[type=text], input[type=number] {
      width: 95%; padding: 10px; border-radius: 5px; border: 1px solid #ccc;
    }
    button, input[type=button], input[type=submit] {
      background-color: #007bff; color: white; padding: 10px 20px;
      border: none; border-radius: 5px; cursor: pointer; font-size: 16px; width: 100%;
      transition: background-color 0.3s ease;
    }
    button:hover { background-color: #0056b3; }
    .button-link {
      text-decoration: none; background-color:#228b22; color: white;
      padding: 6px 6px; border-radius: 5px; display: inline-block; transition: background-color 0.3s; margin-right: 10px;
    }
    .button-link:hover { background-color: #333333; }
    #text { font-size: 40px; line-height: 40px; margin-top: 20px; }
    #message { margin-top: 20px; font-size: 20px; display: none; white-space: pre-line; }
  </style>
</head>
<body>
  <div class="container">
    <h1>Firestore 生徒設定エディタ</h1>
    <p>生徒のフォント・文字サイズ・行間などを設定します。</p>

    <div class="align-right-container">
      <a href="/" class="button-link"><b>再構成画面へ戻る</b></a>
    </div>

    <form>
      <h2>行間スライダー</h2>
      <label for="lineSlider">行間: <span id="lineValue">1.0</span></label><br>
      <input type="range" id="lineSlider" min="1" max="2" step="0.1" value="1">

      <h2>文字サイズスライダー</h2>
      <label for="fontSizeSlider">文字サイズ: <span id="sizeValue">40</span>px</label><br>
      <input type="range" id="fontSizeSlider" min="20" max="60" value="40">

      <h2>フォントを選択</h2>
      <select id="fontSelect">
        <option value="Noto Sans JP">ゴシック</option>
        <option value="Noto Serif JP">明朝</option>
        <option value="Kosugi Maru">丸ゴシック</option>
        <option value="IPAexゴシック">IPAexゴシック</option>
        <option value="IPAex明朝">IPAex明朝</option>
        <option value="Verdana, sans-serif">Verdana（代替英字フォント）</option>
      </select>

      <div id="text">
        こんにちは<br>おはよう
      </div>

      <h2>設定確認</h2>
      <button type="button" id="showBtn">確認</button>
      <div id="message">ここに設定内容が表示されます</div>

      <h2>Firestore 登録</h2>
      <input type="text" id="docId" placeholder="IDを入力">
      <input type="text" id="name" placeholder="名前を入力">
      <input type="number" id="number" placeholder="出席番号を入力">
      <br>
      <button type="button" id="createBtn">Firestoreに登録</button>
    </form>
  </div>

  <!-- JavaScript (Firebase SDK不要: Flask経由API呼び出し想定) -->
  <script>
    const baseLine = 40;
    const maxDiff = 20;

    const text = document.getElementById("text");
    const lineSlider = document.getElementById("lineSlider");
    const fontSizeSlider = document.getElementById("fontSizeSlider");
    const fontSelect = document.getElementById("fontSelect");
    const lineValue = document.getElementById("lineValue");
    const sizeValue = document.getElementById("sizeValue");
    const message = document.getElementById("message");

    function updateStyle() {
      const sliderValue = Number(lineSlider.value);
      const lineHeight = baseLine + sliderValue * maxDiff;
      lineValue.textContent = sliderValue.toFixed(1);
      sizeValue.textContent = fontSizeSlider.value;
      text.style.lineHeight = lineHeight + "px";
      text.style.fontSize = fontSizeSlider.value + "px";
      text.style.fontFamily = fontSelect.value;
    }

    lineSlider.addEventListener("input", updateStyle);
    fontSizeSlider.addEventListener("input", updateStyle);
    fontSelect.addEventListener("change", updateStyle);
    updateStyle();

    document.getElementById("showBtn").addEventListener("click", () => {
      const sliderValue = Number(lineSlider.value);
      const fontSize = fontSizeSlider.value;
      const fontFamily = fontSelect.value;
      message.textContent = `フォント: ${fontFamily}\\n文字サイズ: ${fontSize}px\\n行間倍率: ${sliderValue.toFixed(1)}`;
      message.style.display = "block";
    });

    document.getElementById("createBtn").addEventListener("click", async () => {
      const docId = document.getElementById("docId").value.trim();
      const name = document.getElementById("name").value.trim();
      const number = Number(document.getElementById("number").value);
      const lineHeight = Number(lineSlider.value);
      const fontSize = Number(fontSizeSlider.value);
      const fontSelectValue = fontSelect.value;

      if (!docId || !name || isNaN(number)) {
        alert("ID・名前・番号をすべて入力してください。");
        return;
      }

      try {
        const res = await fetch("/update_firestore", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ id: docId, name, number, lineHeight, fontSize, fontSelect: fontSelectValue })
        });
        const data = await res.json();
        alert(data.message);
      } catch (e) {
        alert("通信エラー: " + e.message);
      }
    });
  </script>
</body>
</html>
"""


# Firestoreの情報変える生徒側へ行く
@app.route('/edit')
def edit_page():
    return HTML_EDIT


@app.route("/update_firestore", methods=["POST"])
def update_firestore():
    data = request.json
    user_id = data.get("user_id", "default_user")
    update_data = {
        "fontSize": data.get("fontSize", 16),
        "lineHeight": data.get("lineHeight", 1.6),
        "fontSelect": data.get("fontSelect", "Kosugi Maru"),
    }
    config_ref.document(user_id).set(update_data)
    return jsonify({"status": "ok", "updated": update_data})


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
                app.logger.info(
                    f"ID '{student_id}' の設定を適用します: {firebase_settings}")
            else:
                app.logger.info(
                    f"ID '{student_id}' は見つかりませんでした。デフォルト設定で処理します。")

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


def convert_neo_to_html(neo_content: str, font_size=16, line_height=1.6, font_select="IPAexGothic", app_root=".") -> str:
    """
    NEOタグ形式テキストをHTMLへ変換し、フォント・行間・サイズを反映する
    """
    import re, html

    html_lines = []
    current_font = font_select
    current_size = font_size
    current_weight = "normal"
    current_line_height = line_height

    # 各行を解析
    for line in neo_content.splitlines():
        line = line.strip()
        if not line:
            continue

        # フォント指定
        if line.startswith("[フォント:"):
            font_match = re.search(r"\[フォント:(.*?)\]", line)
            size_match = re.search(r"\[サイズ:(.*?)\]", line)
            weight_match = re.search(r"\[ウェイト:(.*?)\]", line)
            text_match = re.search(r"\](.+)", line)

            if font_match:
                current_font = font_match.group(1).strip()
            if size_match:
                try:
                    current_size = float(size_match.group(1).strip())
                except:
                    pass
            if weight_match:
                current_weight = weight_match.group(1).strip()

            text_content = text_match.group(1).strip() if text_match else ""
            html_lines.append(
                f'<p style="font-family:{current_font}; font-size:{current_size}px; font-weight:{current_weight}; line-height:{current_line_height};">'
                f'{html.escape(text_content)}</p>'
            )

        # 行間設定
        elif line.startswith("[行間]"):
            try:
                current_line_height = float(line.replace("[行間]", "").strip())
            except:
                pass

        # 画像挿入
        elif line.startswith("[画像:"):
            img_match = re.match(r"\[画像:(.*?):([\d\.]+):([\d\.]+):([\d\.]+):([\d\.]+)\]", line)
            if img_match:
                img_path = img_match.group(1)
                img_rel_path = img_path.replace(app_root, "").replace("/home/runner/workspace", "").lstrip("/")
                img_width = img_match.group(4)
                img_height = img_match.group(5)
                html_lines.append(
                    f'<img src="/{img_rel_path}" style="width:{img_width}px; height:{img_height}px; display:block; margin:8px auto;">'
                )

        # 通常テキスト
        else:
            html_lines.append(
                f'<p style="font-family:{current_font}; font-size:{current_size}px; font-weight:{current_weight}; line-height:{current_line_height};">'
                f'{html.escape(line)}</p>'
            )

    # HTML全体
    html_output = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{
                font-family: '{font_select}';
                font-size: {font_size}px;
                line-height: {line_height};
                color: #111;
                background: #fff;
                margin: 24px;
                padding: 0;
            }}
            img {{
                max-width: 90%;
                border-radius: 8px;
            }}
        </style>
    </head>
    <body>
        {''.join(html_lines)}
    </body>
    </html>
    """

    return html_output


def create_pdf_with_weasyprint(neo_content, output_path, app_root, firebase_settings=None):
    """
    neo_content を解析して HTML を作り、必要なフォントをすべて @font-face で定義して
    WeasyPrint に渡して PDF を生成する（画像は file:// 経由で埋め込み）。
    """
    print("=== NEO解析内容 (先頭800文字) ===")
    print(neo_content[:800])
    from weasyprint import HTML, CSS
    import re, os, html as pyhtml

    try:
        # --- 1) 使われているフォント名を収集 ---
        font_names = set(re.findall(r'\[フォント:(.*?)\]', neo_content))
        # デフォルトフォントも入れておく
        if firebase_settings and firebase_settings.get("fontSelect"):
            font_names.add(firebase_settings.get("fontSelect"))
        if not font_names:
            font_names.add("IPAexGothic")

        # --- 2) 各フォント名をファイルパスに解決して @font-face を作る ---
        font_face_rules = []
        for fname in sorted(font_names):
            # get_font_path は既に定義されている関数を使う
            path = get_font_path(app_root, fname)
            if not path:
                # フォントが見つからなければ ipaex を fallback として使う
                path = get_font_path(app_root, "IPAex明朝") or get_font_path(app_root, "IPAexゴシック")
            if path:
                # file:// フルパスで指定
                font_face_rules.append(
                    f"@font-face {{ font-family: '{fname}'; src: url('file://{path}'); }}"
                )
            else:
                print(f"⚠️ フォントファイル見つからず: {fname}")

        # --- 3) HTML ブロックを作る ---
        html_blocks = []
        current_font = None
        current_size = None
        current_weight = None
        current_lineheight = None

        for raw_line in neo_content.splitlines():
            line = raw_line.strip()
            if not line:
                continue

            # 行間はここでは無視（必要なら current_lineheight を取り込む）
            if line.startswith("[行間]"):
                # 任意処理：行間を CSS 単位に変換したい場合はここで current_lineheight に格納
                try:
                    current_lineheight = float(line.replace("[行間]", "").strip())
                except:
                    current_lineheight = None
                continue

            # 画像タグ
            if line.startswith("[画像:"):
                parts = re.findall(r"\[画像:(.*?):([\d\.]+):([\d\.]+):([\d\.]+):([\d\.]+)\]", line)
                if parts:
                    img_path, x, y, w, h = parts[0]
                    # 画像はローカルファイル経由で埋め込む（WeasyPrint が file:// をサポート）
                    img_file_url = f"file://{os.path.abspath(img_path)}"
                    html_blocks.append(f'<div style="text-align:center; margin: 1em 0;"><img src="{img_file_url}" style="max-width:90%;"></div>')
                continue

            # フォント/サイズ/ウェイトタグを探す
            font_match = re.search(r"\[フォント:(.*?)\]", line)
            size_match = re.search(r"\[サイズ:(.*?)\]", line)
            weight_match = re.search(r"\[ウェイト:(.*?)\]", line)

            text = re.sub(r"\[.*?\]", "", line).strip()
            if not text:
                continue

            # 決定したフォント情報を使って p タグを作る
            used_font = font_match.group(1).strip() if font_match else (firebase_settings.get("fontSelect") if firebase_settings else "IPAexGothic")
            used_size = size_match.group(1).strip() if size_match else (str(firebase_settings.get("fontSize")) if firebase_settings else "16")
            used_weight = weight_match.group(1).strip() if weight_match else "normal"

            # line-height の反映（もし current_lineheight があれば）
            lh_css = "line-height:1.6;"
            if current_lineheight:
                # neo の行間が px ベースだったら相当に大きくなるので簡易変換
                try:
                    # 小〜中程度の値に落とす（必要に応じて調整）
                    lh_val = max(1.0, float(current_lineheight) / 20.0)
                    lh_css = f"line-height:{lh_val};"
                except:
                    pass

            # escape
            esc_text = pyhtml.escape(text)
            html_blocks.append(
                f"<p style=\"font-family:'{used_font}'; font-size:{used_size}px; font-weight:{used_weight}; {lh_css} margin:0.3em 0;\">{esc_text}</p>"
            )

        body_html = "\n".join(html_blocks)

        # --- 4) 最終 HTML テンプレート（フォント定義を head に埋め込む） ---
        css_font_defs = "\n".join(font_face_rules)
        html_template = f"""
        <html lang="ja">
        <head>
            <meta charset="utf-8">
            <style>
                {css_font_defs}
                body {{
                    padding: 1cm;
                    word-wrap: break-word;
                    background: white;
                }}
                img {{ page-break-inside: avoid; max-width:100%; }}
            </style>
        </head>
        <body>
            {body_html}
        </body>
        </html>
        """

        # --- 5) WeasyPrint に書かせる ---
        # base_url は app_root にしておく（ファイル参照の解決に使われる）
        HTML(string=html_template, base_url=app_root).write_pdf(output_path)

        print(f"✅ PDF生成成功: {output_path}")
        return True, None

    except Exception as e:
        print("❌ PDF生成失敗:", e)
        return False, str(e)


def process_pdf(pdf_path: str, firebase_settings: dict | None = None):
    try:
        doc = fitz.open(pdf_path)
        assert isinstance(doc, fitz.Document)
    except Exception as e:
        return f"PDFを開けません: {e}"

    basename = os.path.splitext(os.path.basename(pdf_path))[0]
    dir_name = os.path.join(OUTPUT_FOLDER, basename)
    os.makedirs(dir_name, exist_ok=True)

    # --- Firebase設定を取得 ---
    fs_font_override = firebase_settings.get(
        "fontSelect") if firebase_settings else None
    fs_size_add = float(firebase_settings.get("fontSize",
                                              0)) if firebase_settings else 0.0

    # 出力ファイルパス
    output_file_OG = os.path.join(dir_name, f"{basename}_OG.txt")
    output_file_NEO = os.path.join(dir_name, f"{basename}_NEO.txt")
    output_file_SORTED = os.path.join(dir_name, f"{basename}_SORTED.txt")

    neo, sorted_txt, imgs, og_tagged = [], [], [], []

    # --- ページごとの抽出 ---
    for i, page in enumerate(doc):
        sorted_txt.append(f"\n--- Page {i+1} ---\n")
        elements = []

        # テキスト抽出
        for blk in page.get_text("dict")["blocks"]:
            if blk["type"] == 0:
                text = "".join(span["text"] for ln in blk["lines"]
                               for span in ln["spans"]).strip()
                if text:
                    elements.append({
                        "type": "text",
                        "bbox": blk["bbox"],
                        "content": text
                    })

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
                elements.append({
                    "type": "image",
                    "bbox": bbox,
                    "content": full
                })
            except Exception as e:
                print("画像抽出失敗:", e)

        # 座標順ソート
        elements.sort(key=lambda x: (x["bbox"][1], x["bbox"][0]))

        prev_y = None
        for el in elements:
            y = el["bbox"][1]

            # --- 🔹 行間処理 ---
            if prev_y is not None:
                gap = y - prev_y
                if gap > 0:
                    line_gap = gap
                    # Firestoreの倍率反映（NEO用）
                    if firebase_settings and firebase_settings.get(
                            "lineHeight"):
                        try:
                            multiplier = float(firebase_settings["lineHeight"])
                            line_gap = gap * multiplier
                        except Exception:
                            pass
                    # それぞれに反映
                    neo.append(f"[行間]{line_gap:.2f}\n")  # 生徒設定適用後
                    og_tagged.append(f"[行間]{gap:.2f}\n")  # 元PDF値

            # --- テキスト要素 ---
            if el["type"] == "text":
                text = el["content"]

                # --- 元PDFフォント情報を取得 (OG用)
                try:
                    found_span = None
                    for blk in page.get_text("dict")["blocks"]:
                        if blk["type"] == 0:
                            for line in blk["lines"]:
                                for span in line["spans"]:
                                    if span["text"].strip(
                                    ) and span["text"].strip() in text:
                                        found_span = span
                                        break
                                if found_span:
                                    break
                        if found_span:
                            break

                    if found_span:
                        og_font = found_span.get("font", "Unknown")
                        og_size = found_span.get("size", 12.0)
                        og_weight = "bold" if "Bold" in og_font else "normal"
                    else:
                        og_font, og_size, og_weight = "Unknown", 12.0, "normal"

                except Exception:
                    og_font, og_size, og_weight = "Unknown", 12.0, "normal"

                # --- Firestore設定反映後のフォント (NEO用)
                font = fs_font_override or "IPAexGothic, sans-serif"
                size = og_size + fs_size_add  # 元サイズに加算

                # --- 出力
                neo.append(
                    f"[フォント:{font}][サイズ:{size:.2f}][ウェイト:normal]{text}\n")
                og_tagged.append(
                    f"[フォント:{og_font}][サイズ:{og_size:.2f}][ウェイト:{og_weight}]{text}\n"
                )
                sorted_txt.append(f"テキスト: {text}\n")

                prev_y = el["bbox"][3]

            # --- 画像要素 ---
            elif el["type"] == "image":
                bbox = el["bbox"]
                img_tag = f"[画像:{el['content']}:{bbox[0]:.2f}:{bbox[1]:.2f}:{bbox[2]-bbox[0]:.2f}:{bbox[3]-bbox[1]:.2f}]\n"
                neo.append(img_tag)
                og_tagged.append(img_tag)
                sorted_txt.append(f"[画像] {el['content']} | BBOX: {bbox}\n\n")
                prev_y = bbox[3]

    # --- 出力内容を結合 ---
    neo_content = "".join(neo)
    og_tagged_content = "".join(og_tagged)
    sorted_content = "".join(sorted_txt)

    # --- ファイル保存 ---
    with open(output_file_NEO, "w", encoding="utf-8") as f:
        f.write(neo_content)
    with open(output_file_SORTED, "w", encoding="utf-8") as f:
        f.write(sorted_content)
    with open(output_file_OG, "w", encoding="utf-8") as f:
        f.write(og_tagged_content)

    # --- PDF再構築 ---
    recreated_pdf_filename = f"{basename}_recreated.pdf"
    recreated_pdf_path = os.path.join(dir_name, recreated_pdf_filename)
    pdf_ok, pdf_error = create_pdf_with_weasyprint(
        neo_content,
        recreated_pdf_path,
        app_root,
        firebase_settings=firebase_settings
    )
    recreated_pdf_url = os.path.join(basename, recreated_pdf_filename).replace(
        "\\", "/") if pdf_ok else ""
    if not pdf_ok:
        print("❌ PDF再構成に失敗:", pdf_error)
        recreated_pdf_url = ""
        download_html = "<p style='color:red;'>PDFの再構成に失敗しました。</p>"
    else:
        print("✅ PDF再構成成功:", recreated_pdf_path)
        recreated_pdf_url = os.path.join(basename, recreated_pdf_filename).replace("\\", "/")
        download_html = (
            f'<div class="download-section"><h3>再構成されたPDF</h3>'
            f'<a href="/outputs/{html.escape(recreated_pdf_url)}" '
            f'class="action-link" download>ダウンロード</a></div>'
        )

    # --- NEOテキスト生成（追加） ---
    extracted_text = "".join(neo)
    neo_text = extracted_text

    font_size = firebase_settings.get("fontSize", 16) if firebase_settings else 16
    line_height = firebase_settings.get("lineHeight", 1.6) if firebase_settings else 1.6
    font_select = firebase_settings.get("fontSelect", "IPAexGothic") if firebase_settings else "IPAexGothic"

    # --- HTML生成 ---
    styled_neo_html = convert_neo_to_html(
        neo_text,
        font_size,
        line_height,
        font_select,
        app_root
    )
    
    image_gallery_html = "".join(
        f'<a href="/outputs/{html.escape(url)}" target="_blank">'
        f'<img src="/outputs/{html.escape(url)}" alt="image"></a>'
        for url in imgs) or "<p>画像は抽出されませんでした。</p>"

    download_html = (
        f'<div class="download-section"><h3>再構成されたPDF</h3>'
        f'<a href="/outputs/{html.escape(recreated_pdf_url)}" class="action-link" download>ダウンロード</a></div>'
        if pdf_ok else "")

    # --- 統合HTML ---
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
                <details><summary>OGテキスト (タグ付き)</summary><div class="content-box">{html.escape(og_tagged_content)}</div></details>
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