"""
replitで変更したデータをGitHubに反映させるときは次のコードをShellにコピペ

git add .
git commit -m "update: "
git push

↑git commit -m "update"の中に更新内容を書く 別にupdateのままでもおけ丸水産
"""

# Flask関連
from flask import Flask, request, jsonify, send_file, render_template
from werkzeug.utils import secure_filename

# 標準ライブラリ
import os
import re
import tempfile
import html
import html as pyhtml
import json
from datetime import datetime, timedelta, timezone
import shutil
from bs4 import BeautifulSoup
import time
import mimetypes

# PDF操作関連
import pymupdf as fitz
from weasyprint import HTML, CSS
from weasyprint.urls import path2url

# フォント関連
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

# Firebase関連
import firebase_admin
from firebase_admin import credentials, firestore

# デバッグ・ログ関連
import logging
from logging.handlers import RotatingFileHandler


# ログ設定
def setup_logging():
    """
    ログ設定（日本標準時対応・Flask互換）
    - JSTで日付フォルダを自動作成（例: logs/2025-10-14）
    - app.log（INFO以上） / error.log（WARNING以上）を自動分離
    - 2MB×5世代ローテーション
    - Flaskや他ライブラリの初期化済logging設定を上書き
    """

    # 日本標準時（JST）
    JST = timezone(timedelta(hours=9), name="Asia/Tokyo")
    logging.Formatter.converter = lambda *args: datetime.now(JST).timetuple()

    # 日付フォルダ
    today_str = datetime.now(JST).strftime("%Y-%m-%d")
    log_dir = os.path.join("logs", today_str)
    os.makedirs(log_dir, exist_ok=True)

    # 設定
    max_bytes = 2_000_000  # 2MB
    backup_count = 7
    log_format = "%(asctime)s [%(levelname)s] %(name)s - %(message)s"
    formatter = logging.Formatter(log_format)

    # INFO以上: app.log
    app_handler = RotatingFileHandler(
        os.path.join(log_dir, "app.log"),
        maxBytes=max_bytes,
        backupCount=backup_count,
        encoding="utf-8"
    )
    app_handler.setLevel(logging.INFO)
    app_handler.setFormatter(formatter)

    # WARNING以上: error.log
    error_handler = RotatingFileHandler(
        os.path.join(log_dir, "error.log"),
        maxBytes=max_bytes,
        backupCount=backup_count,
        encoding="utf-8"
    )
    error_handler.setLevel(logging.WARNING)
    error_handler.setFormatter(formatter)

    # コンソール出力
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)

    # 既存ハンドラを削除して再構成
    root_logger = logging.getLogger()
    for handler in root_logger.handlers[:]:
        root_logger.removeHandler(handler)

    root_logger.setLevel(logging.INFO)
    root_logger.addHandler(app_handler)
    root_logger.addHandler(error_handler)
    root_logger.addHandler(console_handler)

    # 動作確認用ログ
    logger = logging.getLogger("pdf_remaker")
    logger.info("✅ ログ初期化完了")
    logger.info(f"✅ 日付（JST）: {today_str}")
    logger.info(f"✅ ログディレクトリ: {log_dir}")
    logger.info("✅ app.log / error.log 分離・ローテーション有効")

    return logger


def cleanup_old_logs(base_dir: str, days_to_keep: int, logger_obj):
    """
    base_dir 内の YYYY-MM-DD フォルダをチェックし、
    days_to_keep 日より古いフォルダを削除。
    """
    if not os.path.isdir(base_dir):
        logger_obj.info("cleanup_old_logs: no logs directory found (%s)", base_dir)
        return

    now = datetime.now()
    cutoff = now - timedelta(days=days_to_keep)

    for folder_name in os.listdir(base_dir):
        folder_path = os.path.join(base_dir, folder_name)
        if not os.path.isdir(folder_path):
            continue

        try:
            folder_date = datetime.strptime(folder_name, "%Y-%m-%d")
        except ValueError:
            # 日付形式でないフォルダは無視
            continue

        if folder_date < cutoff:
            try:
                shutil.rmtree(folder_path)
                logger_obj.info("🧹 cleanup_old_logs: removed old log folder %s", folder_path)
            except Exception as e:
                logger_obj.exception("cleanup_old_logs: failed to remove %s: %s", folder_path, e)


# ログ初期化
logger = setup_logging()

# 古いログを自動削除
days_to_keep = int(os.environ.get("LOG_DAYS_TO_KEEP", "7"))  # 7日保持
cleanup_old_logs("logs", days_to_keep, logger)

# Flask・環境設定
print("(;^ω^) 起動中static.")
print(f"DEBUG: fitz module path: {fitz.__file__}")
print(f"DEBUG: fitz.open available: {hasattr(fitz, 'open')}")

app_root = os.path.dirname(os.path.abspath(__file__))

# フォント設定
FONT_FILE_MAP = {
    "Noto Serif JP": "static/fonts/NotoSerifJP-Regular.ttf",
    "明朝体, serif": "static/fonts/NotoSerifJP-Regular.ttf",
    "IPAex明朝": "static/fonts/ipaexg.ttf",
    "Noto Sans JP": "static/fonts/NotoSansJP-Regular.ttf",
    "ゴシック体, sans-serif": "static/fonts/NotoSansJP-Regular.ttf",
    "IPAexゴシック": "static/fonts/ipaexg.ttf",
    "Kosugi Maru": "static/fonts/KosugiMaru-Regular.ttf",
    "Verdana, sans-serif": "static/fonts/NotoSansJP-Regular.ttf",
    "Arial, sans-serif": "static/fonts/NotoSansJP-Regular.ttf"
}


def get_font_path(app_root, font_family_name="IPAexGothic"):
    font_file = FONT_FILE_MAP.get(font_family_name, "ipaexg.ttf")
    if not os.path.isabs(font_file):
        font_path = os.path.join(app_root, font_file)
    else:
        font_path = font_file

    font_path = os.path.abspath(font_path)
    if not os.path.exists(font_path):
        fallback_path = os.path.join(app_root, "static/fonts", "ipaexg.ttf")
        if os.path.exists(fallback_path):
            logger.info(f"✅ フォントファイルが見つかりました: {fallback_path}")
            return fallback_path
        else:
            logger.warning(f"⚠️ フォントファイルが存在しません: {fallback_path}")
            return None
    return font_path


font_path = get_font_path(app_root, "IPAexGothic")
font_url = path2url(font_path) if font_path else None

# Firebase 初期化
try:
    service_key_json = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS_JSON")
    if service_key_json:
        # Render等のサーバ環境
        with tempfile.NamedTemporaryFile(delete=False,
                                         suffix=".json",
                                         mode="w") as temp_file:
            temp_file.write(service_key_json)
            temp_file_path = temp_file.name
        cred = credentials.Certificate(temp_file_path)
        firebase_admin.initialize_app(cred)
        logger.info("✅ Firebase初期化: 環境変数から読み込み成功")
    else:
        # ローカル環境
        cred = credentials.Certificate("serAccoCaMnNeMg.json")
        firebase_admin.initialize_app(cred)
        logger.info("✅ Firebase初期化: serAccoCaMnNeMg.jsonから読み込み成功")

    db = firestore.client()
    config_ref = db.collection("messages")
    logger.info("✅ Firestore接続成功")

except Exception as e:
    logger.critical(f"Firebase初期化エラー: {e}", exc_info=True)
    raise SystemExit("Firebase初期化に失敗しました。")


def get_firestore_config(user_id="default_user"):
    logger.info("get_firestore_config: loading config for user_id=%s", user_id)
    try:
        doc = config_ref.document(user_id).get()
        if doc.exists:
            data = doc.to_dict()
            logger.debug("get_firestore_config: found document %s -> %s",
                         user_id, data)
            return data
        else:
            # Firestoreにまだ設定がない場合、デフォルトを作成
            default_config = {
                "fontSize": 16,
                "lineHeight": 1.6,
                "fontSelect": "Kosugi Maru"
            }
            config_ref.document(user_id).set(default_config)
            logger.info(
                "get_firestore_config: created default config for new user_id=%s",
                user_id)
            return default_config
    except Exception as e:
        logger.exception(
            "get_firestore_config: Firestore access failed for user_id=%s",
            user_id)
        # エラー時には安全なデフォルトを返す
        return {"fontSize": 16, "lineHeight": 1.6, "fontSelect": "Kosugi Maru"}


def get_document(collection_name, doc_id):
    try:
        logger.info(
            f"get_document: loading document '{doc_id}' from collection '{collection_name}'"
        )
        doc_ref = db.collection(collection_name).document(doc_id)
        doc = doc_ref.get()
        if doc.exists:
            return doc.to_dict()
        else:
            logger.warning(f"get_document: document '{doc_id}' not found.")
            return None
    except Exception as e:
        logger.exception("Firestoreアクセス中にエラーが発生しました")
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


# 戻る
@app.route('/return')
def return_page():
    return render_template("upload.html", page_name="upload")


# Firestoreの情報変える生徒側へ行く
@app.route('/edit')
def edit_page():
    return render_template("edit.html", page_name="edit")


@app.route("/update_firestore", methods=["POST"])
def update_firestore():
    try:
        data = request.get_json()
        doc_id = data.get("id")
        if not doc_id:
            return jsonify({"message": "IDが指定されていません。"}), 400

        db.collection("messages").document(doc_id).set(data)
        logger.info(f"Firestore updated for id={doc_id}")
        return jsonify({"message": f"{doc_id} の設定を登録しました！"})

    except Exception:
        logger.exception("Firestore更新中にエラーが発生しました")
        return jsonify({"message": "Firestore更新中に内部エラーが発生しました。"}), 500


# Firestoreのメッセージ取得
@app.route("/get_message", methods=["GET"])
def get_message_api():
    doc_id = request.args.get("id", "").strip()
    logger.info("get_message called for id=%s", doc_id)

    if not doc_id:
        logger.warning("get_message: no id provided")
        return jsonify({"error": "IDが指定されていません"}), 400

    data = get_document("messages", doc_id)
    if not data:
        logger.info("get_message: id not found: %s", doc_id)
        return jsonify({"error": f"ID '{doc_id}' は存在しません"}), 404

    logger.info("get_message: found config for id=%s", doc_id)
    return jsonify({
        k: data.get(k, "N/A")
        for k in ["fontSelect", "fontSize", "lineHeight"]
    } | {"id": doc_id})


@app.route("/", methods=["GET", "POST"])
def upload_pdf():
    if request.method != "POST":
        logger.debug("upload_pdf: GET request — rendering upload page")
        return render_template("upload.html", page_name="upload")

    logger.info("upload_pdf: POST request received")

    # ファイル存在チェック
    if "file" not in request.files or not request.files["file"].filename:
        logger.warning("upload_pdf: no file in request")
        return "ファイルが選択されていません。"

    uploaded_file = request.files["file"]
    filename = uploaded_file.filename or ""
    logger.info(f"upload_pdf: uploaded filename={filename}")

    # PDF以外は拒否（早期returnでネスト削減）
    if not filename.lower().endswith(".pdf"):
        logger.warning(f"upload_pdf: uploaded file is not a PDF: {filename}")
        return "PDFファイルをアップロードしてください。"

    # student_id設定確認
    student_id = request.form.get("student_id", "").strip()
    logger.info(f"upload_pdf: student_id={student_id or '<none>'}")

    firebase_settings = None
    if student_id:
        firebase_settings = get_document("messages", student_id)
        if firebase_settings:
            logger.info(f"upload_pdf: applying firebase settings for id={student_id}")
        else:
            logger.info(f"upload_pdf: no firebase settings found for id={student_id}; using defaults")

    try:
        filename = secure_filename(filename)
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        uploaded_file.save(filepath)
        logger.info(f"upload_pdf: saved file to {filepath}")

        result_html = process_pdf(filepath, firebase_settings)
        logger.info(f"upload_pdf: process_pdf completed for {filepath}")
        return result_html

    except Exception as e:
        logger.exception(f"upload_pdf: error processing uploaded file {filename}")
        return f"処理中にエラーが発生しました: {e}", 500


@app.route('/outputs/<path:filepath>')
def serve_output_file(filepath):
    try:
        logger.info(f"serve_output_file: request for {filepath}")
        safe_path = os.path.normpath(filepath)
        full_path = os.path.join(OUTPUT_FOLDER, safe_path)
        full_path = os.path.abspath(full_path)
        output_folder_abs = os.path.abspath(OUTPUT_FOLDER)

        # 出力フォルダ外へのアクセスを防ぐ
        if not (full_path.startswith(output_folder_abs + os.path.sep) or full_path == output_folder_abs):
            return jsonify({"message": "不正なパスです"}), 400

        if not os.path.isfile(full_path):
            return jsonify({"message": "ファイルが見つかりません。"}), 404

        # MIMEタイプを推測して inline で返す（iframe 表示用）
        mimetype, _ = mimetypes.guess_type(full_path)
        if mimetype is None:
            mimetype = "application/octet-stream"

        logger.info(f"serve_output_file: sending file {full_path} with mimetype {mimetype}")
        return send_file(full_path, mimetype=mimetype, as_attachment=False)

    except Exception as e:
        logger.exception("ファイル送信中にエラーが発生しました")
        return jsonify({"message": "内部エラーが発生しました。ログをご確認ください。"}), 500


@app.route("/result")
def result_page():
    try:
        # ここで実際のデータを渡してレンダリング
        pdf_name = request.args.get("pdf_name", "output.pdf")
        dir_name = request.args.get("dir_name", "output")

        # ダミーデータ（テスト用）
        styled_neo_html = "<p>スタイル付きNEOテキストの例</p>"
        neo_content = "<p>NEOタグ付きテキストの例</p>"
        og_tagged_content = "<p>OGタグ付きテキストの例</p>"
        sorted_content = "<p>時系列ソートの例</p>"
        image_gallery_html = "<p>抽出画像の例</p>"
        imgs = []

        return render_template(
            "result.html",
            pdf_name=pdf_name,
            dir_name=dir_name,
            styled_neo_html=styled_neo_html,
            neo_content=neo_content,
            og_tagged_content=og_tagged_content,
            sorted_content=sorted_content,
            image_gallery_html=image_gallery_html,
            imgs=imgs,
            download_html="""
                <a href='/outputs/{0}' target='_blank'>PDFを開く</a>
            """.format(pdf_name)
        )

    except Exception as e:
        logger.exception("result_page: エラーが発生しました")
        return f"エラーが発生しました: {e}", 500


@app.route("/logs")
def view_logs():
    try:
        log_base_dir = "logs"
        log_dirs = sorted(
            [d for d in os.listdir(log_base_dir) if os.path.isdir(os.path.join(log_base_dir, d))],
            reverse=True
        )

        all_logs = []
        for d in log_dirs:
            log_dir = os.path.join(log_base_dir, d)
            app_log_path = os.path.join(log_dir, "app.log")
            error_log_path = os.path.join(log_dir, "error.log")

            app_log = ""
            error_log = ""
            if os.path.exists(app_log_path):
                with open(app_log_path, "r", encoding="utf-8", errors="ignore") as f:
                    app_log = f.read()
            if os.path.exists(error_log_path):
                with open(error_log_path, "r", encoding="utf-8", errors="ignore") as f:
                    error_log = f.read()

            all_logs.append({
                "date": d,
                "app_log": app_log,
                "error_log": error_log,
            })

        # どのフォルダにもログがない場合
        if not any(l["app_log"] or l["error_log"] for l in all_logs):
            message = "現在ログファイルはありません。"
        else:
            message = ""

        return render_template(
            "logs.html",
            page_name="logs",
            message=message,
            all_logs=all_logs
        )

    except Exception as e:
        logging.exception("view_logs: ログ閲覧ページ生成中にエラー発生")
        return f"ログ閲覧ページでエラーが発生しました: {e}", 500


# ダウンロード機能（ファイル名指定で送信）
@app.route("/download/<filename>")
def download_file(filename):
    try:
        file_path = os.path.join(OUTPUT_FOLDER, filename)
        if not os.path.isfile(file_path):
            return "指定されたファイルが存在しません。", 404

        logger.info(f"download_file: {filename} を送信します")
        return send_file(file_path, as_attachment=True)

    except Exception as e:
        logger.exception("download_file: 送信エラー")
        return f"ファイル送信中にエラーが発生しました: {e}", 500


def convert_neo_to_html(neo_content: str,
                        font_size=16,
                        line_height=1.6,
                        font_select="IPAexGothic",
                        app_root=".") -> str:
    """
    NEOタグ形式テキストをHTMLへ変換し、フォント・行間・サイズを反映する
    """

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
                except Exception:
                    pass
            if weight_match:
                current_weight = weight_match.group(1).strip()

            text_content = text_match.group(1).strip() if text_match else ""
            html_lines.append(
                f'<p style="font-family:{current_font}; font-size:{current_size}px; font-weight:{current_weight}; line-height:{current_line_height};">'
                f'{html.escape(text_content)}</p>')

        # 行間設定
        elif line.startswith("[行間]"):
            try:
                current_line_height = float(line.replace("[行間]", "").strip())
            except Exception:
                pass

        # 画像挿入
        elif line.startswith("[画像:"):
            img_match = re.match(
                r"\[画像:(.*?):([\d\.]+):([\d\.]+):([\d\.]+):([\d\.]+)\]", line)
            if img_match:
                img_path = img_match.group(1)
                img_rel_path = img_path.replace(app_root, "").replace(
                    "/home/runner/workspace", "").lstrip("/")
                img_width = img_match.group(4)
                img_height = img_match.group(5)
                html_lines.append(
                    f'<img src="/{img_rel_path}" style="width:{img_width}px; height:{img_height}px; display:block; margin:8px auto;">'
                )

        # 通常テキスト
        else:
            html_lines.append(
                f'<p style="font-family:{current_font}; font-size:{current_size}px; font-weight:{current_weight}; line-height:{current_line_height};">'
                f'{html.escape(line)}</p>')

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


def create_pdf_with_weasyprint(neo_content,
                               output_path,
                               app_root,
                               firebase_settings=None):
    """
    neo_content を解析して HTML を作り、必要なフォントをすべて @font-face で定義して
    WeasyPrint に渡して PDF を生成する（画像は file:// 経由で埋め込み）。
    """
    print("=== NEO解析内容 (先頭800文字) ===")
    print(neo_content[:800])

    try:
        # 使われているフォント名を収集
        font_names = set(re.findall(r'\[フォント:(.*?)\]', neo_content))
        # デフォルトフォントも入れておく
        if firebase_settings and firebase_settings.get("fontSelect"):
            font_names.add(firebase_settings.get("fontSelect"))
        if not font_names:
            font_names.add("IPAexGothic")

        # 各フォント名をファイルパスに解決して @font-face を作る
        font_face_rules = []
        for fname in sorted(font_names):
            # get_font_path は既に定義されている関数を使う
            path = get_font_path(app_root, fname)
            if not path:
                # フォントが見つからなければ ipaex を fallback として使う
                path = get_font_path(app_root, "IPAex明朝") or get_font_path(
                    app_root, "IPAexゴシック")
            if path:
                # file:// フルパスで指定
                font_face_rules.append(
                    f"@font-face {{ font-family: '{fname}'; src: url('file://{path}'); }}"
                )
            else:
                print(f"⚠️ フォントファイル見つからず: {fname}")

        # HTML ブロックを作る
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
                    current_lineheight = float(
                        line.replace("[行間]", "").strip())
                except Exception:
                    current_lineheight = None
                continue

            # 画像タグ
            if line.startswith("[画像:"):
                parts = re.findall(
                    r"\[画像:(.*?):([\d\.]+):([\d\.]+):([\d\.]+):([\d\.]+)\]",
                    line)
                if parts:
                    img_path, x, y, w, h = parts[0]
                    # 画像はローカルファイル経由で埋め込む（WeasyPrint が file:// をサポート）
                    img_file_url = f"file://{os.path.abspath(img_path)}"
                    html_blocks.append(
                        f'<div style="text-align:center; margin: 1em 0;"><img src="{img_file_url}" style="max-width:90%;"></div>'
                    )
                continue

            # フォント/サイズ/ウェイトタグを探す
            font_match = re.search(r"\[フォント:(.*?)\]", line)
            size_match = re.search(r"\[サイズ:(.*?)\]", line)
            weight_match = re.search(r"\[ウェイト:(.*?)\]", line)

            text = re.sub(r"\[.*?\]", "", line).strip()
            if not text:
                continue

            # 決定したフォント情報を使って p タグを作る
            used_font = font_match.group(1).strip() if font_match else (
                firebase_settings.get("fontSelect")
                if firebase_settings else "IPAexGothic")
            used_size = size_match.group(1).strip() if size_match else (
                str(firebase_settings.get("fontSize"))
                if firebase_settings else "16")
            used_weight = weight_match.group(
                1).strip() if weight_match else "normal"

            # line-height の反映（もし current_lineheight があれば）
            lh_css = "line-height:1.6;"
            if current_lineheight:
                # neo の行間が px ベースだったら相当に大きくなるので簡易変換
                try:
                    # 小〜中程度の値に落とす（必要に応じて調整）
                    lh_val = max(1.0, float(current_lineheight) / 20.0)
                    lh_css = f"line-height:{lh_val};"
                except Exception:
                    pass

            # escape
            esc_text = pyhtml.escape(text)
            html_blocks.append(
                f"<p style=\"font-family:'{used_font}'; font-size:{used_size}px; font-weight:{used_weight}; {lh_css} margin:0.3em 0;\">{esc_text}</p>"
            )

        body_html = "\n".join(html_blocks)

        # 最終 HTML テンプレート（フォント定義を head に埋め込む）
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

        # WeasyPrint に書かせる
        # base_url は app_root にしておく（ファイル参照の解決に使われる）
        HTML(string=html_template, base_url=app_root).write_pdf(output_path)

        print(f"✅ PDF生成成功: {output_path}")
        return True, None

    except Exception as e:
        print("❌ PDF生成失敗:", e)
        return False, str(e)


def process_pdf(pdf_path: str, firebase_settings: dict | None = None):
    pdf_name = os.path.basename(pdf_path)
    try:
        doc = fitz.open(pdf_path)
        assert isinstance(doc, fitz.Document)
    except Exception as e:
        return f"PDFを開けません: {e}"

    basename = os.path.splitext(os.path.basename(pdf_path))[0]
    dir_name = os.path.join(OUTPUT_FOLDER, basename)
    os.makedirs(dir_name, exist_ok=True)

    # Firebase設定を取得
    fs_font_override = firebase_settings.get(
        "fontSelect") if firebase_settings else None
    fs_size_add = float(firebase_settings.get("fontSize",
                                              0)) if firebase_settings else 0.0

    # 出力ファイルパス
    output_file_OG = os.path.join(dir_name, f"{basename}_OG.txt")
    output_file_NEO = os.path.join(dir_name, f"{basename}_NEO.txt")
    output_file_SORTED = os.path.join(dir_name, f"{basename}_SORTED.txt")

    neo, sorted_txt, imgs, og_tagged = [], [], [], []

    # ページごとの抽出
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

            # 行間処理
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

            # テキスト要素
            if el["type"] == "text":
                text = el["content"]

                # 元PDFフォント情報を取得 (OG用)
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

                # Firestore設定反映後のフォント (NEO用)
                font = fs_font_override or "IPAexGothic, sans-serif"
                size = og_size + fs_size_add  # 元サイズに加算

                # 出力
                neo.append(
                    f"[フォント:{font}][サイズ:{size:.2f}][ウェイト:normal]{text}\n")
                og_tagged.append(
                    f"[フォント:{og_font}][サイズ:{og_size:.2f}][ウェイト:{og_weight}]{text}\n"
                )
                sorted_txt.append(f"テキスト: {text}\n")

                prev_y = el["bbox"][3]

            # 画像要素
            elif el["type"] == "image":
                bbox = el["bbox"]
                img_tag = f"[画像:{el['content']}:{bbox[0]:.2f}:{bbox[1]:.2f}:{bbox[2]-bbox[0]:.2f}:{bbox[3]-bbox[1]:.2f}]\n"
                neo.append(img_tag)
                og_tagged.append(img_tag)
                sorted_txt.append(f"[画像] {el['content']} | BBOX: {bbox}\n\n")
                prev_y = bbox[3]

    # 出力内容を結合
    neo_content = "".join(neo)
    og_tagged_content = "".join(og_tagged)
    sorted_content = "".join(sorted_txt)

    # ファイル保存
    with open(output_file_NEO, "w", encoding="utf-8") as f:
        f.write(neo_content)
    with open(output_file_SORTED, "w", encoding="utf-8") as f:
        f.write(sorted_content)
    with open(output_file_OG, "w", encoding="utf-8") as f:
        f.write(og_tagged_content)

    # PDF再構築
    recreated_pdf_filename = f"{basename}_recreated.pdf"
    recreated_pdf_path = os.path.join(dir_name, recreated_pdf_filename)
    pdf_ok, pdf_error = create_pdf_with_weasyprint(
        neo_content,
        recreated_pdf_path,
        app_root,
        firebase_settings=firebase_settings)
    recreated_pdf_url = os.path.join(basename, recreated_pdf_filename).replace(
        "\\", "/") if pdf_ok else ""
    if not pdf_ok:
        print("❌ PDF再構成に失敗:", pdf_error)
        recreated_pdf_url = ""
        download_html = "<p style='color:red;'>PDFの再構成に失敗しました。</p>"
    else:
        print("✅ PDF再構成成功:", recreated_pdf_path)
        recreated_pdf_url = os.path.join(basename,
                                         recreated_pdf_filename).replace(
                                             "\\", "/")
        download_html = (
            f'<div class="download-section"><h3>再構成されたPDF</h3>'
            f'<a href="/outputs/{html.escape(recreated_pdf_url)}" '
            f'class="action-link" download>ダウンロード</a></div>')

    # NEOテキスト生成（追加）
    extracted_text = "".join(neo)
    neo_text = extracted_text

    font_size = firebase_settings.get("fontSize",
                                      16) if firebase_settings else 16
    line_height = firebase_settings.get("lineHeight",
                                        1.6) if firebase_settings else 1.6
    font_select = firebase_settings.get(
        "fontSelect", "IPAexGothic") if firebase_settings else "IPAexGothic"

    # HTML生成
    styled_neo_html = convert_neo_to_html(neo_text, font_size, line_height,
                                          font_select, app_root)

    image_gallery_html = "".join(
        f'<a href="/outputs/{html.escape(url)}" target="_blank">'
        f'<img src="/outputs/{html.escape(url)}" alt="image"></a>'
        for url in imgs) or "<p>画像は抽出されませんでした。</p>"

    download_html = (
        f'<div class="download-section"><h3>再構成されたPDF</h3>'
        f'<a href="/outputs/{html.escape(recreated_pdf_url)}" class="action-link" download>ダウンロード</a></div>'
        if pdf_ok else "")

    return render_template(
        "result.html",
        pdf_name=pdf_name,
        dir_name=dir_name,
        download_html=download_html,
        recreated_pdf_url=recreated_pdf_url,
        imgs=imgs,
        styled_neo_html=sanitize_html_for_result(styled_neo_html),
        neo_content=sanitize_html_for_result(neo_content),
        og_tagged_content=sanitize_html_for_result(og_tagged_content),
        sorted_content=sanitize_html_for_result(sorted_content),
        image_gallery_html=image_gallery_html
    )


def sanitize_html_for_result(html):
    """結果ページ用のHTMLをクリーン化（生徒設定フォントなどを除去）"""
    if not html:
        return ""

    # <style>タグを全削除
    html = re.sub(r"<style.*?>.*?</style>", "", html, flags=re.DOTALL)

    # インラインstyle属性を削除（font-family, line-heightなど）
    html = re.sub(r'style="[^"]*"', "", html)

    # spanなどの余分なタグを整理
    html = re.sub(r'\s+', ' ', html)

    return html.strip()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 3000))
    print("DEBUG: Logging handlers:", logging.getLogger().handlers)
    app.run(debug=False, host="0.0.0.0", port=port)