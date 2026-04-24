from __future__ import annotations

from datetime import datetime
import hashlib
from pathlib import Path
import tempfile
import webbrowser

from flask import Flask, jsonify, render_template, request

from compare_contracts import ComparisonRejectedError, compare_contract_files, get_runtime_dir


ALLOWED_EXTENSIONS = {".docx"}

app = Flask(__name__, template_folder="templates")


def _is_docx(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/api/compare")
def compare_api():
    file_a = request.files.get("fileA")
    file_b = request.files.get("fileB")
    output_dir_raw = (request.form.get("outputDir") or "").strip()

    if not file_a or not file_b:
        return jsonify({"ok": False, "message": "请上传两份 docx 文件。"}), 400

    if not _is_docx(file_a.filename or "") or not _is_docx(file_b.filename or ""):
        return jsonify({"ok": False, "message": "只支持 .docx 文件。"}), 400

    runtime_dir = get_runtime_dir()
    output_dir = Path(output_dir_raw).expanduser() if output_dir_raw else runtime_dir

    if not output_dir.is_absolute():
        output_dir = (runtime_dir / output_dir).resolve()

    output_dir.mkdir(parents=True, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_excel = output_dir / f"文档差异结果{ts}.xlsx"

    with tempfile.TemporaryDirectory(prefix="contract_cmp_", dir=str(runtime_dir)) as tmp:
        tmp_dir = Path(tmp)
        # 避免中文文件名经清洗后冲突，固定使用不同的临时文件名。
        file_a_path = tmp_dir / "first_upload.docx"
        file_b_path = tmp_dir / "second_upload.docx"
        file_a.save(file_a_path)
        file_b.save(file_b_path)

        hash_a = hashlib.sha256(file_a_path.read_bytes()).hexdigest()
        hash_b = hashlib.sha256(file_b_path.read_bytes()).hexdigest()
        if hash_a == hash_b:
            return jsonify({"ok": False, "message": "两份文件完全一致，请确认上传了两个不同版本。"}), 400

        try:
            display_a = (file_a.filename or "").strip() or "第一份文档.docx"
            display_b = (file_b.filename or "").strip() or "第二份文档.docx"
            diff_count = compare_contract_files(
                file_a_path,
                file_b_path,
                output_excel,
                first_doc_name=display_a,
                second_doc_name=display_b,
            )
        except ComparisonRejectedError as exc:
            return jsonify({"ok": False, "message": str(exc)}), 400
        except Exception as exc:  # noqa: BLE001
            return jsonify({"ok": False, "message": f"分析失败：{exc}"}), 500

    return jsonify(
        {
            "ok": True,
            "message": "分析完成，结果已保存到指定目录。",
            "diffCount": diff_count,
            "outputPath": str(output_excel),
        }
    )


@app.get("/api/health")
def health():
    return jsonify({"ok": True})


def main() -> None:
    host = "127.0.0.1"
    port = 8080
    webbrowser.open(f"http://{host}:{port}", new=2)
    app.run(host=host, port=port, debug=False)


if __name__ == "__main__":
    main()
