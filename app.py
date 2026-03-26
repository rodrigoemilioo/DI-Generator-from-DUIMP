"""
app.py
------
Servidor Flask principal.
Rotas:
  GET  /           → Página de upload
  POST /converter  → Processa o arquivo e retorna o Excel para download
  GET  /health     → Verificação de saúde da API
"""

import os
import uuid
from pathlib import Path
from datetime import datetime

from flask import (Flask, render_template, request,
                   send_file, jsonify, flash, redirect, url_for)
from werkzeug.utils import secure_filename

from parser import parse_file
from processor import process_data
from generator import generate_excel

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURAÇÃO
# ─────────────────────────────────────────────────────────────────────────────

BASE_DIR     = Path(__file__).parent
UPLOAD_DIR   = BASE_DIR / "uploads"
OUTPUT_DIR   = BASE_DIR / "outputs"
ALLOWED_EXT  = {".xml", ".json", ".xlsx", ".xls", ".xlsm"}
MAX_SIZE_MB  = 20

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "duimp-di-converter-key-2024")
app.config["MAX_CONTENT_LENGTH"] = MAX_SIZE_MB * 1024 * 1024


# ─────────────────────────────────────────────────────────────────────────────
# UTILITÁRIOS
# ─────────────────────────────────────────────────────────────────────────────

def _allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXT


def _cleanup_old_files(directory: Path, max_age_hours: int = 2):
    """Remove arquivos gerados há mais de X horas."""
    now = datetime.now().timestamp()
    for f in directory.glob("*"):
        if f.is_file():
            age_hours = (now - f.stat().st_mtime) / 3600
            if age_hours > max_age_hours:
                try:
                    f.unlink()
                except OSError:
                    pass


# ─────────────────────────────────────────────────────────────────────────────
# ROTAS
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    """Página inicial com formulário de upload."""
    return render_template("index.html")


@app.route("/converter", methods=["POST"])
def converter():
    """
    Recebe o arquivo via POST, processa e retorna o Excel.
    Campos do form:
      - file          : arquivo obrigatório
      - taxa_cambio   : taxa de câmbio manual (opcional, float)
    """
    # ── Validação do upload ──
    if "file" not in request.files:
        flash("Nenhum arquivo enviado.", "danger")
        return redirect(url_for("index"))

    uploaded_file = request.files["file"]
    if not uploaded_file.filename:
        flash("Nome de arquivo vazio.", "danger")
        return redirect(url_for("index"))

    if not _allowed_file(uploaded_file.filename):
        ext_lista = ", ".join(sorted(ALLOWED_EXT))
        flash(f"Formato não suportado. Use: {ext_lista}", "danger")
        return redirect(url_for("index"))

    # ── Taxa de câmbio manual ──
    taxa_str = request.form.get("taxa_cambio", "").strip()
    taxa_override = None
    if taxa_str:
        try:
            taxa_override = float(taxa_str.replace(",", "."))
            if taxa_override <= 0:
                raise ValueError
        except ValueError:
            flash("Taxa de câmbio inválida. Use um número positivo (ex: 5.24).", "warning")
            return redirect(url_for("index"))

    # ── Salva o arquivo enviado ──
    safe_name   = secure_filename(uploaded_file.filename)
    unique_id   = uuid.uuid4().hex[:8]
    input_path  = UPLOAD_DIR / f"{unique_id}_{safe_name}"
    uploaded_file.save(str(input_path))

    # ── Processamento ──
    output_path = OUTPUT_DIR / f"DI_Convertida_{unique_id}.xlsx"
    try:
        # 1. Parse
        raw = parse_file(str(input_path))

        if not raw.get("adicoes"):
            flash("Nenhuma adição encontrada no arquivo. "
                  "Verifique se o formato está correto.", "warning")
            return redirect(url_for("index"))

        # 2. Processo
        processed = process_data(raw, taxa_cambio_override=taxa_override)

        # 3. Gera Excel
        generate_excel(processed, str(output_path))

        _cleanup_old_files(UPLOAD_DIR)
        _cleanup_old_files(OUTPUT_DIR)

        num_adicoes = len(processed["adicoes"])
        total_trib  = processed["totais"]["total_tributos"]

        # Download direto
        return send_file(
            str(output_path),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=f"DI_{processed['header'].get('numero_duimp', unique_id)}.xlsx"
        )

    except (FileNotFoundError, ValueError) as e:
        flash(f"Erro no processamento: {e}", "danger")
        return redirect(url_for("index"))
    except Exception as e:
        app.logger.exception("Erro inesperado ao converter DUIMP")
        flash(f"Erro inesperado: {e}", "danger")
        return redirect(url_for("index"))
    finally:
        # Remove arquivo de upload após processamento
        try:
            input_path.unlink(missing_ok=True)
        except Exception:
            pass


@app.route("/api/converter", methods=["POST"])
def api_converter():
    """
    Versão API JSON do conversor.
    Retorna JSON com resumo + link para download.
    Aceita multipart/form-data com campo 'file'.
    """
    if "file" not in request.files:
        return jsonify({"error": "Campo 'file' ausente no request."}), 400

    uploaded_file = request.files["file"]
    if not _allowed_file(uploaded_file.filename):
        return jsonify({"error": "Formato não suportado."}), 400

    taxa_str = request.form.get("taxa_cambio", "")
    taxa_override = None
    if taxa_str:
        try:
            taxa_override = float(taxa_str.replace(",", "."))
        except ValueError:
            pass

    safe_name  = secure_filename(uploaded_file.filename)
    unique_id  = uuid.uuid4().hex[:8]
    input_path = UPLOAD_DIR / f"{unique_id}_{safe_name}"
    uploaded_file.save(str(input_path))
    output_path = OUTPUT_DIR / f"DI_Convertida_{unique_id}.xlsx"

    try:
        raw       = parse_file(str(input_path))
        processed = process_data(raw, taxa_cambio_override=taxa_override)
        generate_excel(processed, str(output_path))

        return jsonify({
            "status": "success",
            "numero_duimp": processed["header"].get("numero_duimp", ""),
            "total_adicoes": len(processed["adicoes"]),
            "totais": processed["totais"],
            "taxa_cambio": processed["taxa_cambio_usada"],
            "download_id": unique_id,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        try:
            input_path.unlink(missing_ok=True)
        except Exception:
            pass


@app.route("/download/<unique_id>")
def download(unique_id: str):
    """Rota de download pelo ID gerado pela API."""
    # Segurança: valida que o ID é hexadecimal
    if not all(c in "0123456789abcdef" for c in unique_id):
        return jsonify({"error": "ID inválido"}), 400

    matches = list(OUTPUT_DIR.glob(f"DI_Convertida_{unique_id}.xlsx"))
    if not matches:
        return jsonify({"error": "Arquivo não encontrado ou expirado."}), 404

    return send_file(
        str(matches[0]),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=matches[0].name
    )


@app.route("/health")
def health():
    return jsonify({"status": "ok", "timestamp": datetime.now().isoformat()})


@app.errorhandler(413)
def file_too_large(e):
    flash(f"Arquivo muito grande. Limite: {MAX_SIZE_MB}MB.", "danger")
    return redirect(url_for("index"))


# ─────────────────────────────────────────────────────────────────────────────
# INICIALIZAÇÃO
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import logging
    logging.basicConfig(level=logging.INFO)
    print("=" * 60)
    print("  DUIMP to DI Converter – Iniciando servidor...")
    print("  Acesse: http://localhost:5000")
    print("=" * 60)
    app.run(debug=True, host="0.0.0.0", port=5000)
