import os
import subprocess
import sys
from flask import Flask, render_template, request, Response, stream_with_context, jsonify, send_from_directory
from werkzeug.utils import secure_filename
from datetime import datetime
import json
import re
import threading
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.json.ensure_ascii = False

# Configurações
UPLOAD_FOLDER = 'templates'
REPORTS_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}

ACTIVE_JOBS = 0
JOBS_LOCK = threading.Lock()

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
TEMPLATE_FILENAME = 'ModeloSolicitacaoMob.xlsx'
TEMPLATE_STATUS_FILE = os.path.join('templates', 'template_update_status.json')


def start_job():
    global ACTIVE_JOBS
    with JOBS_LOCK:
        ACTIVE_JOBS += 1


def end_job():
    global ACTIVE_JOBS
    with JOBS_LOCK:
        if ACTIVE_JOBS > 0:
            ACTIVE_JOBS -= 1


def generate_report_excel(report_data: dict) -> str:
    """
    Gera um arquivo Excel formatado a partir do JSON de relatório emitido pelo PowerShell.
    Retorna o caminho absoluto do arquivo gerado.
    """
    # ── Cores Vestas ──────────────────────────────────────────────────────────
    COLOR_NIGHT_SKY  = "1F3144"
    COLOR_BLUE_SKY   = "005AFF"
    COLOR_LIGHT_GREY = "E3E5E8"
    COLOR_GREEN_FILL = "D4EDDA"
    COLOR_RED_FILL   = "F8D7DA"
    COLOR_WHITE      = "FFFFFF"
    COLOR_LABEL_BG   = "EBF3FB"

    fill_header    = PatternFill("solid", fgColor=COLOR_NIGHT_SKY)
    fill_col_hdr   = PatternFill("solid", fgColor=COLOR_NIGHT_SKY)
    fill_label     = PatternFill("solid", fgColor=COLOR_LABEL_BG)
    fill_success   = PatternFill("solid", fgColor=COLOR_GREEN_FILL)
    fill_error     = PatternFill("solid", fgColor=COLOR_RED_FILL)
    fill_light     = PatternFill("solid", fgColor=COLOR_LIGHT_GREY)

    font_white_bold  = Font(name="Calibri", bold=True,  color=COLOR_WHITE, size=13)
    font_col_hdr     = Font(name="Calibri", bold=True,  color=COLOR_WHITE, size=10)
    font_label       = Font(name="Calibri", bold=True,  color=COLOR_NIGHT_SKY, size=10)
    font_value       = Font(name="Calibri", bold=False, color=COLOR_NIGHT_SKY, size=10)
    font_value_bold  = Font(name="Calibri", bold=True,  color=COLOR_NIGHT_SKY, size=10)

    thin_border_side = Side(style="thin", color="CCCCCC")
    thin_border      = Border(
        left=thin_border_side, right=thin_border_side,
        top=thin_border_side,  bottom=thin_border_side
    )

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Relatório de Submissão"

    items             = report_data.get("items", [])
    submission_dt     = report_data.get("submission_datetime", "")
    id_mob            = report_data.get("id_mobilizacao", "")
    requester_name    = report_data.get("requester_name", "")
    requester_email   = report_data.get("requester_email", "")

    # ── Coletar todos os display names usados, preservando ordem de aparição ──
    display_cols: list[str] = []
    seen: set[str] = set()
    for item in items:
        for col_name in (item.get("fields") or {}).keys():
            if col_name not in seen:
                seen.add(col_name)
                display_cols.append(col_name)

    total_cols = 2 + len(display_cols)  # "ID SP" + "Status" + campo...
    last_col_letter = get_column_letter(max(total_cols, 3))

    # ── ROW 1 – Título principal ───────────────────────────────────────────────
    ws.merge_cells(f"A1:{last_col_letter}1")
    title_cell = ws["A1"]
    title_cell.value = "RELATÓRIO DE SUBMISSÃO — MOBILIZAÇÕES VESTAS"
    title_cell.fill  = fill_header
    title_cell.font  = font_white_bold
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # ── ROWS 2-5 – Metadados da submissão ─────────────────────────────────────
    meta_rows = [
        ("Data/Hora da Submissão:", submission_dt),
        ("ID de Mobilização:",     id_mob),
        ("Solicitante:",           requester_name),
        ("E-mail:",                requester_email),
    ]
    for r_offset, (label, value) in enumerate(meta_rows, start=2):
        label_cell = ws.cell(row=r_offset, column=1, value=label)
        label_cell.fill  = fill_label
        label_cell.font  = font_label
        label_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        label_cell.border = thin_border

        value_cell = ws.cell(row=r_offset, column=2, value=value)
        value_cell.fill  = fill_light
        value_cell.font  = font_value_bold
        value_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        value_cell.border = thin_border

        # Mesclar colunas restantes do valor
        if total_cols > 2:
            ws.merge_cells(start_row=r_offset, start_column=2,
                           end_row=r_offset,   end_column=total_cols)
        ws.row_dimensions[r_offset].height = 18

    # ── ROW 6 – Linha de separação ────────────────────────────────────────────
    ws.row_dimensions[6].height = 8

    # ── ROW 7 – Cabeçalho das colunas ─────────────────────────────────────────
    hdr_row = 7
    headers = ["ID SP", "Status"] + display_cols
    for col_idx, header_text in enumerate(headers, start=1):
        cell = ws.cell(row=hdr_row, column=col_idx, value=header_text)
        cell.fill      = fill_col_hdr
        cell.font      = font_col_hdr
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = thin_border
    ws.row_dimensions[hdr_row].height = 22

    # ── ROWS 8+ – Dados ───────────────────────────────────────────────────────
    for row_offset, item in enumerate(items, start=hdr_row + 1):
        is_error = "Erro" in str(item.get("status", ""))
        row_fill = fill_error if is_error else fill_success

        # ID SP
        id_sp_cell = ws.cell(row=row_offset, column=1, value=item.get("id_sp", ""))
        id_sp_cell.fill      = row_fill
        id_sp_cell.font      = font_value_bold
        id_sp_cell.alignment = Alignment(horizontal="center", vertical="center")
        id_sp_cell.border    = thin_border

        # Status
        status_cell = ws.cell(row=row_offset, column=2, value=item.get("status", ""))
        status_cell.fill      = row_fill
        status_cell.font      = font_value
        status_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
        status_cell.border    = thin_border

        # Campos do item
        fields = item.get("fields") or {}
        for col_offset, col_name in enumerate(display_cols, start=3):
            val = fields.get(col_name, "")
            data_cell = ws.cell(row=row_offset, column=col_offset, value=val)
            data_cell.fill      = row_fill
            data_cell.font      = font_value
            data_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
            data_cell.border    = thin_border

        ws.row_dimensions[row_offset].height = 18

    # ── Largura das colunas ───────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 10   # ID SP
    ws.column_dimensions["B"].width = 30   # Status
    for col_offset, col_name in enumerate(display_cols, start=3):
        estimated = max(14, min(len(col_name) + 4, 40))
        ws.column_dimensions[get_column_letter(col_offset)].width = estimated

    # ── Congelar cabeçalho ────────────────────────────────────────────────────
    ws.freeze_panes = f"A{hdr_row + 1}"

    # ── Salvar ────────────────────────────────────────────────────────────────
    os.makedirs(REPORTS_FOLDER, exist_ok=True)
    safe_id  = re.sub(r"[^A-Za-z0-9]", "_", id_mob) if id_mob else "sem_id"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_filename = f"Relatorio_Mob_{safe_id}_{timestamp}.xlsx"
    report_path     = os.path.abspath(os.path.join(REPORTS_FOLDER, report_filename))
    wb.save(report_path)
    return report_filename, report_path


def build_powershell_command(script_path, file_path, sheet_name):
    """Executa scripts PowerShell com stdout em UTF-8 para preservar acentuação."""
    utf8_preamble = (
        "[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false); "
        "[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false); "
        "$OutputEncoding = [Console]::OutputEncoding; "
        "chcp 65001 > $null; "
    )
    escaped_script_path = script_path.replace("'", "''")
    escaped_file_path = file_path.replace("'", "''")
    escaped_sheet_name = sheet_name.replace("'", "''")

    return [
        "powershell.exe",
        "-ExecutionPolicy", "Bypass",
        "-NoProfile",
        "-Command",
        (
            f"{utf8_preamble}& '{escaped_script_path}' "
            f"-ExcelPath '{escaped_file_path}' -SheetName '{escaped_sheet_name}'"
        )
    ]


def build_powershell_populate_command(script_path, file_path):
    """Executa Populate-SharePointList.ps1 sem o parâmetro -SheetName (processamento unificado)."""
    utf8_preamble = (
        "[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false); "
        "[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false); "
        "$OutputEncoding = [Console]::OutputEncoding; "
        "chcp 65001 > $null; "
    )
    escaped_script_path = script_path.replace("'", "''")
    escaped_file_path = file_path.replace("'", "''")

    return [
        "powershell.exe",
        "-ExecutionPolicy", "Bypass",
        "-NoProfile",
        "-Command",
        (
            f"{utf8_preamble}& '{escaped_script_path}' "
            f"-ExcelPath '{escaped_file_path}'"
        )
    ]


def build_powershell_template_update_command(script_path, template_path):
    """Executa script PowerShell de atualização do template com stdout em UTF-8."""
    utf8_preamble = (
        "[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false); "
        "[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false); "
        "$OutputEncoding = [Console]::OutputEncoding; "
        "chcp 65001 > $null; "
    )
    escaped_script_path = script_path.replace("'", "''")
    escaped_template_path = template_path.replace("'", "''")

    return [
        "powershell.exe",
        "-ExecutionPolicy", "Bypass",
        "-NoProfile",
        "-Command",
        (
            f"{utf8_preamble}& '{escaped_script_path}' "
            f"-TemplatePath '{escaped_template_path}'"
        )
    ]


def get_template_status():
    if not os.path.exists(TEMPLATE_STATUS_FILE):
        return {'last_updated': None}

    try:
        with open(TEMPLATE_STATUS_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
            return {'last_updated': data.get('last_updated')}
    except Exception:
        return {'last_updated': None}


def save_template_status(iso_datetime):
    payload = {'last_updated': iso_datetime}
    with open(TEMPLATE_STATUS_FILE, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def decode_powershell_output(raw_output):
    """Decodifica stdout do Windows PowerShell 5.1 preservando acentuação."""
    utf8_text = raw_output.decode('utf-8', errors='replace')
    cp1252_text = raw_output.decode('cp1252', errors='replace')
    repaired_cp1252 = repair_mojibake(cp1252_text)

    utf8_score = score_decoded_text(utf8_text)
    repaired_cp1252_score = score_decoded_text(repaired_cp1252)

    if repaired_cp1252_score < utf8_score:
        return repaired_cp1252
    return utf8_text


def repair_mojibake(text):
    """Corrige trechos UTF-8 lidos como cp1252 sem afetar texto já correto."""
    previous_text = text
    for _ in range(3):
        repaired_text = repair_mojibake_once(previous_text)
        if score_decoded_text(repaired_text) >= score_decoded_text(previous_text):
            return previous_text
        previous_text = repaired_text
    return previous_text


def repair_mojibake_once(text):
    parts = re.split(r'(\s+)', text)
    repaired_parts = []
    for part in parts:
        if not part or part.isspace() or not has_mojibake_markers(part):
            repaired_parts.append(part)
            continue

        try:
            repaired_candidate = part.encode('cp1252', errors='strict').decode('utf-8', errors='strict')
        except (UnicodeEncodeError, UnicodeDecodeError):
            repaired_parts.append(part)
            continue

        if score_decoded_text(repaired_candidate) <= score_decoded_text(part):
            repaired_parts.append(repaired_candidate)
        else:
            repaired_parts.append(part)

    return ''.join(repaired_parts)


def has_mojibake_markers(text):
    return any(marker in text for marker in ('Ã', 'Â', 'â'))


def score_decoded_text(text):
    replacement_penalty = text.count('�') * 10
    mojibake_penalty = sum(text.count(marker) for marker in ('Ã', 'Â', 'â')) * 6
    return replacement_penalty + mojibake_penalty

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/download-template')
def download_template():
    return send_from_directory(app.config['UPLOAD_FOLDER'], TEMPLATE_FILENAME, as_attachment=True)


@app.route('/template-update-status', methods=['GET'])
def template_update_status():
    return jsonify(get_template_status()), 200


@app.route('/update-template', methods=['POST'])
def update_template():
    start_job()
    try:
        template_path = os.path.abspath(os.path.join(app.config['UPLOAD_FOLDER'], TEMPLATE_FILENAME))
        if not os.path.exists(template_path):
            return jsonify({'status': 'error', 'message': 'Template não encontrado.'}), 404

        script_path = os.path.abspath('Update-ExcelTemplateChoices.ps1')
        if not os.path.exists(script_path):
            return jsonify({'status': 'error', 'message': 'Script de atualização não encontrado.'}), 500

        cmd = build_powershell_template_update_command(script_path, template_path)

        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            universal_newlines=False
        )
        raw_output = process.communicate()[0]
        process.wait()
        full_output = decode_powershell_output(raw_output)

        if process.returncode != 0:
            return jsonify({
                'status': 'error',
                'message': 'Falha ao atualizar template.',
                'log': full_output
            }), 500

        updated_at = datetime.now().isoformat()
        save_template_status(updated_at)

        return jsonify({
            'status': 'success',
            'message': 'Template atualizado com sucesso.',
            'last_updated': updated_at,
            'log': full_output
        }), 200
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': f'Erro ao atualizar template: {str(e)}'
        }), 500
    finally:
        end_job()

@app.route('/validate', methods=['POST'])
def validate():
    """Fase 1: Executa validação separada antes do upload"""
    start_job()
    try:
        file = request.files.get('file')
        sheet_name = request.form.get('sheet', 'PESSOAS')
        sheet_name = (sheet_name or 'PESSOAS').strip() or 'PESSOAS'

        if not file or not allowed_file(file.filename):
            return jsonify({'status': 'error', 'errors': ['Arquivo inválido ou não enviado.']}), 400

        # Bloqueia template vazio para evitar validação/upload sem conteúdo.
        try:
            file.stream.seek(0, os.SEEK_END)
            uploaded_size = file.stream.tell()
            file.stream.seek(0)
        except Exception:
            uploaded_size = None

        if uploaded_size == 0:
            return jsonify({'status': 'error', 'errors': ['O arquivo enviado está vazio.']}), 400

        filename = secure_filename(file.filename)
        if not filename:
            return jsonify({'status': 'error', 'errors': ['Nome de arquivo inválido.']}), 400

        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)

        if os.path.getsize(file_path) == 0:
            if os.path.exists(file_path):
                os.remove(file_path)
            return jsonify({'status': 'error', 'errors': ['O arquivo enviado está vazio.']}), 400

        script_path = os.path.abspath("Validate-ExcelData.ps1")
        cmd = build_powershell_command(script_path, file_path, sheet_name)

        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            universal_newlines=False
        )
        raw_output = process.communicate()[0]
        process.wait()
        full_output = decode_powershell_output(raw_output)

        # Extrair JSON de validação da saída
        start_marker = '---VALIDATION_JSON_START---'
        end_marker = '---VALIDATION_JSON_END---'
        
        if start_marker in full_output and end_marker in full_output:
            start_idx = full_output.index(start_marker) + len(start_marker)
            end_idx = full_output.index(end_marker)
            json_str = full_output[start_idx:end_idx].strip()
            result = json.loads(json_str)

            total_lines = result.get('total_lines', 0)
            try:
                total_lines = int(total_lines)
            except (TypeError, ValueError):
                total_lines = 0

            if total_lines <= 0:
                return jsonify({
                    'status': 'error',
                    'errors': ['Nenhuma linha válida encontrada para importar no SharePoint.'],
                    'log': full_output,
                    'filename': filename,
                    'total_lines': total_lines
                }), 200
            
            # Adicionar log da validação para debug
            result['log'] = full_output
            
            # Incluir filename para a fase 2 reutilizar
            result['filename'] = filename
            return jsonify(result), 200
        else:
            return jsonify({
                'status': 'error',
                'errors': ['Não foi possível obter resultado da validação.'],
                'log': full_output,
                'filename': filename
            }), 200

    except Exception as e:
        return jsonify({
            'status': 'error',
            'errors': [f'Erro ao executar validação: {str(e)}']
        }), 500
    finally:
        end_job()


@app.route('/run-script', methods=['POST'])
def run_script():
    """Fase 2: Upload para SharePoint (usa o arquivo já salvo pela validação)"""
    data = request.get_json()
    if not data:
        return Response("Erro: Dados não enviados.", status=400)

    filename = data.get('filename', '')

    if not filename or not allowed_file(filename):
        return Response("Erro: Arquivo inválido.", status=400)

    file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(filename))
    if not os.path.exists(file_path):
        return Response("Erro: Arquivo não encontrado. Execute a validação primeiro.", status=400)

    # Caminho absoluto para o script PowerShell
    script_path = os.path.abspath("Populate-SharePointList.ps1")

    # Processamento unificado: PESSOAS + EQUIPAMENTOS (sem -SheetName)
    cmd = build_powershell_populate_command(script_path, file_path)

    def generate():
        start_job()
        yield f"Arquivo recebido: {filename}\n"
        yield f"Processando abas PESSOAS e EQUIPAMENTOS...\n"
        yield f"Executando script PowerShell...\n{'-'*30}\n"

        report_filename = None

        try:
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                bufsize=1,
                universal_newlines=False
            )

            detected_error_in_output = False
            error_markers = [
                "FALHA CRÍTICA",
                "UPLOAD CANCELADO",
                "--- RESULT: ERROR ---",
                "Write-Error",
                "Erro ao adicionar item",
                "Não foi possível gerar um ID_Mobilizacao único",
            ]

            # Acumulador para extrair o bloco JSON do relatório
            output_buffer: list[str] = []
            in_report_json = False
            report_json_lines: list[str] = []

            REPORT_START = "---REPORT_JSON_START---"
            REPORT_END   = "---REPORT_JSON_END---"

            while True:
                chunk = process.stdout.readline()
                if not chunk:
                    break
                decoded = decode_powershell_output(chunk)

                stripped = decoded.strip()

                # Captura bloco JSON do relatório sem emitir para o stream
                if stripped == REPORT_START:
                    in_report_json = True
                    continue
                if stripped == REPORT_END:
                    in_report_json = False
                    continue
                if in_report_json:
                    report_json_lines.append(stripped)
                    continue

                if any(marker in decoded for marker in error_markers):
                    detected_error_in_output = True
                yield decoded

            process.wait()

            # ── Gerar Excel do relatório ──────────────────────────────────────
            if report_json_lines:
                try:
                    report_data = json.loads("".join(report_json_lines))
                    report_filename, _ = generate_report_excel(report_data)
                    yield f"\n---REPORT_FILE:{report_filename}---\n"
                except Exception as exc:
                    yield f"\n[AVISO] Não foi possível gerar o relatório Excel: {exc}\n"

            if process.returncode == 0 and not detected_error_in_output:
                yield f"\n{'-'*30}\n[SUCESSO] Processo finalizado com código 0.\n"
            else:
                yield f"\n{'-'*30}\n[ERRO] Processo finalizado com código {process.returncode}.\n"

        except Exception as e:
            yield f"\n[ERRO DE EXECUÇÃO]: {str(e)}\n"
        finally:
            end_job()

    return Response(stream_with_context(generate()), content_type='text/plain; charset=utf-8')

@app.route('/shutdown', methods=['POST'])
def shutdown():
    """Encerra o servidor Flask."""
    def _shutdown():
        os._exit(0)
    t = threading.Timer(0.5, _shutdown)
    t.daemon = True
    t.start()
    return jsonify({'message': 'Servidor encerrado.'}), 200


@app.route('/list-reports')
def list_reports():
    """Lista os relatórios Excel gerados, ordenados do mais recente para o mais antigo."""
    reports_abs = os.path.abspath(REPORTS_FOLDER)
    if not os.path.isdir(reports_abs):
        return jsonify({'reports': []})
    files = []
    for fname in os.listdir(reports_abs):
        if fname.lower().endswith('.xlsx'):
            fpath = os.path.join(reports_abs, fname)
            try:
                mtime = os.path.getmtime(fpath)
                size  = os.path.getsize(fpath)
            except OSError:
                continue
            files.append({'name': fname, 'mtime': mtime, 'size': size})
    files.sort(key=lambda f: f['mtime'], reverse=True)
    for f in files:
        from datetime import datetime as _dt
        f['modified'] = _dt.fromtimestamp(f['mtime']).strftime('%d/%m/%Y %H:%M:%S')
        del f['mtime']
    return jsonify({'reports': files})


@app.route('/download-report/<path:filename>')
def download_report(filename):
    """Serve o relatório Excel gerado após a submissão."""
    safe_name = secure_filename(filename)
    reports_abs = os.path.abspath(REPORTS_FOLDER)
    file_abs    = os.path.abspath(os.path.join(reports_abs, safe_name))
    # Previne path traversal
    if not file_abs.startswith(reports_abs + os.sep):
        return jsonify({'error': 'Acesso negado.'}), 403
    if not os.path.exists(file_abs):
        return jsonify({'error': 'Relatório não encontrado.'}), 404
    return send_from_directory(reports_abs, safe_name, as_attachment=True)


if __name__ == '__main__':
    print("Servidor rodando em http://localhost:5000")
    app.run(debug=False, host='0.0.0.0', port=5000, threaded=True)
