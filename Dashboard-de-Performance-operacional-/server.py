"""
Projeto Demo — Servidor Flask v2 — Performance Corrigida
"""

import os, threading, webbrowser, traceback, datetime as _dt
from datetime import datetime
from pathlib import Path

from flask import Flask, request, jsonify, send_from_directory
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment

BASE_DIR   = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
PORT       = int(os.environ.get("PORT", 5050))

# ── Caminho do banco ──────────────────────────────────────────────────────────
_env_data  = os.environ.get("APP_DATA", "")
DATA_DIR   = Path(_env_data) if _env_data else (BASE_DIR / "data")
EXCEL_PATH = DATA_DIR / "relatorios.xlsx"
EXCEL_BACKUP = BASE_DIR / "data" / "relatorios_backup.xlsx"

app = Flask(__name__, static_folder=str(STATIC_DIR))

def _salvar_excel(wb):
    import shutil
    wb.save(EXCEL_PATH)
    def _backup():
        try:
            EXCEL_BACKUP.parent.mkdir(parents=True, exist_ok=True)
            if EXCEL_PATH.resolve() != EXCEL_BACKUP.resolve():
                shutil.copy2(EXCEL_PATH, EXCEL_BACKUP)
        except Exception as e:
            print(f"[AVISO] Backup local falhou: {e}")
    threading.Thread(target=_backup, daemon=True).start()

lock = threading.RLock()

ABA_TURNOS   = "Turnos"
ABA_DESCARGA = "Descarga"
ABA_PARADAS  = "Paradas"
ABA_REL      = "RelatorioTurno"

EQ_MAP = {"Moega 1":"m1","Moega 2":"m2","m1":"m1","m2":"m2"}

# ── Inicializar banco ─────────────────────────────────────────────────────────
def garantir_banco():
    try:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    if not EXCEL_PATH.exists():
        inicializar_excel()

def inicializar_excel():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if EXCEL_PATH.exists(): return
    wb = Workbook()
    wb.remove(wb.active)
    specs = {
        ABA_TURNOS:   ["id","data","periodo","turno","encarregado","operador","bal1","bal2","criado_em"],
        ABA_DESCARGA: ["id","turno_id","equipamento","produto","hora_ini","hora_fim","quantidade"],
        ABA_PARADAS:  ["id","turno_id","equipamento","causa","categoria","hora_ini","hora_fim","duracao_min","obs"],
        ABA_REL:      ["id","turno_id","dds","qual_desvio","ssma_desvio","outros_desvio",
                       "prev_soja_prog","prev_soja_real","prev_milho_prog","prev_milho_real",
                       "prev_farelo_prog","prev_farelo_real","saldo_json","hora_a_hora_json"],
    }
    fill = PatternFill("solid", fgColor="1A4F7A")
    font = Font(color="FFFFFF", bold=True)
    for aba, cols in specs.items():
        ws = wb.create_sheet(aba)
        ws.append(cols)
        for cell in ws[1]:
            cell.fill = fill; cell.font = font
            cell.alignment = Alignment(horizontal="center")
    _salvar_excel(wb)

def proximo_id(ws):
    return max((r[0] for r in ws.iter_rows(min_row=2, values_only=True) if r and r[0]), default=0) + 1

def buscar_turno_id(wb, data, periodo):
    for row in wb[ABA_TURNOS].iter_rows(min_row=2, values_only=True):
        if not row or not row[1]: continue
        row_data = row[1].strftime("%Y-%m-%d") if isinstance(row[1], (_dt.datetime, _dt.date)) else str(row[1])[:10]
        if row_data == data and row[2] == periodo:
            return row[0]
    return None

@app.route("/")
def index(): return send_from_directory(STATIC_DIR, "relatorio.html")

@app.route("/dashboard")
def dashboard(): return send_from_directory(STATIC_DIR, "dashboard.html")

@app.route("/api/turno")
def get_turno():
    data    = request.args.get("data")
    periodo = request.args.get("periodo")
    if not data or not periodo:
        return jsonify({"erro": "data e periodo obrigatorios"}), 400

    with lock:
        # data_only=True é vital aqui para ler o valor calculado/digitado e não a fórmula
        wb  = load_workbook(EXCEL_PATH, data_only=True) 
        tid = buscar_turno_id(wb, data, periodo)
        if tid is None:
            return jsonify({"encontrado": False})

        # Pega a linha do turno
        tr = next((r for r in wb[ABA_TURNOS].iter_rows(min_row=2, values_only=True) if r[0] == tid), None)

        # CORREÇÃO DA DESCARGA: Mapeamento explícito dos campos
        descarga = []
        for r in wb[ABA_DESCARGA].iter_rows(min_row=2, values_only=True):
            if r[0] and r[1] == tid:
                descarga.append({
                    "equipamento": r[2], 
                    "produto": r[3],
                    "hora_ini": str(r[4])[:5] if r[4] else "", # Pega apenas HH:MM
                    "hora_fim": str(r[5])[:5] if r[5] else "", # <--- AQUI ESTAVA O ERRO
                    "quantidade": r[6]
                })

        # CORREÇÃO DAS PARADAS
        paradas = []
        for r in wb[ABA_PARADAS].iter_rows(min_row=2, values_only=True):
            if r[0] and r[1] == tid:
                paradas.append({
                    "equipamento": r[2], 
                    "causa": r[3], 
                    "categoria": r[4],
                    "hora_ini": str(r[5])[:5] if r[5] else "",
                    "hora_fim": str(r[6])[:5] if r[6] else "",
                    "duracao_min": r[7], 
                    "obs": r[8]
                })

    return jsonify({
        "encontrado": True, 
        "turno_id": tid,
        "turno":       tr[3] if tr else "",
        "encarregado": tr[4] if tr else "",
        "operador":    tr[5] if tr else "",
        "bal1":        tr[6] if tr else "",
        "bal2":        tr[7] if tr else "",
        "descarga":    descarga, 
        "paradas":     paradas,
    })
@app.route("/api/salvar", methods=["POST"])
def salvar_turno():
    try:
        p = request.get_json(silent=True)
        data, periodo = p.get("data"), p.get("periodo")
        with lock:
            garantir_banco()
            wb = load_workbook(EXCEL_PATH)
            ws_t, ws_d, ws_p = wb[ABA_TURNOS], wb[ABA_DESCARGA], wb[ABA_PARADAS]
            tid = buscar_turno_id(wb, data, periodo)
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if tid is None:
                tid = proximo_id(ws_t)
                ws_t.append([tid, data, periodo, p.get("turno",""), p.get("encarregado",""), p.get("operador",""), p.get("bal1",""), p.get("bal2",""), now])
            else:
                for row in ws_t.iter_rows(min_row=2):
                    if row[0].value == tid:
                        row[3].value, row[4].value, row[5].value = p.get("turno",""), p.get("encarregado",""), p.get("operador","")
                        row[6].value, row[7].value = p.get("bal1",""), p.get("bal2","")
                        break
            
            for ws, lista, cols in [(ws_d, p.get("descarga", []), 7), (ws_p, p.get("paradas", []), 9)]:
                linhas_manter = [[c.value for c in r] for r in ws.iter_rows(min_row=2) if r[0].value is not None and r[1].value != tid]
                if ws.max_row > 1: ws.delete_rows(2, ws.max_row)
                for l in linhas_manter: ws.append(l)
                base = proximo_id(ws) - 1
                for i, item in enumerate(lista, 1):
                    row_data = [base + i, tid] + [item.get(k, "") for k in (["equipamento","produto","hora_ini","hora_fim","quantidade"] if ws==ws_d else ["equipamento","causa","categoria","hora_ini","hora_fim","duracao_min","obs"])]
                    ws.append(row_data)
            _salvar_excel(wb)
        return jsonify({"ok": True, "turno_id": tid})
    except Exception as e:
        return jsonify({"erro": str(e)}), 500

@app.route("/api/dashboard")
def api_dashboard():
    try:
        data_ini  = request.args.get("data_ini", "")
        data_fim  = request.args.get("data_fim",  "")
        periodo_f_raw = request.args.get("periodo", "")
        moega_f   = request.args.get("moega", "")
        periodo_f = periodo_f_raw  # Excel armazena o valor original ("07X13"), sem conversão

        def _parse_hm(v):
            if v is None or v == "": return None
            if isinstance(v, _dt.time): return v.hour*60+v.minute
            if isinstance(v, _dt.datetime): return v.hour*60+v.minute
            if isinstance(v, _dt.timedelta): return int(v.total_seconds()/60)
            try:
                s = str(v).strip()
                if ":" in s:
                    h, m = s.split(":")[:2]
                    return int(h)*60+int(m)
            except: pass
            return None

        _PER_INI = {"p0":60,"p1":420,"p2":780,"p3":1140}
        _PER_FIM = {"p0":420,"p1":780,"p2":1140,"p3":1500}
        GRAOS    = {"Soja","Milho"}
        EQL      = {"Moega 1":"m1","Moega 2":"m2","m1":"m1","m2":"m2"}

        with lock:
            garantir_banco()
            wb   = load_workbook(EXCEL_PATH, data_only=True)
            ws_t = wb[ABA_TURNOS]
            ws_d = wb[ABA_DESCARGA]
            ws_p = wb[ABA_PARADAS]

            # Filtrar turnos
            turnos_filt = {}
            for row in ws_t.iter_rows(min_row=2, values_only=True):
                tid, data_raw, per = row[0], row[1], row[2]
                if not tid or not data_raw: continue
                dt_s = data_raw.strftime("%Y-%m-%d") if isinstance(data_raw, (_dt.datetime, _dt.date)) else str(data_raw)[:10]
                if data_ini and dt_s < data_ini: continue
                if data_fim  and dt_s > data_fim:  continue
                if periodo_f and per != periodo_f: continue
                turnos_filt[tid] = {"data": dt_s, "periodo": per}

            if not turnos_filt:
                return jsonify({"descarga":{},"paradas":[],"por_dia":[],"vagh":{},
                                "kpis":{},"num_dias":0,"num_turnos":0})

            # Ferro por turno: guarda intervalos (ini_m, fim_m) para intersectar com cada produto
            ferro_intervals = {}  # {tid: {eq_k: [(ini_m, fim_m), ...]}}
            for row in ws_p.iter_rows(min_row=2, values_only=True):
                tid, eq_raw, causa, cat, h_i, h_f, dur = row[1], row[2], row[3], row[4], row[5], row[6], row[7]
                if tid not in turnos_filt or cat != "ferro": continue
                eq_k_p = EQL.get(eq_raw, eq_raw)
                ini_p = _parse_hm(h_i)
                fim_p = _parse_hm(h_f)
                if ini_p is None or fim_p is None: continue
                if fim_p <= ini_p: fim_p += 1440
                if tid not in ferro_intervals: ferro_intervals[tid] = {}
                if eq_k_p not in ferro_intervals[tid]: ferro_intervals[tid][eq_k_p] = []
                ferro_intervals[tid][eq_k_p].append((ini_p, fim_p))

            def ferro_no_intervalo(tid, eq_k, prod_ini, prod_fim):
                if prod_fim <= prod_ini: prod_fim += 1440
                total = 0
                for (fi, ff) in ferro_intervals.get(tid, {}).get(eq_k, []):
                    overlap = min(ff, prod_fim) - max(fi, prod_ini)
                    if overlap > 0: total += overlap
                    if prod_fim > 1440:
                        ov2 = min(ff + 1440, prod_fim) - max(fi + 1440, prod_ini)
                        if ov2 > 0: total += ov2
                return total

            # Paradas
            paradas_agg = {}
            paradas_dia  = {}
            for row in ws_p.iter_rows(min_row=2, values_only=True):
                tid, eq_raw, causa, cat, h_i, h_f, dur = row[1], row[2], row[3], row[4], row[5], row[6], row[7]
                if tid not in turnos_filt or not causa: continue
                dur = dur or 0
                if moega_f and eq_raw != moega_f: continue
                if causa not in paradas_agg:
                    paradas_agg[causa] = {"total_min":0,"ferro_min":0,"term_min":0,
                                          "ocorrencias":0,"equipamentos":set(),"ocorr_list":[]}
                paradas_agg[causa]["total_min"]   += dur
                paradas_agg[causa]["ocorrencias"] += 1
                paradas_agg[causa]["equipamentos"].add(eq_raw)
                paradas_agg[causa]["ocorr_list"].append({
                    "data": turnos_filt[tid]["data"], "eq": eq_raw,
                    "hora_ini": str(h_i)[:5] if h_i else "",
                    "hora_fim":  str(h_f)[:5] if h_f else "",
                    "dur_min": dur,
                })
                if cat == "ferro": paradas_agg[causa]["ferro_min"] += dur
                else:              paradas_agg[causa]["term_min"]  += dur
                dt = turnos_filt[tid]["data"]
                if dt not in paradas_dia: paradas_dia[dt] = {"ferro_min":0,"term_min":0}
                if cat == "ferro": paradas_dia[dt]["ferro_min"] += dur
                else:              paradas_dia[dt]["term_min"]  += dur

            # Descarga + vag/hora
            # Taxa por produto: (qtd) / ((fim - ini) - ferro_no_intervalo) * 60
            descarga     = {"Soja":{"m1":0,"m2":0},"Milho":{"m1":0,"m2":0},"Farelo":{"m1":0,"m2":0}}
            descarga_dia = {}
            taxas_gbl    = {"m1":{"graos":{"qtd":0,"min":0},"farelo":{"qtd":0,"min":0}},
                          "m2":{"graos":{"qtd":0,"min":0},"farelo":{"qtd":0,"min":0}}}
            taxas_dia    = {}  # dt -> {eq_k: {tipo: {qtd, min}}}
            _desc_det    = {}  # (eq_k, prod) -> {eq, prod, qtd, intervalos:[{ini,fim}]}
            for row in ws_d.iter_rows(min_row=2, values_only=True):
                tid, eq_raw, prod, h_ini, h_fim, qtd = row[1], row[2], row[3], row[4], row[5], row[6]
                if tid not in turnos_filt: continue
                if not prod: continue
                if moega_f and eq_raw != moega_f: continue
                eq_k = EQL.get(eq_raw, "")
                if not eq_k: continue
                qtd = qtd or 0
                dt  = turnos_filt[tid]["data"]

                # Acumular descarga (para KPIs e por_dia)
                if prod in descarga:
                    descarga[prod][eq_k] += qtd
                if dt not in descarga_dia:
                    descarga_dia[dt] = {"Soja":0,"Milho":0,"Farelo":0,"total":0}
                if prod in descarga_dia[dt]:
                    descarga_dia[dt][prod] += qtd
                descarga_dia[dt]["total"] += qtd

                # Guardar detalhe de horários para cálculo de vagh no dashboard
                _desc_det_key = (eq_k, prod, tid)
                if _desc_det_key not in _desc_det:
                    _desc_det[_desc_det_key] = {"eq":eq_k,"prod":prod,"qtd":0,"intervalos":[]}
                _desc_det[_desc_det_key]["qtd"] += qtd
                hi_str = str(h_ini)[:5] if h_ini else ""
                hf_str = str(h_fim)[:5] if h_fim else ""
                if hi_str and hf_str:
                    _desc_det[_desc_det_key]["intervalos"].append({"ini":hi_str,"fim":hf_str})

                # Acumular para cálculo de taxa por produto
                ini_m = _parse_hm(h_ini)
                fim_m = _parse_hm(h_fim)
                if ini_m is None or fim_m is None or ini_m == fim_m:
                    continue  # sem horário não é possível calcular taxa
                if fim_m < ini_m: fim_m += 1440

                # Ferro apenas dentro do intervalo deste produto
                f_prod = ferro_no_intervalo(tid, eq_k, ini_m, fim_m)
                tempo_liq_min = (fim_m - ini_m) - f_prod
                if tempo_liq_min <= 0 or qtd <= 0:
                    continue

                tipo = "graos" if prod in GRAOS else "farelo"

                taxas_gbl[eq_k][tipo]["qtd"] += qtd
                taxas_gbl[eq_k][tipo]["min"] += tempo_liq_min
                if dt not in taxas_dia:
                    taxas_dia[dt] = {"m1":{"graos":{"qtd":0,"min":0},"farelo":{"qtd":0,"min":0}},
                                     "m2":{"graos":{"qtd":0,"min":0},"farelo":{"qtd":0,"min":0}}}
                taxas_dia[dt][eq_k][tipo]["qtd"] += qtd
                taxas_dia[dt][eq_k][tipo]["min"] += tempo_liq_min

        # KPIs e serialização (fora do lock — só operações em memória)
        def taxa_pond(acc): return round(acc["qtd"] / (acc["min"] / 60.0), 2) if acc["qtd"] > 0 and acc["min"] > 0 else 0

        ts = descarga["Soja"]["m1"]   + descarga["Soja"]["m2"]
        tm = descarga["Milho"]["m1"]  + descarga["Milho"]["m2"]
        tf = descarga["Farelo"]["m1"] + descarga["Farelo"]["m2"]
        tg = ts + tm + tf
        t_ferro = sum(v["ferro_min"] for v in paradas_agg.values())
        t_term  = sum(v["term_min"]  for v in paradas_agg.values())

        vagh = {
            "m1": {"graos": taxa_pond(taxas_gbl["m1"]["graos"]),
                   "farelo": taxa_pond(taxas_gbl["m1"]["farelo"])},
            "m2": {"graos": taxa_pond(taxas_gbl["m2"]["graos"]),
                   "farelo": taxa_pond(taxas_gbl["m2"]["farelo"])},
        }

        por_dia = []
        for dt in sorted(set(descarga_dia) | set(paradas_dia)):
            d   = descarga_dia.get(dt, {})
            p2  = paradas_dia.get(dt, {})
            vd  = taxas_dia.get(dt, {})
            por_dia.append({
                "data":   dt,
                "Soja":   round(d.get("Soja",0),1),
                "Milho":  round(d.get("Milho",0),1),
                "Farelo": round(d.get("Farelo",0),1),
                "total":  round(d.get("total",0),1),
                "ferro_min":     p2.get("ferro_min",0),
                "term_min":      p2.get("term_min",0),
                "vagh_m1_graos":  taxa_pond(vd.get("m1",{}).get("graos",{"qtd":0,"min":0})),
                "vagh_m1_farelo": taxa_pond(vd.get("m1",{}).get("farelo",{"qtd":0,"min":0})),
                "vagh_m2_graos":  taxa_pond(vd.get("m2",{}).get("graos",{"qtd":0,"min":0})),
                "vagh_m2_farelo": taxa_pond(vd.get("m2",{}).get("farelo",{"qtd":0,"min":0})),
            })

        paradas_list = sorted([
            {"causa":c, "equipamentos":list(v["equipamentos"]),
             "total_min":v["total_min"], "ferro_min":v["ferro_min"],
             "term_min":v["term_min"],   "ocorrencias":v["ocorrencias"],
             "ocorr_list":sorted(v["ocorr_list"], key=lambda x:x["data"])}
            for c,v in paradas_agg.items()
        ], key=lambda x: x["total_min"], reverse=True)

        return jsonify({
            "_v": 2,
            "descarga":        descarga,
            "descarga_detalhe": list(_desc_det.values()),
            "paradas":         paradas_list,
            "por_dia":         por_dia,
            "vagh":            vagh,
            "num_dias":        len(set(t["data"] for t in turnos_filt.values())),
            "num_turnos":      len(turnos_filt),
            "kpis": {
                "total_geral":  round(tg, 1),
                "total_soja":   round(ts, 1),
                "total_milho":  round(tm, 1),
                "total_farelo": round(tf, 1),
                "ferro_min":    t_ferro,
                "term_min":     t_term,
            },
        })

    except Exception as e:
        print(traceback.format_exc())
        return jsonify({"erro": str(e), "trace": traceback.format_exc()}), 500


if __name__ == "__main__":
    garantir_banco()
    threading.Thread(target=lambda: (__import__('time').sleep(1), webbrowser.open(f"http://localhost:{PORT}")), daemon=True).start()
    app.run(host="0.0.0.0", port=PORT, debug=False)