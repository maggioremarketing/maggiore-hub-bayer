#!/usr/bin/env python3
"""
HUB Sync Script — Atualiza o Consumer Insights Hub a partir do HUB_control.xlsx

Uso:
    python sync_hub.py

Requisitos:
    pip install pandas openpyxl

O script lê o arquivo HUB_control.xlsx (mesma pasta) e atualiza:
  - learningJourneyMap (links dos Learning Journeys)
  - Status de cancelled para projetos sem link
  - Projetos de SL (links dos reports)
"""

import re, json, os, sys
import pandas as pd

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
HTML_PATH = os.path.join(SCRIPT_DIR, "bayer-consumer-insights-hub.html")
XLS_PATH = os.path.join(SCRIPT_DIR, "HUB_control.xlsx")

def main():
    if not os.path.exists(XLS_PATH):
        print(f"ERRO: {XLS_PATH} não encontrado"); sys.exit(1)
    if not os.path.exists(HTML_PATH):
        print(f"ERRO: {HTML_PATH} não encontrado"); sys.exit(1)

    html = open(HTML_PATH, "r", encoding="utf-8").read()
    print(f"HTML carregado: {len(html):,} bytes")

    # ── IEP: Ler planilha ──
    df_iep = pd.read_excel(XLS_PATH, sheet_name="IEP")
    print(f"IEP: {len(df_iep)} linhas")

    brand_to_html = {"MiraLAX": "Miralax", "Shopper": "Shoppers"}
    new_lj = {}
    cancelled_keys = set()

    for _, row in df_iep.iterrows():
        brand_raw = str(row.get('Brand', '')).strip()
        exp = str(row.get('Experiment', '')).strip()
        link1 = str(row.get('Link 1', '')).strip() if pd.notna(row.get('Link 1')) else ''
        link2 = str(row.get('Link 2', '')).strip() if pd.notna(row.get('Link 2')) else ''
        status = str(row.get('Status', '')).strip().lower()

        if brand_raw == 'nan' or exp == 'nan':
            continue

        brand_html = brand_to_html.get(brand_raw, brand_raw)
        key = f"{brand_html}|{exp}"

        if status == 'cancelled' or (not link1 and not link2):
            cancelled_keys.add(key)
            continue

        ljs = []
        proj_name = str(row.get('Project Name', '')).strip()
        if link1 and link1 != 'cancelled':
            ljs.append({"file": f"{exp} - {proj_name}.pptx", "url": link1})
        if link2 and link2 != 'nan':
            ljs.append({"file": f"{exp} - {proj_name} (2).pptx", "url": link2})
        if ljs:
            new_lj[key] = ljs

    # Atualizar learningJourneyMap
    m = re.search(r'const learningJourneyMap = \{[^;]+\};', html, re.DOTALL)
    if m:
        html = html[:m.start()] + f"const learningJourneyMap = {json.dumps(new_lj, ensure_ascii=False)};" + html[m.end():]
        print(f"  learningJourneyMap: {len(new_lj)} entradas")

    # Atualizar rounds=0 para cancelled
    m2 = re.search(r'const iepProjects\s*=\s*(\[.*?\]);', html, re.DOTALL)
    if m2:
        iep = json.loads(m2.group(1))
        changes = 0
        for p in iep:
            key = f"{p['brand']}|{p['experiment_id']}"
            if key in cancelled_keys and p['rounds'] > 0:
                p['rounds'] = 0
                changes += 1
            elif key in new_lj and p['rounds'] == 0:
                p['rounds'] = 1
                changes += 1
        html = html[:m2.start()] + f"const iepProjects = {json.dumps(iep, ensure_ascii=False)};" + html[m2.end():]
        print(f"  iepProjects: {changes} alterações de rounds")

    # ── SL: Ler planilha ──
    df_sl = pd.read_excel(XLS_PATH, sheet_name="SL")
    print(f"SL: {len(df_sl)} linhas")

    m3 = re.search(r'const projects\s*=\s*(\[.*?\]);', html, re.DOTALL)
    if m3:
        sl_current = json.loads(m3.group(1))
        # Atualizar links dos reports existentes
        link_updates = 0
        for _, row in df_sl.iterrows():
            link = str(row.get('Link', '')).strip() if pd.notna(row.get('Link')) else ''
            title = str(row.get('Report Title', '')).strip()
            brand = str(row.get('Brand', '')).strip()
            year = row.get('Year', '')
            if not link or not title:
                continue
            for p in sl_current:
                if p.get('theme') == title and p.get('brand') == brand and p.get('year') == year:
                    if p.get('link') != link:
                        p['link'] = link
                        link_updates += 1
                    break
        html = html[:m3.start()] + f"const projects = {json.dumps(sl_current, ensure_ascii=False)};" + html[m3.end():]
        print(f"  SL links atualizados: {link_updates}")

    # Salvar
    open(HTML_PATH, "w", encoding="utf-8").write(html)
    print(f"\nHTML salvo: {len(html):,} bytes")
    print("Sync completo!")

if __name__ == "__main__":
    main()
