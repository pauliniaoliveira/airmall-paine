#!/usr/bin/env python3
"""
Airmall 2.0 — Gerador de Relatório Web
=======================================
Use este script sempre que tiver um novo arquivo Excel de dados.

Uso:
  python3 gerar_relatorio.py <caminho_do_excel>

Exemplo:
  python3 gerar_relatorio.py "Airmall 2.0 - integridade vendas-novo.xlsx"

O script gera/atualiza o arquivo index.html na mesma pasta.
Depois é só fazer upload da pasta 'airmall' no Netlify.
"""
import sys, json, datetime, re
import pandas as pd
import numpy as np

# ── Args ────────────────────────────────────────────────────────────────────
if len(sys.argv) < 2:
    print("Uso: python3 gerar_relatorio.py <arquivo.xlsx>")
    sys.exit(1)

EXCEL_PATH = sys.argv[1]
SCRIPT_DIR = __import__('os').path.dirname(__import__('os').path.abspath(__file__))
HTML_PATH  = __import__('os').path.join(SCRIPT_DIR, "index.html")

print(f"📂 Lendo: {EXCEL_PATH}")
df = pd.read_excel(EXCEL_PATH, sheet_name='Resultado da consulta')
df.columns = df.columns.str.strip()

# ── Campos base ─────────────────────────────────────────────────────────────
df['date'] = pd.to_datetime(df['created_at']).dt.date
for c in ['gmv','fee_service','provider_incentive','take_rate','profit']:
    df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

def cia_grp(c):
    c = str(c).upper()
    if 'LA' in c or 'TAM' in c or 'JJ' in c: return 'LATAM'
    if 'G3' in c or 'GOL' in c or 'VRG' in c: return 'GOL'
    if 'AD' in c or 'AZUL' in c: return 'AZUL'
    return 'Outros'

df['cia_group'] = df['outbound_cia'].apply(cia_grp)

def cred_tipo(c):
    c = str(c).lower()
    if 'nunca' in c or 'proib' in c: return 'PROIBIDA'
    if 'test' in c: return 'TESTE'
    if 'busca' in c: return 'BUSCA'
    if 'cota' in c or 'offline' in c or 'fatur' in c: return 'COTAÇÃO'
    return 'OUTRO'

df['cred_tipo'] = df['credential'].apply(cred_tipo)

TODAY    = df['date'].max()
SAME_DLW = TODAY - datetime.timedelta(days=7)
START_7D = TODAY - datetime.timedelta(days=6)

# ── KPIs ────────────────────────────────────────────────────────────────────
def day_kpis(dfd):
    b = dfd[dfd['original']=='Sim']
    e = dfd[dfd['original']=='Não']
    n = len(dfd)
    tr_e = e['take_rate'].mean() if len(e)>0 else np.nan
    tr_b = b['take_rate'].mean() if len(b)>0 else np.nan
    return {
        'total': n,
        'cnt_esp': int(len(e)),
        'pct_esp': round(len(e)/n*100,1) if n>0 else 0,
        'gmv': round(e['gmv'].sum(),2),
        'receita': round((e['fee_service']+e['provider_incentive']).sum(),2),
        'take_rate': round(tr_e*100,2) if not np.isnan(tr_e) else None,
        'take_rate_pub': round(tr_b*100,2) if not np.isnan(tr_b) else None,
        'delta_tr': round((tr_e-tr_b)*100,2) if not np.isnan(tr_e) and not np.isnan(tr_b) else None,
        'rec_incr': round((e['fee_service']+e['provider_incentive']).sum()-(b['fee_service']+b['provider_incentive']).sum(),2) if len(b)>0 else 0,
        'incentivo_excl': round(e['provider_incentive'].sum()-b['provider_incentive'].sum(),2) if len(b)>0 else 0,
    }

med7_raw = day_kpis(df[df['date']>=START_7D])
media7d  = {k: round(v/7,2) if isinstance(v,(int,float)) and not isinstance(v,bool) and v is not None else v for k,v in med7_raw.items()}

painel = {
    'hoje':    day_kpis(df[df['date']==TODAY]),
    'sem_ant': day_kpis(df[df['date']==SAME_DLW]),
    'media7d': media7d,
    'periodo': day_kpis(df),
}

# ── Diário ──────────────────────────────────────────────────────────────────
daily = []
for d in sorted(df['date'].unique()):
    dfd = df[df['date']==d]
    b = dfd[dfd['original']=='Sim']; e = dfd[dfd['original']=='Não']
    tr_e = e['take_rate'].mean() if len(e)>0 else np.nan
    tr_b = b['take_rate'].mean() if len(b)>0 else np.nan
    daily.append({
        'date': str(d), 'total': len(dfd),
        'pct_esp': round(len(e)/len(dfd)*100,1) if len(dfd)>0 else 0,
        'gmv': round(e['gmv'].sum(),2),
        'receita': round((e['fee_service']+e['provider_incentive']).sum(),2),
        'delta_tr': round((tr_e-tr_b)*100,2) if not np.isnan(tr_e) and not np.isnan(tr_b) else None,
    })

# ── Por CIA ─────────────────────────────────────────────────────────────────
cias = ['LATAM','GOL','AZUL','Outros']
por_cia = {cia: day_kpis(df[df['cia_group']==cia]) for cia in cias}

# ── Busca vs Emissão por CIA ─────────────────────────────────────────────────
bvse_cia = []
for cia in cias:
    dc = df[df['cia_group']==cia]
    b = dc[dc['original']=='Sim']; e = dc[dc['original']=='Não']
    bvse_cia.append({
        'cia': cia, 'total': len(dc),
        'pct_esp': round(len(e)/len(dc)*100,1) if len(dc)>0 else 0,
        'gmv_esp': round(e['gmv'].sum(),2),
        'tr_pub': round(b['take_rate'].mean()*100,2) if len(b)>0 else None,
        'tr_esp': round(e['take_rate'].mean()*100,2) if len(e)>0 else None,
        'delta_tr': round((e['take_rate'].mean()-b['take_rate'].mean())*100,2) if len(b)>0 and len(e)>0 else None,
    })

# ── Provider × CIA ───────────────────────────────────────────────────────────
prov_cia = []
for prov in sorted(df['provider_name'].unique()):
    row = {'provider': prov}
    for cia in cias:
        sub = df[(df['provider_name']==prov)&(df['cia_group']==cia)]
        b = sub[sub['original']=='Sim']; e = sub[sub['original']=='Não']
        row[cia+'_cnt'] = len(sub)
        row[cia+'_pct_esp'] = round(len(e)/len(sub)*100,1) if len(sub)>0 else 0
        row[cia+'_delta_tr'] = round((e['take_rate'].mean()-b['take_rate'].mean())*100,2) if len(b)>0 and len(e)>0 else None
    prov_cia.append(row)

# ── Hoje Busca ───────────────────────────────────────────────────────────────
hoje_busca = []
for cia in cias:
    tc = len(df[(df['date']==TODAY)&(df['cia_group']==cia)&(df['original']=='Sim')])
    for prov, grp in df[(df['date']==TODAY)&(df['cia_group']==cia)&(df['original']=='Sim')].groupby('provider_name'):
        hoje_busca.append({'cia':cia,'provider':prov,'qtd':len(grp),'pct_cia':round(len(grp)/tc*100,1) if tc>0 else 0,'tr':round(grp['take_rate'].mean()*100,2)})

# ── Hoje Emissão ─────────────────────────────────────────────────────────────
hoje_emi = []
for cia in cias:
    tc = len(df[(df['date']==TODAY)&(df['cia_group']==cia)])
    for prov, grp in df[(df['date']==TODAY)&(df['cia_group']==cia)].groupby('provider_name'):
        for orig, sgrp in grp.groupby('original'):
            tipo = 'ESPECIAL' if orig=='Não' else 'PÚBLICA'
            hoje_emi.append({'cia':cia,'provider':prov,'tipo':tipo,'qtd':len(sgrp),'pct_cia':round(len(sgrp)/tc*100,1) if tc>0 else 0,'tr':round(sgrp['take_rate'].mean()*100,2)})

# ── Credenciais ──────────────────────────────────────────────────────────────
cred_data = []
for (prov,cred,tipo), grp in df.groupby(['provider_name','credential','cred_tipo']):
    b = grp[grp['original']=='Sim']; e = grp[grp['original']=='Não']
    cred_data.append({
        'provider':prov,'credential':cred,'tipo':tipo,'total':len(grp),
        'pct_esp': round(len(e)/len(grp)*100,1) if len(grp)>0 else 0,
        'delta_tr': round((e['take_rate'].mean()-b['take_rate'].mean())*100,2) if len(b)>0 and len(e)>0 else None,
        'alerta': tipo in ('PROIBIDA','TESTE') or (tipo=='BUSCA' and len(e)>0),
    })

# ── ADVPs ─────────────────────────────────────────────────────────────────────
advp_data = []
for advp, grp in df.groupby('outbound_advp'):
    b = grp[grp['original']=='Sim']; e = grp[grp['original']=='Não']
    advp_data.append({'advp':str(advp),'total':len(grp),'pct_esp':round(len(e)/len(grp)*100,1) if len(grp)>0 else 0,'gmv':round(grp['gmv'].sum(),2),'delta_tr':round((e['take_rate'].mean()-b['take_rate'].mean())*100,2) if len(b)>0 and len(e)>0 else None})
advp_data.sort(key=lambda x: x['total'], reverse=True)

# ── JSON final ───────────────────────────────────────────────────────────────
data = {
    'meta': {'today':str(TODAY),'same_dlw':str(SAME_DLW),'period_start':str(df['date'].min()),'period_end':str(TODAY),'generated_at':datetime.datetime.now().strftime('%d/%m/%Y %H:%M')},
    'painel': painel, 'daily': daily, 'por_cia': por_cia, 'bvse_cia': bvse_cia,
    'prov_cia': prov_cia, 'hoje_busca': hoje_busca, 'hoje_emi': hoje_emi,
    'cred_data': cred_data, 'advp_data': advp_data[:60], 'cias': cias,
}

json_str = json.dumps(data, ensure_ascii=False, default=str)

# ── Injetar no HTML ───────────────────────────────────────────────────────────
with open(HTML_PATH, 'r', encoding='utf-8') as f:
    html = f.read()

# Substitui o bloco de dados entre const D = { ... };
html_new = re.sub(r'const D = \{.*?\};', f'const D = {json_str};', html, flags=re.DOTALL)

with open(HTML_PATH, 'w', encoding='utf-8') as f:
    f.write(html_new)

print(f"✅ Relatório atualizado: {HTML_PATH}")
print(f"   Período: {df['date'].min()} → {TODAY}")
print(f"   Total emissões: {len(df):,}")
print(f"\n👉 Próximos passos:")
print("   1. Acesse https://netlify.com/drop")
print("   2. Arraste a pasta 'airmall' para a área de upload")
print("   3. Compartilhe o link gerado com a equipe!")
