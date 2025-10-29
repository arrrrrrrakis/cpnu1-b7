# -*- coding: utf-8 -*-
"""
Arquivo: alocacao_v5_DA.py
Descrição: Implementação do algoritmo de Aceitação Diferida (Deferred Acceptance – Gale–Shapley)
para alocação muitos-para-um com cotas por tipo (PI, PCD, PPP, AMPLA), compatível com Google Colab.

Entradas:
  - Planilha Excel 'b7p.xlsx' (aba 'lista'), com as mesmas colunas do seu script original.
Saídas:
  - CSV/XLSX com detalhes, resumo por cota, contagem, não alocados e auditoria de empates.

Uso (no Google Colab):
  1) Faça upload deste arquivo e do 'b7p.xlsx'.
  2) from alocacao_v5_DA import main; main()
     (ou simplesmente execute o arquivo como script).
"""
import pandas as pd, numpy as np
from collections import defaultdict
from pathlib import Path

# ---- Colab I/O ----
try:
    import google.colab.files as gfiles
    IN_COLAB = True
except Exception:
    IN_COLAB = False

def main():
    if IN_COLAB:
        print("Por favor, faça o upload do arquivo 'b7p.xlsx'")
        uploaded = gfiles.upload()
        if 'b7p.xlsx' not in uploaded:
            raise SystemExit("Arquivo 'b7p.xlsx' não encontrado. Renomeie e tente de novo.")
        ARQUIVO = 'b7p.xlsx'
        OUT_DIR = Path('/content')
    else:
        ARQUIVO = 'b7p.xlsx'   # ajuste se necessário
        OUT_DIR = Path('.')

    ABA = 'lista'
    # Saídas (sufixo v5_DA)
    DETALHES_CSV        = OUT_DIR/'detalhes_alocacao_v5_DA.csv'
    XLSX_PATH           = OUT_DIR/'resultado_alocacao_v5_DA.xlsx'
    NAO_CSV             = OUT_DIR/'nao_alocados_v5_DA.csv'
    AUDIT_TIE_CSV       = OUT_DIR/'auditoria_desempates_v5_DA.csv'
    CLASSIF_INICIAL_CSV = OUT_DIR/'prioridade_por_cargo_v5_DA.csv'

    # ---- Configuração de vagas por cargo (copiada do seu script) ----
    VAGAS_POR_CARGO = {
        ('AGU','B7-01-A'):{'AMPLA':72,'PPP':15,'PCD':6,'PI':0,'TOTAL':93},
        ('AGU','B7-01-B'):{'AMPLA':38,'PPP':7,'PCD':5,'PI':0,'TOTAL':50},
        ('AGU','B7-01-C'):{'AMPLA':1,'PPP':0,'PCD':0,'PI':0,'TOTAL':1},
        ('AGU','B7-01-D'):{'AMPLA':24,'PPP':6,'PCD':2,'PI':0,'TOTAL':32},
        ('AGU','B7-01-E'):{'AMPLA':1,'PPP':0,'PCD':0,'PI':0,'TOTAL':1},
        ('FUNAI','B7-02-A'):{'AMPLA':26,'PPP':14,'PCD':7,'PI':20,'TOTAL':67},
        ('FUNAI','B7-02-B'):{'AMPLA':2,'PPP':0,'PCD':0,'PI':0,'TOTAL':2},
        ('FUNAI','B7-02-C'):{'AMPLA':6,'PPP':3,'PCD':3,'PI':5,'TOTAL':17},
        ('FUNAI','B7-02-D'):{'AMPLA':12,'PPP':6,'PCD':3,'PI':7,'TOTAL':28},
        ('FUNAI','B7-02-E'):{'AMPLA':16,'PPP':7,'PCD':2,'PI':9,'TOTAL':34},
        ('IBGE','B7-03-A'):{'AMPLA':13,'PPP':2,'PCD':1,'PI':0,'TOTAL':16},
        ('IBGE','B7-03-B'):{'AMPLA':1,'PPP':0,'PCD':1,'PI':0,'TOTAL':2},
        ('IBGE','B7-03-C'):{'AMPLA':3,'PPP':1,'PCD':0,'PI':0,'TOTAL':4},
        ('IBGE','B7-03-D'):{'AMPLA':73,'PPP':19,'PCD':7,'PI':0,'TOTAL':99},
        ('IBGE','B7-03-E'):{'AMPLA':3,'PPP':0,'PCD':0,'PI':0,'TOTAL':3},
        ('IBGE','B7-03-F'):{'AMPLA':1,'PPP':0,'PCD':0,'PI':0,'TOTAL':1},
        ('IBGE','B7-03-G'):{'AMPLA':14,'PPP':3,'PCD':2,'PI':0,'TOTAL':19},
        ('IBGE','B7-03-H'):{'AMPLA':3,'PPP':0,'PCD':0,'PI':0,'TOTAL':3},
        ('IBGE','B7-03-I'):{'AMPLA':25,'PPP':6,'PCD':3,'PI':0,'TOTAL':34},
        ('IBGE','B7-03-J'):{'AMPLA':6,'PPP':1,'PCD':1,'PI':0,'TOTAL':8},
        ('INCRA','B7-04-A'):{'AMPLA':7,'PPP':1,'PCD':1,'PI':0,'TOTAL':9},
        ('INCRA','B7-04-B'):{'AMPLA':14,'PPP':3,'PCD':1,'PI':0,'TOTAL':18},
        ('INCRA','B7-04-C'):{'AMPLA':43,'PPP':10,'PCD':3,'PI':0,'TOTAL':56},
        ('INCRA','B7-04-D'):{'AMPLA':118,'PPP':27,'PCD':7,'PI':0,'TOTAL':152},
        ('INEP','B7-05-A'):{'AMPLA':9,'PPP':2,'PCD':0,'PI':0,'TOTAL':11},
        ('MAPA','B7-06-A'):{'AMPLA':10,'PPP':1,'PCD':0,'PI':0,'TOTAL':11},
        ('MCTI','B7-07-A'):{'AMPLA':1,'PPP':1,'PCD':0,'PI':0,'TOTAL':2},
        ('MCTI','B7-07-B'):{'AMPLA':3,'PPP':0,'PCD':0,'PI':0,'TOTAL':3},
        ('MCTI','B7-07-C'):{'AMPLA':3,'PPP':2,'PCD':0,'PI':0,'TOTAL':5},
        ('MCTI','B7-07-D'):{'AMPLA':1,'PPP':1,'PCD':1,'PI':0,'TOTAL':3},
        ('MCTI','B7-07-E'):{'AMPLA':52,'PPP':7,'PCD':2,'PI':0,'TOTAL':61},
        ('MINC','B7-08-A'):{'AMPLA':56,'PPP':16,'PCD':5,'PI':0,'TOTAL':77},
        ('MGI','B7-09-A'):{'AMPLA':147,'PPP':38,'PCD':12,'PI':0,'TOTAL':197},
        ('MGI','B7-09-B'):{'AMPLA':4,'PPP':2,'PCD':1,'PI':0,'TOTAL':7},
        ('MGI','B7-09-C'):{'AMPLA':2,'PPP':0,'PCD':0,'PI':0,'TOTAL':2},
        ('MGI','B7-09-D'):{'AMPLA':4,'PPP':1,'PCD':0,'PI':0,'TOTAL':5},
        ('MGI','B7-09-F'):{'AMPLA':8,'PPP':2,'PCD':1,'PI':0,'TOTAL':11},
        ('MJSP','B7-10-A'):{'AMPLA':110,'PPP':23,'PCD':10,'PI':0,'TOTAL':143},
        ('MS','B7-11-A'):{'AMPLA':1,'PPP':2,'PCD':1,'PI':0,'TOTAL':4},
        ('MDIC','B7-12-A'):{'AMPLA':51,'PPP':15,'PCD':5,'PI':0,'TOTAL':71},
        ('MPO','B7-13-A'):{'AMPLA':37,'PPP':9,'PCD':3,'PI':0,'TOTAL':49},
        ('PREVIC','B7-14-A'):{'AMPLA':3,'PPP':0,'PCD':1,'PI':0,'TOTAL':4},
    }

    # ---- Carregar e normalizar ----
    xl = pd.ExcelFile(ARQUIVO)
    df = xl.parse(ABA)
    df.columns = (df.columns.astype(str)
                  .str.replace('\n',' ', regex=False)
                  .str.replace('\r',' ', regex=False)
                  .str.strip())

    def first_non_empty(row):
        for x in row:
            if pd.isna(x):
                continue
            s = str(x).strip()
            if s and s.lower()!='nan' and s!='-':
                return s
        return ''

    for keyword, newname in [('Indígena','Class_Indígena'),
                             ('PCD','Class_PCD2'),
                             ('Negra','Class_Negra3')]:
        cols = [c for c in df.columns if keyword.lower() in c.lower()]
        df[newname] = df[cols].apply(first_non_empty, axis=1) if cols else ''

    df['Interesse_norm'] = df.get('Interesse?', 'sim').astype(str).str.strip().str.lower()

    for col in ['Órgão','Cód. Cargo Edital','Cargo','Nome']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    df['Nota Final']    = pd.to_numeric(df['Nota Final'].astype(str).str.replace(',','.',regex=False).str.replace(' ','',regex=False), errors='coerce')
    df['Ordem_Pref']    = pd.to_numeric(df.get('Ordem_Pref'), errors='coerce')
    df['Posição Real']  = pd.to_numeric(df.get('Posição Real'), errors='coerce')
    df['CandID']        = df.get('Inscrição', df.get('Nome')).astype(str)
    df['Chave']         = list(zip(df['Órgão'].astype(str), df['Cód. Cargo Edital'].astype(str)))

    # Universo válido
    val = df[(df['Interesse_norm']=='sim') &
             (df['Nota Final'].notna()) &
             (df['Ordem_Pref'].notna()) &
             (df['Chave'].isin(VAGAS_POR_CARGO.keys()))].copy()

    # Cota do candidato (etiqueta)
    val['Cota'] = np.where(val['Class_Indígena'].astype(str).str.strip()!='', 'PI',
                   np.where(val['Class_PCD2'].astype(str).str.strip()!='', 'PCD',
                   np.where(val['Class_Negra3'].astype(str).str.strip()!='', 'PPP', 'AMPLA')))

    # Prioridade do cargo (determinística)
    val['rank_tuple'] = list(zip(-val['Nota Final'], val['Posição Real'].fillna(1e9), val['CandID']))

    # Preferências do candidato (lista ordenada por Ordem_Pref)
    prefs_cand = {}
    for cid, g in val.groupby('CandID'):
        prefs_cand[cid] = list(g.sort_values('Ordem_Pref')[['Órgão','Cód. Cargo Edital']].apply(tuple, axis=1))

    # Prioridade dos cargos e índice rápido
    prioridade_cargo = {}
    prior_index      = {}
    for chave, g in val.groupby('Chave'):
        orden = g.sort_values('rank_tuple')['CandID'].tolist()
        prioridade_cargo[chave] = orden
        prior_index[chave] = {c:i for i,c in enumerate(orden)}

    # Capacidades por cota
    cap = {ch: {'AMPLA':cfg['AMPLA'], 'PPP':cfg['PPP'], 'PCD':cfg['PCD'], 'PI':cfg['PI']}
           for ch, cfg in VAGAS_POR_CARGO.items()}

    # Mapa cota do candidato
    cand_cota_map = (val.sort_values('Ordem_Pref')
                       .groupby('CandID')['Cota'].first().to_dict())

    # Info detalhada por (CandID, Chave)
    DETAIL = {}
    for _, row in val.iterrows():
        DETAIL[(row['CandID'], row['Chave'])] = {
            'Nome': row['Nome'],
            'Inscrição': row.get('Inscrição', None),
            'Nota Final': row['Nota Final'],
            'Ordem_Pref': row['Ordem_Pref'],
            'Cargo': row['Cargo'],
            'rank_tuple': row['rank_tuple'],
            'Cota': row['Cota']
        }

    # Exporta prioridade inicial por cargo (auditoria)
    (pd.DataFrame([
        {'Orgao': org, 'Codigo_Cargo': cod, 'Pos': i+1, 'CandID': cid}
        for (org,cod), ordem in prioridade_cargo.items()
        for i, cid in enumerate(ordem)
    ])).to_csv(CLASSIF_INICIAL_CSV, index=False)

    # ---- DA ----
    aceitos = {ch: {'AMPLA':[], 'PPP':[], 'PCD':[], 'PI':[]} for ch in VAGAS_POR_CARGO.keys()}
    proxima = {cid:0 for cid in prefs_cand.keys()}
    livres  = set(prefs_cand.keys())

    def cotas_do_candidato(cid):
        # Política: cotista tenta AMPLA primeiro; se não couber, tenta sua cota
        c = cand_cota_map.get(cid, 'AMPLA')
        if c == 'PI':    return ['AMPLA', 'PI']
        if c == 'PCD':   return ['AMPLA', 'PCD']
        if c == 'PPP':   return ['AMPLA', 'PPP']
        return ['AMPLA']

    def inserir_ordenado(chave, cota, cid):
        lst = aceitos[chave][cota]
        if cid in lst:
            return None
        lst.append(cid)
        idxs = prior_index.get(chave, {})
        lst.sort(key=lambda x: idxs.get(x, 10**9))
        if len(lst) > cap[chave][cota]:
            rejeitado = lst.pop()
            return rejeitado
        return None

    total_propostas = 0
    while True:
        progresso = False
        for cid in list(livres):
            prefs = prefs_cand.get(cid, [])
            i = proxima[cid]
            if i >= len(prefs):
                livres.discard(cid)
                continue
            chave = prefs[i]
            proxima[cid] += 1
            total_propostas += 1
            progresso = True
            aceito_aqui = False
            for cota in cotas_do_candidato(cid):
                if cap[chave][cota] <= 0:
                    continue
                r = inserir_ordenado(chave, cota, cid)
                if r is None:
                    livres.discard(cid); aceito_aqui = True; break
                elif r != cid:
                    livres.add(r); livres.discard(cid); aceito_aqui = True; break
            # se não aceito em nenhuma cota, permanece livre para propor ao próximo cargo
        if not progresso:
            break

    # ---- Construir saída ----
    final_rows = []
    for (org, cod), cotas in aceitos.items():
        for cota_usada, lst in cotas.items():
            for cid in lst:
                det = DETAIL.get((cid, (org,cod)))
                if det is None:
                    # fallback: construir det mínimo
                    det = {'Cargo':'', 'Nome':'', 'Inscrição':None, 'Nota Final':None, 'Ordem_Pref':None, 'rank_tuple':None}
                final_rows.append({
                    'Orgao': org,
                    'Codigo_Cargo': cod,
                    'Cargo': det.get('Cargo',''),
                    'CandID': cid,
                    'Nome': det.get('Nome',''),
                    'Inscricao': det.get('Inscrição', None),
                    'Nota_Final': det.get('Nota Final'),
                    'Preferencia': det.get('Ordem_Pref'),
                    'Tipo_Cota_Candidato': cand_cota_map.get(cid, 'AMPLA'),
                    'Tipo_Vaga_Usada': cota_usada,
                    'rank_tuple': det.get('rank_tuple'),
                    'Destino': 'será nomeado'
                })

    final_df = pd.DataFrame(final_rows)

    # ---- Salvar ----
    final_df.to_csv(DETALHES_CSV, index=False)
    with pd.ExcelWriter(XLSX_PATH, engine='openpyxl') as writer:
        final_df.to_excel(writer, sheet_name='Detalhes v5_DA', index=False)
        # Resumo por cargo vs configuração
        resumo = []
        for (org, cod), grp in final_df.groupby(['Orgao','Codigo_Cargo']):
            cfg = VAGAS_POR_CARGO[(org, cod)]
            vc  = grp['Tipo_Vaga_Usada'].value_counts().to_dict()
            pre = {k: vc.get(k,0) for k in ['AMPLA','PPP','PCD','PI']}
            resumo.append({
                'Órgão': org, 'Código Cargo': cod,
                'Vagas AMPLA': cfg['AMPLA'], 'Vagas PPP': cfg['PPP'], 'Vagas PCD': cfg['PCD'], 'Vagas PI': cfg['PI'], 'Total Vagas': cfg['TOTAL'],
                'Preenchidas AMPLA': pre['AMPLA'], 'Preenchidas PPP': pre['PPP'], 'Preenchidas PCD': pre['PCD'], 'Preenchidas PI': pre['PI'],
                'Total Preenchidas (calc)': sum(pre.values())
            })
        pd.DataFrame(resumo).to_excel(writer, sheet_name='Resumo Vagas', index=False)
        (final_df.groupby(['Orgao','Codigo_Cargo','Tipo_Vaga_Usada']).size().reset_index(name='Qtde')
         ).to_excel(writer, sheet_name='Contagem por Cota', index=False)

    # Não alocados
    all_cands   = set(val['CandID'])
    alloc_cands = set(final_df['CandID'])
    nao_cids    = sorted(list(all_cands - alloc_cands))
    nao_rows = []
    for cid in nao_cids:
        g = val[val['CandID']==cid].sort_values('Ordem_Pref')
        nao_rows.append({
            'CandID': cid,
            'Nome': g['Nome'].iloc[0] if len(g)>0 else '',
            'Inscricao': g['Inscrição'].iloc[0] if ('Inscrição' in g.columns and len(g)>0) else None,
            'Preferencias_listadas': int(g['Ordem_Pref'].nunique()) if len(g)>0 else 0,
            'Motivo': 'sem alocação após DA'
        })
    pd.DataFrame(nao_rows).to_csv(NAO_CSV, index=False)

    # Auditoria de empates
    val['_ordem_planilha'] = val.reset_index().index
    records = []
    src = val.copy()
    src.rename(columns={'Órgão':'Orgao','Cód. Cargo Edital':'Codigo_Cargo'}, inplace=True)
    for (org, cod, tipo), grp in final_df.groupby(['Orgao','Codigo_Cargo','Tipo_Vaga_Usada']):
        m = grp.merge(src[['Orgao','Codigo_Cargo','CandID','Posição Real','Nota Final','_ordem_planilha']],
                      on=['Orgao','Codigo_Cargo','CandID'], how='left')
        for nota, g in m.groupby('Nota Final', dropna=False):
            if len(g) <= 1:
                continue
            g = g.reset_index(drop=True)
            g_sorted = g.sort_values(['Posição Real','_ordem_planilha'])
            final_seq  = list(g['CandID'])
            expect_seq = list(g_sorted['CandID'])
            if final_seq != expect_seq:
                records.append({
                    'Orgao': org, 'Codigo_Cargo': cod, 'Tipo_Vaga_Usada': tipo,
                    'Nota_Final_tie': float(nota) if pd.notna(nota) else None,
                    'Final_seq': final_seq,
                    'Esperado_por_PosicaoReal_seq': expect_seq
                })
    pd.DataFrame(records).to_csv(AUDIT_TIE_CSV, index=False)

    out = {
        'detalhes_csv': str(DETALHES_CSV),
        'xlsx':          str(XLSX_PATH),
        'nao_csv':       str(NAO_CSV),
        'audit_ties_csv':str(AUDIT_TIE_CSV),
        'prioridade_inicial_csv': str(CLASSIF_INICIAL_CSV),
        'total_final':   len(final_df)
    }
    print("\n--- Processamento (DA) concluído ---")
    print(out)

    # Download no Colab
    if IN_COLAB:
        for p in [DETALHES_CSV, XLSX_PATH, NAO_CSV, CLASSIF_INICIAL_CSV, AUDIT_TIE_CSV]:
            try:
                gfiles.download(str(p))
            except Exception as e:
                print(f"Falha ao baixar {p}: {e}")

if __name__ == '__main__':
    main()
