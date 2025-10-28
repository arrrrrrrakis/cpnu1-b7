import pandas as pd
import numpy as np
from collections import defaultdict
from pathlib import Path
import google.colab.files
import io

# --- 1. Upload do Arquivo ---
print("Por favor, faça o upload do arquivo 'b7p.xlsx'")
uploaded = google.colab.files.upload()

# Nome do arquivo de entrada esperado
FILE_NAME = 'b7p.xlsx' 
ABA = 'lista'

if FILE_NAME not in uploaded:
    print(f"\nErro: O arquivo '{FILE_NAME}' não foi encontrado.")
    print("Por favor, renomeie seu arquivo para 'b7p.xlsx' e execute a célula novamente.")
else:
    print(f"\nArquivo '{FILE_NAME}' carregado com sucesso. Iniciando processamento...")

    # --- Início do Script Original (com caminhos adaptados para o Colab) ---
    
    # v4_fix2 REVERT: remove editorial order from ranking; use it only for audit of tie-breaks
    # Ranking per cargo by (-Nota Final, Posição Real, CandID). Ordem_Pref only for cross-cargo choice.
    # Backfill allows temporary duplicates; iterative global conflict resolution by Ordem_Pref (tie by rank_tuple).

    # O arquivo de entrada agora é o arquivo que você carregou
    ARQUIVO = FILE_NAME

    # O diretório de saída agora é o /content/ do Colab
    OUT_DIR = Path('/content/')
    DETALHES_CSV = OUT_DIR/'detalhes_alocacao_v4fix2_revert_FULL.csv'
    XLSX_PATH = OUT_DIR/'resultado_alocacao_v4fix2_revert_FULL.xlsx'
    NAO_CSV = OUT_DIR/'nao_alocados_v4fix2_revert_FULL.csv'
    CLASSIF_INICIAL_CSV = OUT_DIR/'classificados_por_cargo_inicial_v4fix2_revert_FULL.csv'
    AUDIT_TIE_CSV = OUT_DIR/'auditoria_desempates_revert_FULL.csv'

    VAGAS_POR_CARGO = {
        ('AGU', 'B7-01-A'): {'AMPLA': 72, 'PPP': 15, 'PCD': 6, 'PI': 0, 'TOTAL': 93},
        ('AGU', 'B7-01-B'): {'AMPLA': 38, 'PPP': 7, 'PCD': 5, 'PI': 0, 'TOTAL': 50},
        ('AGU', 'B7-01-C'): {'AMPLA': 1, 'PPP': 0, 'PCD': 0, 'PI': 0, 'TOTAL': 1},
        ('AGU', 'B7-01-D'): {'AMPLA': 24, 'PPP': 6, 'PCD': 2, 'PI': 0, 'TOTAL': 32},
        ('AGU', 'B7-01-E'): {'AMPLA': 1, 'PPP': 0, 'PCD': 0, 'PI': 0, 'TOTAL': 1},
        ('FUNAI', 'B7-02-A'): {'AMPLA': 26, 'PPP': 14, 'PCD': 7, 'PI': 20, 'TOTAL': 67},
        ('FUNAI', 'B7-02-B'): {'AMPLA': 2, 'PPP': 0, 'PCD': 0, 'PI': 0, 'TOTAL': 2},
        ('FUNAI', 'B7-02-C'): {'AMPLA': 6, 'PPP': 3, 'PCD': 3, 'PI': 5, 'TOTAL': 17},
        ('FUNAI', 'B7-02-D'): {'AMPLA': 12, 'PPP': 6, 'PCD': 3, 'PI': 7, 'TOTAL': 28},
        ('FUNAI', 'B7-02-E'): {'AMPLA': 16, 'PPP': 7, 'PCD': 2, 'PI': 9, 'TOTAL': 34},
        ('IBGE', 'B7-03-A'): {'AMPLA': 13, 'PPP': 2, 'PCD': 1, 'PI': 0, 'TOTAL': 16},
        ('IBGE', 'B7-03-B'): {'AMPLA': 1, 'PPP': 0, 'PCD': 1, 'PI': 0, 'TOTAL': 2},
        ('IBGE', 'B7-03-C'): {'AMPLA': 3, 'PPP': 1, 'PCD': 0, 'PI': 0, 'TOTAL': 4},
        ('IBGE', 'B7-03-D'): {'AMPLA': 73, 'PPP': 19, 'PCD': 7, 'PI': 0, 'TOTAL': 99},
        ('IBGE', 'B7-03-E'): {'AMPLA': 3, 'PPP': 0, 'PCD': 0, 'PI': 0, 'TOTAL': 3},
        ('IBGE', 'B7-03-F'): {'AMPLA': 1, 'PPP': 0, 'PCD': 0, 'PI': 0, 'TOTAL': 1},
        ('IBGE', 'B7-03-G'): {'AMPLA': 14, 'PPP': 3, 'PCD': 2, 'PI': 0, 'TOTAL': 19},
        ('IBGE', 'B7-03-H'): {'AMPLA': 3, 'PPP': 0, 'PCD': 0, 'PI': 0, 'TOTAL': 3},
        ('IBGE', 'B7-03-I'): {'AMPLA': 25, 'PPP': 6, 'PCD': 3, 'PI': 0, 'TOTAL': 34},
        ('IBGE', 'B7-03-J'): {'AMPLA': 6, 'PPP': 1, 'PCD': 1, 'PI': 0, 'TOTAL': 8},
        ('INCRA', 'B7-04-A'): {'AMPLA': 7, 'PPP': 1, 'PCD': 1, 'PI': 0, 'TOTAL': 9},
        ('INCRA', 'B7-04-B'): {'AMPLA': 14, 'PPP': 3, 'PCD': 1, 'PI': 0, 'TOTAL': 18},
        ('INCRA', 'B7-04-C'): {'AMPLA': 43, 'PPP': 10, 'PCD': 3, 'PI': 0, 'TOTAL': 56},
        ('INCRA', 'B7-04-D'): {'AMPLA': 118, 'PPP': 27, 'PCD': 7, 'PI': 0, 'TOTAL': 152},
        ('INEP', 'B7-05-A'): {'AMPLA': 9, 'PPP': 2, 'PCD': 0, 'PI': 0, 'TOTAL': 11},
        ('MAPA', 'B7-06-A'): {'AMPLA': 10, 'PPP': 1, 'PCD': 0, 'PI': 0, 'TOTAL': 11},
        ('MCTI', 'B7-07-A'): {'AMPLA': 1, 'PPP': 1, 'PCD': 0, 'PI': 0, 'TOTAL': 2},
        ('MCTI', 'B7-07-B'): {'AMPLA': 3, 'PPP': 0, 'PCD': 0, 'PI': 0, 'TOTAL': 3},
        ('MCTI', 'B7-07-C'): {'AMPLA': 3, 'PPP': 2, 'PCD': 0, 'PI': 0, 'TOTAL': 5},
        ('MCTI', 'B7-07-D'): {'AMPLA': 1, 'PPP': 1, 'PCD': 1, 'PI': 0, 'TOTAL': 3},
        ('MCTI', 'B7-07-E'): {'AMPLA': 52, 'PPP': 7, 'PCD': 2, 'PI': 0, 'TOTAL': 61},
        ('MINC', 'B7-08-A'): {'AMPLA': 56, 'PPP': 16, 'PCD': 5, 'PI': 0, 'TOTAL': 77},
        ('MGI', 'B7-09-A'): {'AMPLA': 147, 'PPP': 38, 'PCD': 12, 'PI': 0, 'TOTAL': 197},
        ('MGI', 'B7-09-B'): {'AMPLA': 4, 'PPP': 2, 'PCD': 1, 'PI': 0, 'TOTAL': 7},
        ('MGI', 'B7-09-C'): {'AMPLA': 2, 'PPP': 0, 'PCD': 0, 'PI': 0, 'TOTAL': 2},
        ('MGI', 'B7-09-D'): {'AMPLA': 4, 'PPP': 1, 'PCD': 0, 'PI': 0, 'TOTAL': 5},
        ('MGI', 'B7-09-F'): {'AMPLA': 8, 'PPP': 2, 'PCD': 1, 'PI': 0, 'TOTAL': 11},
        ('MJSP', 'B7-10-A'): {'AMPLA': 110, 'PPP': 23, 'PCD': 10, 'PI': 0, 'TOTAL': 143},
        ('MS', 'B7-11-A'): {'AMPLA': 1, 'PPP': 2, 'PCD': 1, 'PI': 0, 'TOTAL': 4},
        ('MDIC', 'B7-12-A'): {'AMPLA': 51, 'PPP': 15, 'PCD': 5, 'PI': 0, 'TOTAL': 71},
        ('MPO', 'B7-13-A'): {'AMPLA': 37, 'PPP': 9, 'PCD': 3, 'PI': 0, 'TOTAL': 49},
        ('PREVIC', 'B7-14-A'): {'AMPLA': 3, 'PPP': 0, 'PCD': 1, 'PI': 0, 'TOTAL': 4},
    }

    # --- Load & normalize ---
    xl = pd.ExcelFile(ARQUIVO)
    df = xl.parse(ABA)
    df.columns = (df.columns.astype(str)
                  .str.replace('\n',' ', regex=False)
                  .str.replace('\r',' ', regex=False)
                  .str.replace('  ',' ', regex=False)
                  .str.strip())

    # Consolidate quota columns

    def first_non_empty(row):
        for x in row:
            if pd.isna(x):
                continue
            s = str(x).strip()
            if s and s.lower() != 'nan' and s != '-':
                return s
        return ''

    for keyword, newname in [('Indígena','Class_Indígena'), ('PCD','Class_PCD2'), ('Negra','Class_Negra3')]:
        cols = [c for c in df.columns if keyword.lower() in c.lower()]
        df[newname] = df[cols].apply(first_non_empty, axis=1) if cols else ''

    # Interesse
    df['Interesse_norm'] = df.get('Interesse?', 'sim').astype(str).str.strip().str.lower()

    # Key fields & numerics
    for col in ['Órgão','Cód. Cargo Edital','Cargo','Nome']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    df['Nota Final'] = pd.to_numeric(df['Nota Final'].astype(str).str.replace(',', '.', regex=False).str.replace(' ', '', regex=False), errors='coerce')
    df['Ordem_Pref'] = pd.to_numeric(df.get('Ordem_Pref'), errors='coerce')
    df['Posição Real'] = pd.to_numeric(df.get('Posição Real'), errors='coerce')

    df['CandID'] = df.get('Inscrição', df.get('Nome')).astype(str)
    df['Chave'] = list(zip(df['Órgão'].astype(str), df['Cód. Cargo Edital'].astype(str)))

    # Filter valid universe
    val = df[(df['Interesse_norm']=='sim') & df['Nota Final'].notna() & df['Ordem_Pref'].notna() & df['Chave'].isin(VAGAS_POR_CARGO.keys())].copy()

    # Candidate quota
    val['Cota'] = np.where(val['Class_Indígena'].astype(str).str.strip() != '', 'PI',
                    np.where(val['Class_PCD2'].astype(str).str.strip() != '', 'PCD',
                    np.where(val['Class_Negra3'].astype(str).str.strip() != '', 'PPP', 'AMPLA')))

    # Ranking tuple (reverted): NO Ordem_Pref here
    val['rank_tuple'] = list(zip(-val['Nota Final'], val['Posição Real'].fillna(1e9), val['CandID']))

    # DETAIL for backfill
    key = list(zip(val['CandID'], val['Chave']))
    cols_keep = ['Nome','Inscrição','Nota Final','Ordem_Pref','Cota','Cargo','rank_tuple']
    vals = val[cols_keep].to_dict(orient='records')
    DETAIL = {k:v for k,v in zip(key, vals)}

    # ---------- STEP 1: classification per cargo (using rank_tuple) ----------
    classificados = []
    backlogs = {}
    for chave, cfg in VAGAS_POR_CARGO.items():
        sub = val[val['Chave']==chave].sort_values('rank_tuple').copy()
        usados = {'PI':0,'PCD':0,'PPP':0,'AMPLA':0}
        sel_idx = set()

        # Reserves
        for cota in ['PI','PCD','PPP']:
            q = cfg[cota]
            if q <= 0: continue
            elig = sub[sub['Cota']==cota]
            for i, row in elig.iterrows():
                if usados[cota] >= q: break
                if i in sel_idx: continue
                classificados.append({
                    'Orgao': row['Órgão'], 'Codigo_Cargo': row['Cód. Cargo Edital'], 'Cargo': row['Cargo'],
                    'CandID': row['CandID'], 'Nome': row['Nome'], 'Inscricao': row.get('Inscrição'),
                    'Nota_Final': row['Nota Final'], 'Preferencia': row['Ordem_Pref'],
                    'Tipo_Cota_Candidato': row['Cota'], 'Tipo_Vaga_Usada': cota,
                    'rank_tuple': row['rank_tuple']
                })
                usados[cota]+=1; sel_idx.add(i)

        # Ampla (reversion)
        total = cfg['TOTAL']
        restantes = total - sum(usados.values())
        if restantes > 0:
            rem = sub[~sub.index.isin(sel_idx)]
            for i, row in rem.head(restantes).iterrows():
                classificados.append({
                    'Orgao': row['Órgão'], 'Codigo_Cargo': row['Cód. Cargo Edital'], 'Cargo': row['Cargo'],
                    'CandID': row['CandID'], 'Nome': row['Nome'], 'Inscricao': row.get('Inscrição'),
                    'Nota_Final': row['Nota Final'], 'Preferencia': row['Ordem_Pref'],
                    'Tipo_Cota_Candidato': row['Cota'], 'Tipo_Vaga_Usada': 'AMPLA',
                    'rank_tuple': row['rank_tuple']
                })
                usados['AMPLA']+=1; sel_idx.add(i)

        # Backlog residual
        backlog = []
        for i, row in sub.iterrows():
            if i in sel_idx: continue
            backlog.append({'CandID': row['CandID'], 'Chave': chave})
        backlogs[chave] = backlog

    classif_df = pd.DataFrame(classificados)
    classif_df.drop(columns=['rank_tuple']).to_csv(CLASSIF_INICIAL_CSV, index=False)
    print(f"Arquivo CSV de classificação inicial salvo em: {CLASSIF_INICIAL_CSV}")

    # ---------- STEP 2: backfill (allow duplicates) + iterative global conflict resolution ----------

    cargo_sel = defaultdict(list)
    for _, r in classif_df.sort_values(['Orgao','Codigo_Cargo','Tipo_Vaga_Usada','rank_tuple']).iterrows():
        cargo_sel[(r['Orgao'], r['Codigo_Cargo'])].append(dict(r))

    # Backfill: no candidate-exclusion; duplicates allowed; keep data completeness

    def backfill_all(cargo_sel):
        for chave, cfg in VAGAS_POR_CARGO.items():
            lst = cargo_sel[chave]
            counts = {'PI':0,'PCD':0,'PPP':0,'AMPLA':0}
            for r in lst:
                counts[r['Tipo_Vaga_Usada']] += 1
            total = cfg['TOTAL']
            backlog = backlogs.get(chave, [])

            def next_from(tipo):
                for i, row in enumerate(backlog):
                    cid = row['CandID']; ch = row['Chave']
                    if (cid, ch) not in DETAIL: continue
                    det = DETAIL[(cid, ch)]
                    if tipo != 'AMPLA' and det['Cota'] != tipo: continue
                    backlog.pop(i)
                    return cid, det
                return None, None

            for cota in ['PI','PCD','PPP']:
                while counts[cota] < cfg[cota] and len(lst) < total:
                    cid, det = next_from(cota)
                    if cid is None: break
                    lst.append({
                        'Orgao': chave[0], 'Codigo_Cargo': chave[1], 'Cargo': det['Cargo'],
                        'CandID': cid, 'Nome': det['Nome'], 'Inscricao': det['Inscrição'],
                        'Nota_Final': det['Nota Final'], 'Preferencia': det['Ordem_Pref'],
                        'Tipo_Cota_Candidato': det['Cota'], 'Tipo_Vaga_Usada': cota,
                        'rank_tuple': det['rank_tuple']
                    })
                    counts[cota]+=1

            while len(lst) < total:
                cid, det = next_from('AMPLA')
                if cid is None: break
                lst.append({
                    'Orgao': chave[0], 'Codigo_Cargo': chave[1], 'Cargo': det['Cargo'],
                    'CandID': cid, 'Nome': det['Nome'], 'Inscricao': det['Inscrição'],
                    'Nota_Final': det['Nota Final'], 'Preferencia': det['Ordem_Pref'],
                    'Tipo_Cota_Candidato': det['Cota'], 'Tipo_Vaga_Usada': 'AMPLA',
                    'rank_tuple': det['rank_tuple']
                })

    # Global conflict resolution: keep lowest Ordem_Pref; tie-break by rank_tuple

    def resolve_global(cargo_sel):
        ofertas = defaultdict(list)
        for chave, lst in cargo_sel.items():
            for r in lst:
                ofertas[r['CandID']].append({'chave': chave, 'pref': r['Preferencia'], 'rank_tuple': r['rank_tuple']})

        manter = {}
        remover = []
        for cid, lst in ofertas.items():
            lst_sorted = sorted(lst, key=lambda x: (x['pref'], x['rank_tuple']))
            manter[cid] = lst_sorted[0]['chave']
            for opt in lst_sorted[1:]:
                remover.append((cid, opt['chave']))

        mudou = False
        for cid, chave in remover:
            lst = cargo_sel[chave]
            for i, r in enumerate(list(lst)):
                if r['CandID'] == cid:
                    lst.pop(i); mudou = True; break
        return mudou

    # Stabilization loop
    print("Iniciando loop de estabilização (backfill e resolução de conflitos)...")
    for i in range(8):
        backfill_all(cargo_sel)
        changed = resolve_global(cargo_sel)
        print(f"  Iteração {i+1} concluída. Mudanças: {changed}")
        if not changed:
            print("Estabilização alcançada.")
            break

    # ---------- Build final outputs ----------
    final_rows = []
    for (org, cod), lst in cargo_sel.items():
        # sort by tipo-vaga then rank_tuple (Nota desc, Posição Real asc, CandID)
        lst_sorted = sorted(lst, key=lambda r: (
            {'PI':0,'PCD':1,'PPP':2,'AMPLA':3}.get(r['Tipo_Vaga_Usada'], 9),
            r['rank_tuple']
        ))
        for r in lst_sorted:
            final_rows.append({
                'Orgao': org, 'Codigo_Cargo': cod, 'Cargo': r.get('Cargo',''),
                'CandID': r['CandID'], 'Nome': r['Nome'], 'Inscricao': r['Inscricao'],
                'Nota_Final': r.get('Nota_Final'), 'Preferencia': r['Preferencia'],
                'Tipo_Cota_Candidato': r['Tipo_Cota_Candidato'], 'Tipo_Vaga_Usada': r['Tipo_Vaga_Usada'],
                'rank_tuple': r['rank_tuple'],
                'Destino': 'será nomeado'
            })

    final_df = pd.DataFrame(final_rows)
    final_df.to_csv(DETALHES_CSV, index=False)
    print(f"Arquivo CSV de detalhes finais salvo em: {DETALHES_CSV}")

    # XLSX with details + resumo + contagem
    with pd.ExcelWriter(XLSX_PATH, engine='openpyxl') as writer:
        final_df.to_excel(writer, sheet_name='Detalhes v4fix2_revert', index=False)
        # Resumo por cargo vs configuração
        resumo = []
        for (org, cod), grp in final_df.groupby(['Orgao','Codigo_Cargo']):
            cfg = VAGAS_POR_CARGO[(org, cod)]
            vc = grp['Tipo_Vaga_Usada'].value_counts().to_dict()
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
    print(f"Arquivo Excel de resultados salvo em: {XLSX_PATH}")

    # Non-allocated
    all_cands = set(val['CandID'])
    alloc_cands = set(final_df['CandID'])
    nao_cids = sorted(list(all_cands - alloc_cands))
    nao_rows = []
    for cid in nao_cids:
        g = val[val['CandID']==cid].sort_values('Ordem_Pref')
        nao_rows.append({
            'CandID': cid, 'Nome': g['Nome'].iloc[0],
            'Inscricao': g['Inscrição'].iloc[0] if 'Inscrição' in g.columns else None,
            'Preferencias_listadas': int(g['Ordem_Pref'].nunique()),
            'Motivo': 'não classificado no fechamento por cargo'
        })
    nao_df = pd.DataFrame(nao_rows)
    nao_df.to_csv(NAO_CSV, index=False)
    print(f"Arquivo CSV de não alocados salvo em: {NAO_CSV}")

    # ---------- AUDIT: desempates ----------
    # Objetivo: verificar que, em grupos com mesma Nota_Final dentro do cargo/tipo, a ordem está coerente com 'Posição Real'.
    # Também comparar com a ordem da planilha original (índice original) como fallback observacional.

    # reconstruir índice original
    val['_ordem_planilha'] = val.reset_index().index

    records = []
    for (org, cod, tipo), grp in final_df.groupby(['Orgao','Codigo_Cargo','Tipo_Vaga_Usada']):
        # attach source info to compute original positions
        merge_cols = ['Orgao','Codigo_Cargo','CandID']
        src = val.copy()
        src.rename(columns={'Órgão':'Orgao','Cód. Cargo Edital':'Codigo_Cargo'}, inplace=True)
        m = grp.merge(src[['Orgao','Codigo_Cargo','CandID','Posição Real','Nota Final','_ordem_planilha']], on=merge_cols, how='left', suffixes=('',''))
        # identify tie groups by Nota_Final
        for nota, g in m.groupby('Nota_Final', dropna=False):
            if len(g) <= 1:
                continue
            # our order index
            g = g.reset_index(drop=True)
            g['idx_final'] = g.index
            # expected order by Posição Real asc, then ordem_planilha asc
            g_sorted = g.sort_values(['Posição Real','_ordem_planilha'])
            # compare sequences of CandID
            final_seq = list(g['CandID'])
            expect_seq = list(g_sorted['CandID'])
            if final_seq != expect_seq:
                records.append({
                    'Orgao': org, 'Codigo_Cargo': cod, 'Tipo_Vaga_Usada': tipo,
                    'Nota_Final_tie': float(nota) if pd.notna(nota) else None,
                    'Final_seq': final_seq,
                    'Esperado_por_PosicaoReal_seq': expect_seq
                })

    pd.DataFrame(records).to_csv(AUDIT_TIE_CSV, index=False)
    print(f"Arquivo CSV de auditoria de desempates salvo em: {AUDIT_TIE_CSV}")

    # --- Fim do Script Original ---

    # --- 2. Saída Final e Download ---
    
    print("\n--- Processamento concluído ---")

    # Recria o dicionário de saída do script original
    output_files = {
        'detalhes_csv': str(DETALHES_CSV),
        'xlsx': str(XLSX_PATH),
        'nao_csv': str(NAO_CSV),
        'classificados_iniciais_csv': str(CLASSIF_INICIAL_CSV),
        'audit_ties_csv': str(AUDIT_TIE_CSV),
        'total_final': len(final_df)
    }

    print("Resultados:")
    print(output_files)

    print("\nIniciando o download de todos os arquivos gerados...")
    
    # Faz o download de todos os arquivos
    google.colab.files.download(str(DETALHES_CSV))
    google.colab.files.download(str(XLSX_PATH))
    google.colab.files.download(str(NAO_CSV))
    google.colab.files.download(str(CLASSIF_INICIAL_CSV))
    google.colab.files.download(str(AUDIT_TIE_CSV))
    
    print("Downloads concluídos.")

