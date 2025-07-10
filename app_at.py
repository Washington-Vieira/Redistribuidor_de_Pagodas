import pandas as pd
import re
from collections import defaultdict
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook
import streamlit as st
import io

def processar_pagodas(df_original, limite_por_grupo=36, limites_excecao=None):
    if limites_excecao is None:
        limites_excecao = {}
    # Validação
    if 'Semiacabado' not in df_original.columns or 'Pagoda' not in df_original.columns:
        raise ValueError("A planilha deve conter as colunas 'Semiacabado' e 'Pagoda'.")
    regex_pg = re.compile(r'(PGS?|pgs?)(\d{2})/(\d{2})', re.IGNORECASE)
    df = df_original.copy()
    pagoda_original = df['Pagoda'].copy()
    nova_pagoda = df['Pagoda'].copy()
    # --- ETAPA 1: Corrigir grupos quebrados (PGS05 → PGS03...)
    grupos_usados = defaultdict(set)
    for pag in nova_pagoda:
        match = regex_pg.match(str(pag))
        if match:
            prefixo = match.group(1).upper()
            grupo = int(match.group(2))
            grupos_usados[prefixo].add(grupo)
    grupos_ordenados = {prefixo: sorted(list(grupos)) for prefixo, grupos in grupos_usados.items()}
    mapa_grupo_corrigido = {}
    for prefixo, grupos in grupos_ordenados.items():
        for novo_num, grupo_original in enumerate(grupos, start=1):
            mapa_grupo_corrigido[(prefixo, grupo_original)] = novo_num
    for idx, valor in nova_pagoda.items():
        match = regex_pg.match(str(valor))
        if match:
            prefixo = match.group(1).upper()
            grupo = int(match.group(2))
            numero = int(match.group(3))
            novo_grupo = mapa_grupo_corrigido.get((prefixo, grupo))
            nova_pagoda.iloc[idx] = f"{prefixo}{novo_grupo:02}/{numero:02}"
    # --- ETAPA 2: Corrigir buracos nas sequências dentro dos grupos
    grupos = defaultdict(list)
    for idx, pagoda in nova_pagoda.items():
        match = regex_pg.match(str(pagoda))
        if match:
            prefixo = match.group(1).upper()
            grupo = int(match.group(2))
            numero = int(match.group(3))
            chave = f"{prefixo}{grupo:02}"
            grupos[chave].append((numero, idx))
    for chave, lista in grupos.items():
        lista_ordenada = sorted(lista, key=lambda x: x[0])
        # Limite para este grupo (exceção ou padrão)
        limite = limites_excecao.get(chave, limite_por_grupo)
        for i, (_, idx) in enumerate(lista_ordenada):
            novo_num = i + 1
            if novo_num <= limite:
                nova_pagoda.iloc[idx] = f"{chave}/{novo_num:02}"
    # --- ETAPA 3: Redistribuir excedentes para novos grupos
    contador = defaultdict(int)
    for pag in nova_pagoda:
        match = regex_pg.match(str(pag))
        if match:
            prefixo = match.group(1).upper()
            grupo = int(match.group(2))
            numero = int(match.group(3))
            chave = f"{prefixo}{grupo:02}"
            contador[chave] += 1
    def proxima_pagoda(prefixo, contador, limites_excecao, limite_padrao):
        grupo = 1
        while True:
            chave = f"{prefixo}{grupo:02}"
            limite = limites_excecao.get(chave, limite_padrao)
            if contador[chave] < limite:
                contador[chave] += 1
                return f"{chave}/{contador[chave]:02}"
            grupo += 1
    for idx, pagoda in nova_pagoda.items():
        match = regex_pg.match(str(pagoda))
        if match:
            prefixo = match.group(1).upper()
            numero = int(match.group(3))
            grupo = int(match.group(2))
            chave = f"{prefixo}{grupo:02}"
            limite = limites_excecao.get(chave, limite_por_grupo)
            if numero > limite:
                nova_pagoda.iloc[idx] = proxima_pagoda(prefixo, contador, limites_excecao, limite_por_grupo)
    # --- ETAPA 4: Ordenação final por prefixo, grupo e número
    def sort_pagoda(pagoda_str):
        match = regex_pg.match(str(pagoda_str))
        if match:
            prefixo = match.group(1).upper()
            grupo = int(match.group(2))
            numero = int(match.group(3))
            ordem_prefixo = 0 if prefixo == 'PG' else 1
            return (ordem_prefixo, grupo, numero)
        return (99, 999, 999)
    df['Pagoda'] = nova_pagoda
    df = df.sort_values(by='Pagoda', key=lambda col: col.map(sort_pagoda)).reset_index(drop=True)
    # --- ETAPA 5: Gerar log de alterações
    log_alteracoes = []
    for i in range(len(df_original)):
        original = pagoda_original.iloc[i]
        novo = nova_pagoda.iloc[i]
        if original != novo:
            log_alteracoes.append({
                'Linha': i + 2,
                'Semiacabado': df_original.loc[i, 'Semiacabado'],
                'Pagoda_como_estava': original,
                'Pagoda_como_ficou': novo
            })
    return df, log_alteracoes

# --- STREAMLIT APP ---
def main():
    st.title('Redistribuidor de Pagodas (com exceções por grupo)')
    st.write('Faça upload do arquivo Excel do YMS com as colunas Semiacabado e Pagoda. O sistema irá redistribuir para no máximo 36 por grupo, exceto grupos definidos como exceção.')
    uploaded_file = st.file_uploader('Selecione o arquivo Excel', type=['xlsx'])
    limite_padrao = st.number_input('Limite padrão por grupo', min_value=1, max_value=100, value=36)
    st.write('Exceções (ex: PG07=30, PGS03=28)')
    excecoes_str = st.text_input('Exceções (formato: PG07=30,PG08=25)', value='')
    limites_excecao = {}
    if excecoes_str.strip():
        for item in excecoes_str.split(','):
            if '=' in item:
                chave, val = item.split('=')
                limites_excecao[chave.strip().upper()] = int(val.strip())
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        try:
            df_corrigido, log = processar_pagodas(df, limite_por_grupo=limite_padrao, limites_excecao=limites_excecao)
        except Exception as e:
            st.error(str(e))
            return
        st.success('Redistribuição concluída!')
        st.dataframe(df_corrigido.head(20))
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_corrigido.to_excel(writer, index=False, sheet_name='Pagodas_Corrigidas')
            if log:
                df_log = pd.DataFrame(log)
                df_log.to_excel(writer, index=False, sheet_name='Alteracoes')
        st.download_button('Baixar resultado Excel', data=output.getvalue(), file_name='pagodas_corrigidas.xlsx')

if __name__ == '__main__':
    main()

