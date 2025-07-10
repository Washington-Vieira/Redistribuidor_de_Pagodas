import streamlit as st
import pandas as pd
import io

def comparar_planilhas_diferencas(df_pagodas, df_template):
    # Garante que as colunas existam
    if 'Semiacabado' not in df_template.columns or 'Projeto' not in df_template.columns:
        raise ValueError("A planilha de template deve conter as colunas 'Semiacabado' e 'Projeto'.")
    if 'Semiacabado' not in df_pagodas.columns:
        raise ValueError("A planilha de pagodas deve conter a coluna 'Semiacabado'.")
    semiacabados_pagodas = set(df_pagodas['Semiacabado'].astype(str))
    semiacabados_template = set(df_template['Semiacabado'].astype(str))
    # Presentes só no template
    so_no_template = semiacabados_template - semiacabados_pagodas
    # Presentes só no pagodas
    so_no_pagodas = semiacabados_pagodas - semiacabados_template
    # Monta DataFrame de resultado
    resultado = []
    for semi in so_no_template:
        projeto = df_template.loc[df_template['Semiacabado'].astype(str) == semi, 'Projeto'].values[0]
        resultado.append({'Semiacabado': semi, 'Projeto': projeto, 'Ausente em': 'YMS'})
    for semi in so_no_pagodas:
        resultado.append({'Semiacabado': semi, 'Projeto': '', 'Ausente em': 'Obsoleto'})
    return pd.DataFrame(resultado, columns=['Semiacabado', 'Projeto', 'Ausente em'])

def main():
    st.title('Comparador de Semiacabados entre Pagodas Corrigidas e Template')
    st.write('Faça upload do arquivo de pagodas corrigidas e do template (com colunas Semiacabado e Projeto). O sistema irá mostrar os Semiacabados que estão presentes em apenas um dos arquivos e indicar em qual está ausente.')
    pagodas_file = st.file_uploader('Upload do arquivo de pagodas corrigidas (.xlsx)', type=['xlsx'], key='pagodas')
    template_file = st.file_uploader('Upload do template (Semiacabado e Projeto) (.xlsx)', type=['xlsx'], key='template')
    if pagodas_file and template_file:
        # Lê a primeira aba, independentemente do nome
        df_pagodas = pd.read_excel(pagodas_file, sheet_name=0)
        df_template = pd.read_excel(template_file)
        try:
            df_diferencas = comparar_planilhas_diferencas(df_pagodas, df_template)
        except Exception as e:
            st.error(str(e))
            return
        st.success(f'{len(df_diferencas)} Semiacabados presentes em apenas um dos arquivos!')
        st.dataframe(df_diferencas)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_diferencas.to_excel(writer, index=False, sheet_name='Diferencas')
        st.download_button('Baixar resultado Excel', data=output.getvalue(), file_name='diferencas_semiacabados.xlsx')

if __name__ == '__main__':
    main() 