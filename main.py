import streamlit as st
import app_at
import comparador_projetos

def main():
    st.sidebar.title('Menu')
    escolha = st.sidebar.radio('Escolha a funcionalidade:', [
        'Redistribuidor de Pagodas',
        'Comparar SC com IDR'
    ])
    if escolha == 'Redistribuidor de Pagodas':
        app_at.main()
    elif escolha == 'Comparar SC com IDR':
        comparador_projetos.main()

if __name__ == '__main__':
    main() 