import streamlit as st
import pandas as pd
import openpyxl as openpyxl

# Configurações iniciais do Streamlit
st.set_page_config(page_title="Gestão de Colaboradores", layout="wide")
st.title("Gestão de Colaboradores")

# DataFrame inicial para armazenar os dados
if "colaboradores" not in st.session_state:
    st.session_state.colaboradores = pd.DataFrame(
        columns=["SAA", "SAT", "SAO", "COLABORADOR", "SETOR", "CONTRATO", "OBSERVAÇÃO"]
    )

# Função para resetar índice após operações de exclusão
def reset_index():
    st.session_state.colaboradores.reset_index(drop=True, inplace=True)

# Função para carregar o arquivo (Excel ou CSV)
def carregar_arquivo(arquivo):
    try:
        # Verifica a extensão do arquivo para usar o método correto
        if arquivo.type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" or arquivo.type == "application/vnd.ms-excel":
            # Lê o arquivo Excel
            df = pd.read_excel(arquivo, engine="openpyxl")
        elif arquivo.type == "text/csv":
            # Lê o arquivo CSV
            df = pd.read_csv(arquivo)
        else:
            st.error("Formato de arquivo não suportado. Carregue um arquivo Excel (.xlsx, .xls) ou CSV (.csv).")
            return
        
        # Verifica se as colunas essenciais existem
        colunas_esperadas = ["SAA", "SAT", "SAO", "COLABORADOR", "SETOR", "CONTRATO", "OBSERVAÇÃO"]
        if all(coluna in df.columns for coluna in colunas_esperadas):
            st.session_state.colaboradores = df
            st.success("Arquivo carregado com sucesso!")
        else:
            st.error("O arquivo não contém as colunas esperadas.")
    except Exception as e:
        st.error(f"Erro ao carregar o arquivo: {e}")

# Painel de Estatísticas
with st.sidebar:
    st.image("https://automni.com.br/wp-content/uploads/2022/08/Logo-IDLogistics-PNG.svg", use_container_width=True)
    menu = st.radio("Navegação", ["Dashboard", "Consulta e Inclusão", "Tabela de Colaboradores"])

# Dashboard
if menu == "Dashboard":
    # Upload de arquivo (Excel ou CSV)
    arquivo = st.file_uploader("Carregar arquivo", type=["xlsx", "xls", "csv"])
    if arquivo:
        carregar_arquivo(arquivo)

    # Obtenção do total de colaboradores
    total_colaboradores = len(st.session_state.colaboradores)
    
    # Garantir que a comparação não seja sensível a maiúsculas/minúsculas nem a espaços
    colaboradores_ativos = len(st.session_state.colaboradores[st.session_state.colaboradores["OBSERVAÇÃO"].str.strip().str.lower() == "ativo"])
    colaboradores_inativos = len(st.session_state.colaboradores[st.session_state.colaboradores["OBSERVAÇÃO"].str.strip().str.lower() == "inativo"])

    col1, col2, col3 = st.columns(3)
    
    # Total de colaboradores
    with col1:
        st.markdown("""<div style='text-align: center; font-size: 24px; font-weight: bold;'>Total de Colaboradores</div>""", unsafe_allow_html=True)
        st.markdown(f"""<div style='text-align: center; font-size: 36px;'> {total_colaboradores}</div>""", unsafe_allow_html=True)
    
    # Colaboradores Ativos
    with col2:
        st.markdown("""<div style='text-align: center; font-size: 24px; font-weight: bold; color: green;'>Colaboradores Ativos</div>""", unsafe_allow_html=True)
        st.markdown(f"""<div style='text-align: center; font-size: 36px; color: green;'>{colaboradores_ativos}</div>""", unsafe_allow_html=True)
    
    # Colaboradores Inativos
    with col3:
        st.markdown("""<div style='text-align: center; font-size: 24px; font-weight: bold; color: red;'>Colaboradores Inativos</div>""", unsafe_allow_html=True)
        st.markdown(f"""<div style='text-align: center; font-size: 36px; color: red;'>{colaboradores_inativos}</div>""", unsafe_allow_html=True)
    
    # Exibição da lista de colaboradores
    st.subheader("Lista de Colaboradores")
    st.dataframe(st.session_state.colaboradores)

# Consulta e Inclusão
elif menu == "Consulta e Inclusão":
    st.header("Consulta e Inclusão de Colaboradores")

    # Formulário para adicionar ou editar colaborador
    with st.form("incluir_colaborador"):
        col1, col2 = st.columns(2)
        with col1:
            saa = st.text_input("SAA")
            sat = st.text_input("SAT")
            sao = st.text_input("SAO")
            colaborador = st.text_input("Colaborador")
        with col2:
            setor = st.text_input("Setor")
            contrato = st.text_input("Contrato")
            observacao = st.text_area("Observação")

        submit = st.form_submit_button("Incluir/Atualizar")

        if submit:
            novo_dado = {
                "SAA": saa,
                "SAT": sat,
                "SAO": sao,
                "COLABORADOR": colaborador,
                "SETOR": setor,
                "CONTRATO": contrato,
                "OBSERVAÇÃO": observacao,
            }

            st.session_state.colaboradores = pd.concat([ 
                st.session_state.colaboradores, pd.DataFrame([novo_dado])
            ], ignore_index=True)
            st.success("Colaborador incluído/atualizado com sucesso!")

# Tabela de Colaboradores
elif menu == "Tabela de Colaboradores":
    st.header("Tabela de Colaboradores")
    for i, row in st.session_state.colaboradores.iterrows():
        with st.expander(f"{row['COLABORADOR']} ({row['SETOR']})"):
            col1, col2, col3 = st.columns(3)
            with col1:
                st.text_input("SAA", value=row["SAA"], key=f"saa_{i}")
                st.text_input("SAT", value=row["SAT"], key=f"sat_{i}")
                st.text_input("SAO", value=row["SAO"], key=f"sao_{i}")
                st.text_input("Colaborador", value=row["COLABORADOR"], key=f"colaborador_{i}")
            with col2:
                st.text_input("Setor", value=row["SETOR"], key=f"setor_{i}")
                st.text_input("Contrato", value=row["CONTRATO"], key=f"contrato_{i}")
                st.text_area("Observação", value=row["OBSERVAÇÃO"], key=f"observacao_{i}")
            with col3:
                if st.button("Atualizar", key=f"atualizar_{i}"):
                    st.session_state.colaboradores.loc[i] = {
                        "SAA": st.session_state[f"saa_{i}"],
                        "SAT": st.session_state[f"sat_{i}"],
                        "SAO": st.session_state[f"sao_{i}"],
                        "COLABORADOR": st.session_state[f"colaborador_{i}"],
                        "SETOR": st.session_state[f"setor_{i}"],
                        "CONTRATO": st.session_state[f"contrato_{i}"],
                        "OBSERVAÇÃO": st.session_state[f"observacao_{i}"],
                    }
                    st.success("Colaborador atualizado com sucesso!")

                if st.button("Excluir", key=f"excluir_{i}"):
                    st.session_state.colaboradores.drop(i, inplace=True)
                    reset_index()
                    st.success("Colaborador excluído com sucesso!")

    # Botão para exportar a lista para Excel
    st.subheader("Exportar Lista de Colaboradores")
    def exportar_para_excel():
        return st.session_state.colaboradores.to_csv(index=False).encode('utf-8')

    st.download_button(
        label="Exportar Lista para Excel",
        data=exportar_para_excel(),
        file_name="lista_colaboradores.csv",
        mime="text/csv"
    )
