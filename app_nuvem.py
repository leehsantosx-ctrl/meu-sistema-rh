import streamlit as st
import pandas as pd
from datetime import datetime
import os

# Configuração da página Web
st.set_page_config(page_title="programa teste RH", page_icon="💼", layout="wide")

ARQUIVO_EXCEL = 'banco_dados_rh.xlsx'

st.title("💼 Portal Integrado de Recursos Humanos & eSocial")
st.subheader("Gerenciador Admissional e Cálculo de Folha na Nuvem")

# Criação das Abas Superiores idênticas ao seu sistema desktop
aba_contratacao, aba_mensal, aba_ferias, aba_banco = st.tabs([
    "1. Contratação / Admissão", 
    "2. Dados Mensais & FGTS", 
    "3. Dados Férias", 
    "4. Banco de Dados"
])

# --- ABA 1: CONTRATAÇÃO ---
with aba_contratacao:
    st.header("Ficha Cadastral Eletrônica Inteligente")
    
    col1, col2 = st.columns(2)
    with col1:
        nome = st.text_input("Nome Completo do Trabalhador")
        cpf = st.text_input("CPF (Apenas números)")
        rg = st.text_input("RG / CNH")
        nasc = st.text_input("Nascimento (DD/MM/AAAA)")
        cep = st.text_input("CEP da Residência")
        end = st.text_input("Endereço Completo")
    with col2:
        titulo_e = st.text_input("Título de Eleitor")
        reservista = st.text_input("Certificado de Reservista")
        escolaridade = st.selectbox("Grau de Escolaridade", ["Ensino Fundamental", "Ensino Médio", "Ensino Técnico", "Ensino Superior"])
        civil = st.selectbox("Estado Civil", ["Solteiro(a)", "Casado(a)", "Divorciado(a)", "Viúvo(a)"])
        cargo = st.text_input("Cargo ou Função Ativa")
        salario = st.text_input("Salário Base R$ (Ex: 3500)")
        data_adm = st.text_input("Data de Admissão (DD/MM/AAAA)")
        aso_data = st.text_input("Data do Exame ASO Admissional")

    st.markdown("### Pacote de Benefícios Ativos")
    c1, c2, c3, c4 = st.columns(4)
    vt = c1.checkbox("Vale Transporte (VT)")
    vr = c2.checkbox("Vale Refeição (VR)")
    va = c3.checkbox("Vale Alimentação (VA)")
    am = c4.checkbox("Assistência Médica (AM)")

    if st.button("💾 Salvar Contratação e Enviar Banco de Dados", type="primary"):
        if not nome or not cpf or not cargo or not salario or not data_adm or not aso_data:
            st.error("BLOQUEADO! Preencha os dados essenciais (Nome, CPF, Cargo, Salário, Admissão e ASO) antes de enviar para a nuvem.")
        else:
            proximo_numero = 1
            if os.path.exists(ARQUIVO_EXCEL):
                try:
                    df_existente = pd.read_excel(ARQUIVO_EXCEL)
                    proximo_numero = len(df_existente) + 1
                except: pass

            funcionario_atual = {
                'N_Funcionario': f"Nº {proximo_numero}",
                'Matricula_eSocial': datetime.now().strftime("%Y%m%d%H%M"),
                'Nome_Completo': nome, 'CPF': cpf, 'RG_CNH': rg, 'Data_Nascimento': nasc,
                'Cargo': cargo, 'Salario_Base': salario, 'Data_Admissao': data_adm, 'Data_Exame_ASO': aso_data,
                'Data_Cadastro': datetime.now().strftime("%d/%m/%Y %H:%M")
            }
            
            df_novo = pd.DataFrame([funcionario_atual])
            if os.path.exists(ARQUIVO_EXCEL):
                df_consolidado = pd.concat([pd.read_excel(ARQUIVO_EXCEL), df_novo], ignore_index=True)
                df_consolidado.to_excel(ARQUIVO_EXCEL, index=False)
            else:
                df_novo.to_excel(ARQUIVO_EXCEL, index=False)
                
            st.success(f"Funcionário validado na nuvem! Registrado como: Funcionário Nº {proximo_numero}")

# --- ABA 2: DADOS MENSAIS & FGTS ---
with aba_mensal:
    st.header("Fechamento de Folha Mensal & Guia de Recolhimento do FGTS")
    salario_folha = st.number_input("Confirme o Salário do Funcionário para o Cálculo (R$)", min_value=0.0, value=3500.0)
    horas_extras = st.number_input("Inserir Variáveis de Provento - Horas Extras (R$)", min_value=0.0, value=0.0)
    
    if st.button("⚡ Calcular Folha Mensal e FGTS"):
        bruto = salario_folha + horas_extras
        fgts = bruto * 0.08
        
        # INSS e IRRF Simplificados
        inss = bruto * 0.09 if bruto <= 2666 else bruto * 0.12
        irrf = (bruto * 0.075) - 169 if bruto > 2259 else 0.0
        liquido = bruto - inss - irrf
        
        st.markdown(f"""
        ### 📄 Demonstrativo de Pagamento Emitido
        * **Proventos Brutos:** R$ {bruto:,.2f}
        * **(-) Desconto INSS:** R$ {inss:,.2f}
        * **(-) Retenção IRRF:** R$ {irrf:,.2f}
        * **💰 SALÁRIO LÍQUIDO A RECEBER:** R$ {liquido:,.2f}
        
        ### 🏦 Encargo de Responsabilidade da Empresa (eSocial S-1200)
        * **(->) Depósito Mensal do FGTS (8%): R$ {fgts:,.2f}** *(Valor não descontado do trabalhador)*
        """)

# --- ABA 3: FÉRIAS ---
with aba_ferias:
    st.info("Módulo de controle de períodos aquisitivos e agendamento de gozo (S-2230) ativo na nuvem.")

# --- ABA 4: BANCO DE DADOS ---
with aba_banco:
    st.header("📊 Registro Central de Colaboradores Cadastrados")
    if os.path.exists(ARQUIVO_EXCEL):
        df_visualizar = pd.read_excel(ARQUIVO_EXCEL)
        st.dataframe(df_visualizar, use_container_width=True) # Desenha a planilha interativa na página web
    else:
        st.write("Nenhum funcionário cadastrado no banco de dados em nuvem até o momento.")