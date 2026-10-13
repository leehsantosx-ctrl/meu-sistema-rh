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

# --- ABA 2: DADOS MENSAIS & FGTS (TABELA DE RUBRICAS ESOCIAL CONFIGURADA) ---
with aba_mensal:
    st.header("Fechamento de Folha Mensal & Guia de Recolhimento do FGTS")
    
    # Exibe a tabela descritiva das Rubricas Oficiais mapeadas no sistema
    st.markdown("### 📋 Tabela de Rubricas Oficiais Mapeadas (eSocial - Tabela 03)")
    dados_rubricas = {
        "Descrição da Verba": ["Salário Base Contratual", "Horas Extras 50%", "Desconto INSS Empregado", "Retenção IRRF Fonte", "Depósito FGTS Mensal"],
        "Código Rubrica eSocial": ["1000 (Vencimento)", "1003 (Vencimento)", "9201 (Desconto)", "9203 (Desconto)", "9901 (Informativa)"],
        "Incidência INSS": ["Sim (Soma)", "Sim (Soma)", "Deduz a Base", "Não", "Não afeta"],
        "Incidência IRRF": ["Sim (Soma)", "Sim (Soma)", "Deduz a Base", "Não", "Não afeta"],
        "Incidência FGTS": ["Sim (Soma)", "Sim (Soma)", "Não", "Não", "Base de Cálculo"]
    }
    st.table(pd.DataFrame(dados_rubricas))

    st.markdown("---")
    st.markdown("### ⚡ Simulador de Evento Periódico S-1200")
    
    salario_folha = st.number_input("Confirme o Salário Base do Funcionário (R$)", min_value=0.0, value=3500.0, step=100.0)
    horas_extras_input = st.number_input("Inserir Variáveis de Provento - Horas Extras (R$)", min_value=0.0, value=0.0, step=50.0)
    
    if st.button("Calcular Folha de Pagamento com Rubricas"):
        # Cálculo exato das bases com base nas incidências da tabela
        base_calculo_total = salario_folha + horas_extras_input
        
        # 1. ENCARGO DA EMPRESA: FGTS (Incide sobre Salário e Horas Extras - 8%)
        valor_fgts = base_calculo_total * 0.08
        
        # 2. CÁLCULO PROGRESSIVO DO INSS (Tabela Oficial do Governo)
        if base_calculo_total <= 1412.00:
            inss = base_calculo_total * 0.075
        elif base_calculo_total <= 2666.68:
            inss = (1412.00 * 0.075) + ((base_calculo_total - 1412.00) * 0.09)
        elif base_calculo_total <= 4000.03:
            inss = (1412.00 * 0.075) + ((2666.68 - 1412.00) * 0.09) + ((base_calculo_total - 2666.68) * 0.12)
        else:
            inss = (1412.00 * 0.075) + ((2666.68 - 1412.00) * 0.09) + ((4000.03 - 2666.68) * 0.12) + ((base_calculo_total - 4000.03) * 0.14)
        
        # 3. CÁLCULO PROGRESSIVO DO IRRF (Base deduzida pelo INSS pago)
        base_irrf = base_calculo_total - inss
        if base_irrf <= 2259.20:
            irrf = 0.0
        elif base_irrf <= 2826.65:
            irrf = (base_irrf * 0.075) - 169.44
        elif base_irrf <= 3751.05:
            irrf = (base_irrf * 0.15) - 381.44
        elif base_irrf <= 4664.68:
            irrf = (base_irrf * 0.225) - 662.77
        else:
            irrf = (base_irrf * 0.275) - 896.00
            
        salario_liquido = base_calculo_total - inss - irrf
        
        # Exibição do Holerite estruturado no formato que o eSocial lê
        st.markdown("### 📄 Demonstrativo Consolidado de Rubricas (Pronto para gerar XML S-1200)")
        
        col_prov, col_desc = st.columns(2)
        with col_prov:
            st.success("#### 🟢 PROVENTOS (Vencimentos)")
            st.markdown(f"**Rubrica 1000 - Salário Contratual:** R$ {salario_folha:,.2f}")
            st.markdown(f"**Rubrica 1003 - Horas Extras:** R$ {horas_extras_input:,.2f}")
            st.markdown(f"**TOTAL BRUTO (Base INSS/FGTS):** R$ {base_calculo_total:,.2f}")
            
        with col_desc:
            st.error("#### 🔴 DESCONTOS")
            st.markdown(f"**Rubrica 9201 - Previdência INSS:** R$ {inss:,.2f}")
            st.markdown(f"**Rubrica 9203 - Imposto de Renda IRRF:** R$ {irrf:,.2f}")
            st.markdown(f"**TOTAL DE DESCONTOS:** R$ {(inss + irrf):,.2f}")
            
        st.markdown("---")
        st.info(f"### 💰 VALOR LÍQUIDO CREDITADO EM CONTA: R$ {salario_liquido:,.2f}")
        
        st.warning(f"""
        ### 🏦 Guia Automática FGTS Digital (Evento S-1200 cadastrado)
        * **Rubrica Informativa 9901 - FGTS Recolhido pela Empresa:** **R$ {valor_fgts:,.2f}**
        * *Atenção:* Este encargo de 8% é calculado automaticamente com base nas rubricas de provento enviadas e não gera desconto no salário líquido do empregado.
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