import streamlit as st
import pandas as pd
from datetime import datetime
import os
from motores_esocial import gerar_xml_esocial, gerar_xml_s1200_folha, transmitir_com_certificado_real

# Configuração da página Web
st.set_page_config(page_title="Portal de RH & eSocial", page_icon="💼", layout="wide")

ARQUIVO_EXCEL = 'banco_dados_rh.xlsx'

st.title("💼 Portal Integrado de Recursos Humanos & eSocial")
st.subheader("Gerenciador Admissional, Cálculo de Folha e Compliance na Nuvem")

aba_contratacao, aba_mensal, aba_ferias, aba_banco = st.tabs([
    "1. Contratação / Admissão", 
    "2. Dados Mensais & FGTS", 
    "3. Dados Férias & Compliance", 
    "4. Banco de Dados Central"
])

# --- ABA 1: CONTRATAÇÃO ---
with aba_contratacao:
    st.header("Ficha Cadastral Eletrônica Inteligente")
    col1, col2 = st.columns(2)
    with col1:
        nome = st.text_input("Nome Completo do Trabalhador", key="cad_nome")
        cpf = st.text_input("CPF (Apenas números)", key="cad_cpf")
        rg = st.text_input("RG / CNH", key="cad_rg")
        nasc = st.text_input("Nascimento (DD/MM/AAAA)", key="cad_nasc")
        cep = st.text_input("CEP da Residência", key="cad_cep")
        end = st.text_input("Endereço Completo", key="cad_end")
    with col2:
        titulo_e = st.text_input("Título de Eleitor", key="cad_titulo")
        reservista = st.text_input("Certificado de Reservista", key="cad_reservista")
        escolaridade = st.selectbox("Grau de Escolaridade", ["Ensino Fundamental", "Ensino Médio", "Ensino Técnico", "Ensino Superior"], key="cad_esco")
        civil = st.selectbox("Estado Civil", ["Solteiro(a)", "Casado(a)", "Divorciado(a)"], key="cad_civil")
        cargo = st.text_input("Cargo ou Função Ativa", key="cad_cargo")
        salario = st.text_input("Salário Base R$ (Ex: 3500)", key="cad_salario")
        data_adm = st.text_input("Data de Admissão (DD/MM/AAAA)", key="cad_adm")
        aso_data = st.text_input("Data do Exame ASO Admissional", key="cad_aso")

    if st.button("💾 Salvar Contratação e Enviar Banco de Dados", type="primary"):
        if not nome or not cpf or not cargo or not salario or not data_adm or not aso_data:
            st.error("BLOQUEADO! Preencha os dados essenciais antes de enviar.")
        else:
            proximo_numero = 1
            if os.path.exists(ARQUIVO_EXCEL):
                try: proximo_numero = len(pd.read_excel(ARQUIVO_EXCEL)) + 1
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
                pd.concat([pd.read_excel(ARQUIVO_EXCEL), df_novo], ignore_index=True).to_excel(ARQUIVO_EXCEL, index=False)
            else:
                df_novo.to_excel(ARQUIVO_EXCEL, index=False)
                
            st.success(f"Funcionário validado na nuvem! Registrado como: Funcionário Nº {proximo_numero}")
            
            xml_gerado = gerar_xml_esocial(funcionario_atual)
            retorno_api = transmitir_com_certificado_real(xml_gerado.decode('utf-8'))
            
            st.markdown("### 🛰️ Resposta da Conexão Web Service eSocial")
            st.code(retorno_api, language="xml")

# --- ABA 2: DADOS MENSAIS & FGTS ---
with aba_mensal:
    st.header("Fechamento de Folha Mensal & Guia de Recolhimento do FGTS")
    
    if os.path.exists(ARQUIVO_EXCEL):
        df_func = pd.read_excel(ARQUIVO_EXCEL)
        lista_funcionarios = df_func['Nome_Completo'].tolist()
        
        st.markdown("### 👤 Selecione o Funcionário para Fechamento")
        nome_selecionado = st.selectbox("Escolha o colaborador:", lista_funcionarios, key="sel_mensal")
        
        dados_f = df_func[df_func['Nome_Completo'] == nome_selecionado].iloc[0]
        matricula_atual = dados_f['Matricula_eSocial']
        cpf_atual = dados_f['CPF']
        salario_base_planilha = float(str(dados_f['Salario_Base']).replace("R$", "").strip())
        
        st.info(f"Matrícula eSocial: {matricula_atual} | CPF: {cpf_atual}")
        
        salario_folha = st.number_input("Confirme o Salário Base (R$)", min_value=0.0, value=salario_base_planilha, key="num_sal_folha")
        horas_extras_input = st.number_input("Horas Extras Trabalhadas (R$)", min_value=0.0, value=0.0, key="num_he_folha")
        
        if st.button("Calcular Folha de Pagamento com Rubricas"):
            base_calculo_total = salario_folha + horas_extras_input
            valor_fgts = base_calculo_total * 0.08
            
            if base_calculo_total <= 1412.00: inss = base_calculo_total * 0.075
            elif base_calculo_total <= 2666.68: inss = (1412.00 * 0.075) + ((base_calculo_total - 1412.00) * 0.09)
            elif base_calculo_total <= 4000.03: inss = (1412.00 * 0.075) + ((2666.68 - 1412.00) * 0.09) + ((base_calculo_total - 2666.68) * 0.12)
            else: inss = (1412.00 * 0.075) + ((2666.68 - 1412.00) * 0.09) + ((4000.03 - 2666.68) * 0.12) + ((base_calculo_total - 4000.03) * 0.14)
            
            base_irrf = base_calculo_total - inss
            if base_irrf <= 2259.20: irrf = 0.0
            elif base_irrf <= 2826.65: irrf = (base_irrf * 0.075) - 169.44
            elif base_irrf <= 3751.05: irrf = (base_irrf * 0.15) - 381.44
            elif base_irrf <= 4664.68: irrf = (base_irrf * 0.225) - 662.77
            else: irrf = (base_irrf * 0.275) - 896.00
                
            salario_liquido = base_calculo_total - inss - irrf
            
            st.markdown("### 📄 Demonstrativo Consolidado de Rubricas")
            col_prov, col_desc = st.columns(2)
            with col_prov:
                st.success("#### 🟢 PROVENTOS")
                st.markdown(f"**Rubrica 1000 - Salário:** R$ {salario_folha:,.2f}")
                st.markdown(f"**Rubrica 1003 - Horas Extras:** R$ {horas_extras_input:,.2f}")
                st.markdown(f"**TOTAL BRUTO:** R$ {base_calculo_total:,.2f}")
            with col_desc:
                st.error("#### 🔴 DESCONTOS")
                st.markdown(f"**Rubrica 9201 - Previdência INSS:** R$ {inss:,.2f}")
                st.markdown(f"**Rubrica 9203 - Imposto IRRF:** R$ {irrf:,.2f}")
                st.markdown(f"**TOTAL DESCONTOS:** R$ {(inss + irrf):,.2f}")
                
            st.markdown("---")
            st.info(f"### 💰 VALOR LÍQUIDO CREDITADO EM CONTA: R$ {salario_liquido:,.2f}")
            st.warning(f"### 🏦 FGTS Digital (Evento S-1200): R$ {valor_fgts:,.2f}")
            
            st.markdown("---")
            xml_folha = gerar_xml_s1200_folha(matricula_atual, cpf_atual, base_calculo_total, inss, irrf)
            retorno_folha_api = transmitir_com_certificado_real(xml_folha.decode('utf-8'))
            
            st.success(f"🚀 Evento S-1200 de Remuneração gerado para {nome_selecionado}!")
            st.code(retorno_folha_api, language="xml")
    else:
        st.info("Nenhum funcionário encontrado no banco de dados. Cadastre primeiro na Aba 1.")

# --- TÓPICO 2: ABA 3 DE FÉRIAS TOTALMENTE ATIVA ---
with aba_ferias:
    st.header("🌴 Controle de Período Aquisitivo & Férias")
    st.markdown("Configure e simule as férias obrigatórias seguindo as regras da CLT e as notificações do evento **S-2230** do eSocial.")
    
    if os.path.exists(ARQUIVO_EXCEL):
        df_func_f = pd.read_excel(ARQUIVO_EXCEL)
        lista_func_f = df_func_f['Nome_Completo'].tolist()
        
        st.markdown("### 👤 Selecione o Funcionário para Recibo de Férias")
        nome_f_selecionado = st.selectbox("Escolha o colaborador:", lista_func_f, key="sel_ferias")
        
        dados_f_ferias = df_func_f[df_func_f['Nome_Completo'] == nome_f_selecionado].iloc[0]
        salario_base_ferias = float(str(dados_f_ferias['Salario_Base']).replace("R$", "").strip())
        
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            salario_ferias = st.number_input("Confirme o Salário Base para Férias (R$)", min_value=0.0, value=salario_base_ferias, key="num_sal_fer")
            dias_gozo = st.slider("Quantidade de Dias de Férias", min_value=10, max_value=30, value=30, key="slide_dias_fer")
        with col_f2:
            vender_abono = st.radio("Deseja vender 1/3 das férias (Abono Pecuniário)?", ["Não", "Sim"], key="rad_abono")
            adianta_13 = st.radio("Adiantar 1ª Parcela do 13º Salário?", ["Não", "Sim"], key="rad_13")

        # --- CONTINUAÇÃO DA ABA 3: FÉRIAS & COMPLIANCE ---
        if st.button("🌴 Calcular Recibo de Férias (eSocial S-2230)"):
            valor_dias_ferias = (salario_ferias / 30) * dias_gozo
            terco_constitucional = valor_dias_ferias / 3
            
            abono_pecuniario = (salario_ferias / 30) * 10 if vender_abono == "Sim" else 0.0
            terco_abono = abono_pecuniario / 3 if vender_abono == "Sim" else 0.0
            adiantamento_13_valor = salario_ferias / 2 if adianta_13 == "Sim" else 0.0
            
            total_bruto_ferias = valor_dias_ferias + terco_constitucional + abono_pecuniario + terco_abono + adiantamento_13_valor
            
            inss_ferias = total_bruto_ferias * 0.12 if total_bruto_ferias > 2666 else total_bruto_ferias * 0.09
            total_liquido_ferias = total_bruto_ferias - inss_ferias
            
            st.markdown("### 📄 Demonstrativo do Recibo de Férias")
            st.success(f"""
            * **(+) Valor dos {dias_gozo} dias de Gozo:** R$ {valor_dias_ferias:,.2f}
            * **(+) 1/3 Constitucional de Férias:** R$ {terco_constitucional:,.2f}
            * **(+) Abono Pecuniário (10 dias vendidos):** R$ {abono_pecuniario:,.2f}
            * **(+) 1/3 sobre Abono Pecuniário:** R$ {terco_abono:,.2f}
            * **(+) Adiantamento da 1ª Parcela do 13º:** R$ {adiantamento_13_valor:,.2f}
            * ------------------------------------------------------------------
            * **(=) TOTAL BRUTO DE FÉRIAS:** R$ {total_bruto_ferias:,.2f}
            * **(-) Desconto INSS sobre Férias:** R$ {inss_ferias:,.2f}
            * ------------------------------------------------------------------
            * **💰 VALOR LÍQUIDO A PAGAR (2 dias antes do início): R$ {total_liquido_ferias:,.2f}**
            """)
            st.warning("⚠️ **Compliance eSocial:** O aviso de férias deve ser enviado ao governo com no mínimo **30 dias de antecedência** através do evento S-2230.")
    else:
        st.info("Nenhum funcionário cadastrado no banco de dados. Cadastre primeiro na Aba 1.")

# --- ABA 4: BANCO DE DADOS CENTRAL ---
with aba_banco:
    st.header("📊 Registro Central de Colaboradores Cadastrados")
    st.markdown("Esta tabela interativa lê diretamente o arquivo Excel do sistema e lista todos os profissionais.")
    
    if os.path.exists(ARQUIVO_EXCEL):
        df_visualizar = pd.read_excel(ARQUIVO_EXCEL)
        st.dataframe(df_visualizar, use_container_width=True)
        
        # Estrutura de Download corrigida dentro do bloco correto
        with open(ARQUIVO_EXCEL, "rb") as file:
            st.download_button(
                label="📥 Baixar Planilha Excel Completa (.xlsx)",
                data=file,
                file_name="banco_dados_rh_atualizado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        st.markdown("### 📈 Estatísticas Rápidas")
        col_e1, col_e2 = st.columns(2)
        col_e1.metric(label="Total de Funcionários Cadastrados", value=len(df_visualizar))
        col_e2.metric(label="Última Matrícula eSocial Gerada", value=str(df_visualizar['Matricula_eSocial'].iloc[-1]) if len(df_visualizar) > 0 else "N/A")
    else:
        st.info("Nenhum funcionário cadastrado no banco de dados até o momento. Cadastre o primeiro na Aba 1!")
