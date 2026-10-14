import streamlit as st
import pandas as pd
from datetime import datetime
import os
import xml.etree.ElementTree as ET

# Configuração da página Web
st.set_page_config(page_title="Portal de RH & eSocial", page_icon="💼", layout="wide")

ARQUIVO_EXCEL = 'banco_dados_rh.xlsx'

# --- FUNÇÃO INTELEGENTE: GERADOR DE XML BRUTO DO ESOCIAL (S-2200) ---
def gerar_xml_esocial(dados):
    esocial_tag = ET.Element("eSocial", xmlns="http://esocial.gov.br")
    evt_admissao = ET.SubElement(esocial_tag, "evtAdmissao", Id=f"ID1000000000000002026101319000000001")
    
    ide_evento = ET.SubElement(evt_admissao, "ideEvento")
    ET.SubElement(ide_evento, "tpAmb").text = "2"  # 2 = Ambiente de Testes
    ET.SubElement(ide_evento, "procEmi").text = "1"
    ET.SubElement(ide_evento, "verProc").text = "1.0.0"
    
    ide_empregador = ET.SubElement(evt_admissao, "ideEmpregador")
    ET.SubElement(ide_empregador, "tpInsc").text = "1"
    ET.SubElement(ide_empregador, "nrInsc").text = "00000000000000"
    
    trabalhador = ET.SubElement(evt_admissao, "trabalhador")
    ET.SubElement(trabalhador, "cpfTrab").text = str(dados.get('CPF', ''))
    ET.SubElement(trabalhador, "nmTrab").text = str(dados.get('Nome_Completo', ''))
    ET.SubElement(trabalhador, "dtNascto").text = str(dados.get('Data_Nascimento', ''))
    ET.SubElement(trabalhador, "estCiv").text = "1"
    
    vinculo = ET.SubElement(evt_admissao, "vinculo")
    ET.SubElement(vinculo, "matricula").text = str(dados.get('Matricula_eSocial', ''))
    ET.SubElement(vinculo, "tpRegTrab").text = "1"
    ET.SubElement(vinculo, "tpRegPrev").text = "1"
    
    info_clt = ET.SubElement(vinculo, "infoRegime")
    info_clt_detalhe = ET.SubElement(info_clt, "infoCLT")
    ET.SubElement(info_clt_detalhe, "dtAdm").text = str(dados.get('Data_Admissao', ''))
    ET.SubElement(info_clt_detalhe, "tpRegJor").text = "1"
    
    info_contrato = ET.SubElement(vinculo, "infoContrato")
    ET.SubElement(info_contrato, "nmCargo").text = str(dados.get('Cargo', ''))
    
    remuneracao = ET.SubElement(info_contrato, "remuneracao")
    ET.SubElement(remuneracao, "vrSalFx").text = str(dados.get('Salario_Base', '')).replace("R$", "").strip()
    ET.SubElement(remuneracao, "undSalFixo").text = "5"
    
    xml_puro = ET.tostring(esocial_tag, encoding="utf-8", method="xml")
    return b'<?xml version="1.0" encoding="utf-8"?>\n' + xml_puro

# --- CABEÇALHO VISUAL ---
st.title("💼 Portal Integrado de Recursos Humanos & eSocial")
st.subheader("Gerenciador Admissional, Cálculo de Folha e Compliance na Nuvem")

# Abas Superiores 100% Ativas
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
        civil = st.selectbox("Estado Civil", ["Solteiro(a)", "Casado(a)", "Divorciado(a)", "Viúvo(a)"], key="cad_civil")
        cargo = st.text_input("Cargo ou Função Ativa", key="cad_cargo")
        salario = st.text_input("Salário Base R$ (Ex: 3500)", key="cad_salario")
        data_adm = st.text_input("Data de Admissão (DD/MM/AAAA)", key="cad_adm")
        aso_data = st.text_input("Data do Exame ASO Admissional", key="cad_aso")

    st.markdown("### Pacote de Benefícios Ativos")
    c1, c2, c3, c4 = st.columns(4)
    vt = c1.checkbox("Vale Transporte (VT)")
    vr = c2.checkbox("Vale Refeição (VR)")
    va = c3.checkbox("Vale Alimentação (VA)")
    am = c4.checkbox("Assistência Médica (AM)")

    if st.button("💾 Salvar Contratação e Enviar Banco de Dados", type="primary"):
        if not nome or not cpf or not cargo or not salario or not data_adm or not aso_data:
            st.error("BLOQUEADO! Preencha os dados essenciais (Nome, CPF, Cargo, Salário, Admissão e ASO) antes de enviar.")
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
            
            xml_gerado = gerar_xml_esocial(funcionario_atual)
            st.markdown("### 📑 Integração eSocial Pronta")
            st.download_button(
                label="📥 Baixar XML de Admissão Oficial (S-2200)",
                data=xml_gerado,
                file_name=f"S2200_Admissao_{cpf}.xml",
                mime="text/xml"
            )

# --- ABA 2: DADOS MENSAIS & FGTS ---
with aba_mensal:
    st.header("Fechamento de Folha Mensal & Guia de Recolhimento do FGTS")
    
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

# --- ABA 3: FÉRIAS COMPLETA ---
with aba_ferias:
    st.header("🌴 Controle de Período Aquisitivo & Cálculo de Férias CLTs")
    st.markdown("Configure e simule as férias obrigatórias seguindo as regras da CLT e as notificações do evento **S-2230** do eSocial.")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        salario_ferias = st.number_input("Confirme o Salário Base para Férias (R$)", min_value=0.0, value=3500.0, key="fer_sal")

dias_gozo = st.slider("Quantidade de Dias de Férias", min_value=10, max_value=30, value=30)
with col_f2:
        vender_abono = st.radio("Deseja vender 1/3 das férias (Abono Pecuniário)?", ["Não", "Sim"])
        adianta_13 = st.radio("Adiantar 1ª Parcela do 13º Salário?", ["Não", "Sim"])

if st.button("🌴 Calcular Recibo de Férias (eSocial S-2230)"):
        # Cálculo proporcional dos dias de férias
        valor_dias_ferias = (salario_ferias / 30) * dias_gozo
        terco_constitucional = valor_dias_ferias / 3

        abono_pecuniario = (salario_ferias / 30) * 10 if vender_abono == "Sim" else 0.0
        terco_abono = abono_pecuniario / 3 if vender_abono == "Sim" else 0.0
        adiantamento_13_valor = salario_ferias / 2 if adianta_13 == "Sim" else 0.0

        total_bruto_ferias = valor_dias_ferias + terco_constitucional + abono_pecuniario + terco_abono + adiantamento_13_valor

        # Desconto INSS simplificado de férias
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
st.warning("⚠️ Compliance eSocial: O aviso de férias deve ser enviado ao governo com no mínimo 30 dias de antecedência através do evento S-2230.")

# --- ABA 4: BANCO DE DADOS CENTRAL ---
with aba_banco:
    st.header("📊 Registro Central de Colaboradores Cadastrados")
    st.markdown("Esta tabela interativa lê diretamente o arquivo Excel do sistema e lista todos os profissionais.")
    
    if st.button("🔄 Atualizar Tabela do Banco de Dados"):
        st.rerun()
        
    if os.path.exists(ARQUIVO_EXCEL):
        df_visualizar = pd.read_excel(ARQUIVO_EXCEL)
        st.dataframe(df_visualizar, use_container_width=True)
        
        st.markdown("### 📈 Estatísticas Rápidas")
        col_e1, col_e2 = st.columns(2)
        col_e1.metric(label="Total de Funcionários Cadastrados", value=len(df_visualizar))
        col_e2.metric(label="Última Matrícula eSocial Gerada", value=str(df_visualizar['Matricula_eSocial'].iloc[-1]) if len(df_visualizar) > 0 else "N/A")
    else:
        st.info("Nenhum funcionário cadastrado no banco de dados até o momento. Cadastre o primeiro na Aba 1!")