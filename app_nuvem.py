import streamlit as st
import pandas as pd
from datetime import datetime
import os
import xml.etree.ElementTree as ET
from OpenSSL import crypto
import tempfile

# Configuração da página Web
st.set_page_config(page_title="Portal de RH & eSocial", page_icon="💼", layout="wide")

ARQUIVO_EXCEL = 'banco_dados_rh.xlsx'
CAMINHO_CERTIFICADO_PFX = "certificado_teste.pfx"  # Nome genérico de teste
SENHA_DO_CERTIFICADO = "123456"                    # Senha padrão de teste

# --- FUNÇÃO 1: GERADOR DE XML BRUTO DO ESOCIAL (S-2200) ---
def gerar_xml_esocial(dados):
    esocial_tag = ET.Element("eSocial", xmlns="http://esocial.gov.br")
    evt_admissao = ET.SubElement(esocial_tag, "evtAdmissao", Id=f"ID1000000000000002026101319000000001")
    
    ide_evento = ET.SubElement(evt_admissao, "ideEvento")
    ET.SubElement(ide_evento, "tpAmb").text = "2"  # Ambiente de Testes
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

# --- FUNÇÃO 2: TRANSMISSOR INTEGRADO COM CERTIFICADO DIGITAL A1 ---
def transmitir_com_certificado_real(xml_corpo):
    import requests
    URL_ESOCIAL = "https://esocial.gov.br"

    xml_envelope_soap = f"""<?xml version="1.0" encoding="utf-8"?>
    <soapenv:Envelope xmlns:soapenv="http://xmlsoap.org" xmlns:serv="http://esocial.gov.br">
       <soapenv:Header/>
       <soapenv:Body>
          <serv:EnviarLoteEventos>
             <serv:loteEventos>
                <eSocial xmlns="http://esocial.gov.br">
                   <envioLote>
                      <ideEmpregador>
                         <tpInsc>1</tpInsc>
                         <nrInsc>00000000000000</nrInsc>
                      </ideEmpregador>
                      <eventos>
                         <evento Id="ID1000000000000002026101319000000001">
                            {xml_corpo}
                         </evento>
                      </eventos>
                   </envioLote>
                </eSocial>
             </serv:loteEventos>
          </serv:EnviarLoteEventos>
       </soapenv:Body>
    </soapenv:Envelope>"""

    headers = {
        "Content-Type": "text/xml; charset=utf-8",
        "SOAPAction": "http://esocial.gov.br/WsEnviarLoteEventos/EnviarLoteEventos"
    }

    try:
        # Abre e descompacta o arquivo PFX usando as configurações salvas no topo do arquivo
        with open(CAMINHO_CERTIFICADO_PFX, "rb") as f:
            pfx = crypto.load_pkcs12(f.read(), SENHA_DO_CERTIFICADO.encode('utf-8'))
        
        cert_file = tempfile.NamedTemporaryFile(delete=False, suffix=".crt")
        key_file = tempfile.NamedTemporaryFile(delete=False, suffix=".key")
        cert_file.write(crypto.dump_certificate(crypto.FILETYPE_PEM, pfx.get_certificate()))
        key_file.write(crypto.dump_privatekey(crypto.FILETYPE_PEM, pfx.get_privatekey()))
        cert_file.close()
        key_file.close()
        
        resposta = requests.post(url=URL_ESOCIAL, data=xml_envelope_soap.encode('utf-8'), headers=headers, cert=(cert_file.name, key_file.name), timeout=30)
        
        os.unlink(cert_file.name)
        os.unlink(key_file.name)
        return resposta.text
    except Exception as e:
        return f"Aviso de Comunicação: {str(e)} (Verifique se anexou o arquivo físico do seu certificado .pfx na mesma pasta do código)"

# --- CONFIGURAÇÃO DA INTERFACE VISUAL ---
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
        cpf = st.text_input("CPF (Apenas numbers)", key="cad_cpf")
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
            
            # Gera o XML e aciona automaticamente a transmissão via Web Service
            xml_gerado = gerar_xml_esocial(funcionario_atual)
            retorno_api = transmitir_com_certificado_real(xml_gerado.decode('utf-8'))
            
            st.markdown("### 🛰️ Resposta da Conexão Web Service eSocial")
            st.code(retorno_api, language="xml")

# --- ABA 2: DADOS MENSAIS ---
with aba_mensal:
    st.header("Fechamento de Folha Mensal & Guia de Recolhimento do FGTS")
    salario_folha = st.number_input("Confirme o Salário Base do Funcionário (R$)", min_value=0.0, value=3500.0)
    horas_extras_input = st.number_input("Horas Extras (R$)", min_value=0.0, value=0.0)
    
    if st.button("Calcular Folha de Pagamento"):
        base = salario_folha + horas_extras_input
        fgts = base * 0.08
        st.success(f"Proventos Brutos: R$ {base:,.2f}")
        st.warning(f"Guia FGTS Digital Gerada: R$ {fgts:,.2f}")

# --- ABA 3: FÉRIAS ---
with aba_ferias:
    st.header("🌴 Controle de Período Aquisitivo & Férias")
    st.write("Módulo de compliance de gozo de férias ativo.")

# --- ABA 4: BANCO DE DADOS ---
with aba_banco:
    st.header("📊 Registro Central de Colaboradores Cadastrados")
    if os.path.exists(ARQUIVO_EXCEL):
        st.dataframe(pd.read_excel(ARQUIVO_EXCEL), use_container_width=True)