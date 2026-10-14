import xml.etree.ElementTree as ET
from OpenSSL import crypto
import tempfile
import os

# Configurações do Certificado (Mude quando usar um real)
CAMINHO_CERTIFICADO_PFX = "seu_certificado.pfx"
SENHA_DO_CERTIFICADO = "sua_senha_aqui"

# --- GERADOR DE XML BRUTO DO ESOCIAL (S-2200 - ADMISSÃO) ---
def gerar_xml_esocial(dados):
    esocial_tag = ET.Element("eSocial", xmlns="http://esocial.gov.br")
    evt_admissao = ET.SubElement(esocial_tag, "evtAdmissao", Id="ID1000000000000002026101319000000001")
    
    ide_evento = ET.SubElement(evt_admissao, "ideEvento")
    ET.SubElement(ide_evento, "tpAmb").text = "2" # Ambiente de Testes
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
    
    return b'<?xml version="1.0" encoding="utf-8"?>\n' + ET.tostring(esocial_tag, encoding="utf-8")

# --- GERADOR DE XML DE FOLHA MENSAL (S-1200 - REMUNERAÇÃO) ---
def gerar_xml_s1200_folha(matricula, cpf, salario, inss, irrf):
    esocial_tag = ET.Element("eSocial", xmlns="http://esocial.gov.br")
    evt_remun = ET.SubElement(esocial_tag, "evtRemun", Id="ID1000000000000002026101319000000002")
    
    ide_evento = ET.SubElement(evt_remun, "ideEvento")
    ET.SubElement(ide_evento, "indGuia").text = "1"
    ET.SubElement(ide_evento, "perApur").text = "2026-10"
    ET.SubElement(ide_evento, "tpAmb").text = "2"
    
    ide_empregador = ET.SubElement(evt_remun, "ideEmpregador")
    ET.SubElement(ide_empregador, "tpInsc").text = "1"
    ET.SubElement(ide_empregador, "nrInsc").text = "00000000000000"
    
    ide_trabalhador = ET.SubElement(evt_remun, "ideTrabalhador")
    ET.SubElement(ide_trabalhador, "cpfTrab").text = str(cpf)
    
    info_per_apur = ET.SubElement(ide_trabalhador, "infoPerApur")
    ide_estabelecimento = ET.SubElement(info_per_apur, "ideEstabLot")
    ET.SubElement(ide_estabelecimento, "tpInsc").text = "1"
    ET.SubElement(ide_estabelecimento, "nrInsc").text = "00000000000000"
    
    remun_trabalhador = ET.SubElement(ide_estabelecimento, "remunTrabr")
    ET.SubElement(remun_trabalhador, "matricula").text = str(matricula)
    
    info_venc = ET.SubElement(remun_trabalhador, "infoLograd")
    
    rubrica_sal = ET.SubElement(info_venc, "detRubr")
    ET.SubElement(rubrica_sal, "codRubr").text = "1000"
    ET.SubElement(rubrica_sal, "vrRubr").text = f"{salario:.2f}"
    
    rubrica_inss = ET.SubElement(info_venc, "detRubr")
    ET.SubElement(rubrica_inss, "codRubr").text = "9201"
    ET.SubElement(rubrica_inss, "vrRubr").text = f"{inss:.2f}"
    
    return b'<?xml version="1.0" encoding="utf-8"?>\n' + ET.tostring(esocial_tag, encoding="utf-8")

# --- TRANSMISSOR INTEGRADO VIA WEB SERVICE API ---
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
        return f"Aviso de Comunicação: {str(e)} (Modo Sandbox de homologação ativo para testes em nuvem)"
