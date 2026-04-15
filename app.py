import streamlit as st
import pdfplumber
import pandas as pd
import re
from datetime import datetime
from io import BytesIO

# ====================== FUNÇÕES DE EXTRAÇÃO ======================
def limpar_valor(valor):
    if not valor:
        return ""
    valor = re.sub(r'[^\d.,-]', '', str(valor))
    valor = valor.replace('.', '').replace(',', '.')
    try:
        return float(valor)
    except:
        return valor

def extrair_dados_pdf(file_obj):
    with pdfplumber.open(file_obj) as pdf:
        texto_completo = ""
        tabelas = []
        for pagina in pdf.pages:
            texto_completo += (pagina.extract_text() or "") + "\n"
            tabelas.extend(pagina.extract_tables() or [])
    
    texto = texto_completo.upper()
    dados = {
        "CPF/CNPJ": "", "RAZÃO SOCIAL": "", "UF": "", "MUNICÍPIO E ENDEREÇO": "",
        "NÚMERO DE DOCUMENTO": "", "DATA DE EMISSÃO": "", "DATA DE ENTRADA": "",
        "SITUAÇÃO": "", "ACUMULADOR": "", "CFOP": "", "VALOR DE SERVIÇOS": "",
        "VALOR DESCONTO": "", "VALOR CONTÁBIL": "", "BASE DE CÁLCULO": "",
        "ALÍQUOTA ISS": "", "VALOR ISS NORMAL": "", "VALOR ISS RETIDO": "",
        "VALOR IRRF": "", "VALOR PIS": "", "VALOR COFINS": "", "VALOR CSLL": "",
        "VALOR CRF": "", "VALOR INSS": ""
    }

    padroes = {
        "CPF/CNPJ": r'(?:CPF/CNPJ|INSCRIÇÃO NO CPF/CNPJ)[:\s]*([\d./-]+)',
        "RAZÃO SOCIAL": r'(?:RAZÃO SOCIAL|NOME/RAZÃO SOCIAL)[:\s]*(.+?)(?=\s*(UF|MUNICÍPIO|ENDEREÇO|NÚMERO|\d{2}/\d{2}/\d{4}|$))',
        "UF": r'\bUF[:\s]*([A-Z]{2})\b',
        "MUNICÍPIO E ENDEREÇO": r'(?:MUNICÍPIO|ENDEREÇO)[:\s]*(.+?)(?=\s*(NÚMERO DE DOCUMENTO|DATA DE EMISSÃO|$))',
        "NÚMERO DE DOCUMENTO": r'(?:NÚMERO|NFS-E|NOTA FISCAL)[:\s]*(\d+)',
        "DATA DE EMISSÃO": r'(?:DATA DE EMISSÃO|EMISSÃO)[:\s]*(\d{2}/\d{2}/\d{4})',
        "DATA DE ENTRADA": r'(?:DATA DE ENTRADA|ENTRADA)[:\s]*(\d{2}/\d{2}/\d{4})',
        "SITUAÇÃO": r'SITUAÇÃO[:\s]*([A-ZÁÉÍÓÚÇ ]+)',
        "ACUMULADOR": r'ACUMULADOR[:\s]*([A-Z0-9]+)',
        "CFOP": r'CFOP[:\s]*(\d{4})',
        "VALOR DE SERVIÇOS": r'(?:VALOR DOS SERVIÇOS|VALOR SERVIÇO)[:\s]*R?\$?\s*([\d.,]+)',
        "VALOR DESCONTO": r'(?:DESCONTO|VALOR DESCONTO)[:\s]*R?\$?\s*([\d.,]+)',
        "VALOR CONTÁBIL": r'(?:VALOR CONTÁBIL|VALOR CONTABIL)[:\s]*R?\$?\s*([\d.,]+)',
        "BASE DE CÁLCULO": r'(?:BASE DE CÁLCULO|BASE CÁLCULO)[:\s]*R?\$?\s*([\d.,]+)',
        "ALÍQUOTA ISS": r'(?:ALÍQUOTA ISS|ALÍQ\. ISS)[:\s]*([\d.,]+)%?',
        "VALOR ISS NORMAL": r'(?:ISS NORMAL|VALOR ISS)[:\s]*R?\$?\s*([\d.,]+)',
        "VALOR ISS RETIDO": r'(?:ISS RETIDO|RETIDO)[:\s]*R?\$?\s*([\d.,]+)',
        "VALOR IRRF": r'IRRF[:\s]*R?\$?\s*([\d.,]+)',
        "VALOR PIS": r'PIS[:\s]*R?\$?\s*([\d.,]+)',
        "VALOR COFINS": r'COFINS[:\s]*R?\$?\s*([\d.,]+)',
        "VALOR CSLL": r'CSLL[:\s]*R?\$?\s*([\d.,]+)',
        "VALOR CRF": r'CRF[:\s]*R?\$?\s*([\d.,]+)',
        "VALOR INSS": r'INSS[:\s]*R?\$?\s*([\d.,]+)',
    }

    for campo, regex in padroes.items():
        match = re.search(regex, texto_completo, re.IGNORECASE | re.DOTALL)
        if match:
            dados[campo] = match.group(1).strip()

    # Extração de itens
    itens = []
    for tabela in tabelas:
        for linha in tabela:
            if linha and len(linha) >= 3:
                cod = str(linha[0]).strip() if linha[0] else ""
                qtd = str(linha[1]).strip() if len(linha) > 1 else ""
                v_unit = str(linha[2]).strip() if len(linha) > 2 else ""
                if cod and re.search(r'\d', cod):
                    itens.append({
                        "CÓDIGO DO ITEM": cod,
                        "QUANTIDADE": limpar_valor(qtd),
                        "VALOR UNITÁRIO": limpar_valor(v_unit)
                    })
    if not itens:
        linhas_item = re.findall(r'(\d{1,10})\s+(\d+[,.]?\d*)\s+R?\$?\s*([\d.,]+)', texto_completo)
        for cod, qtd, vunit in linhas_item:
            itens.append({
                "CÓDIGO DO ITEM": cod,
                "QUANTIDADE": limpar_valor(qtd),
                "VALOR UNITÁRIO": limpar_valor(vunit)
            })
    return dados, itens

# ====================== INTERFACE STREAMLIT ======================
st.set_page_config(page_title="NF-PDF → Excel Automático", layout="wide", page_icon="🚀")

st.title("🚀 NF-PDF → Excel Automático")
st.markdown("**Faça upload dos PDFs das suas Notas Fiscais e extraia automaticamente todos os campos para uma planilha Excel profissional.**")

uploaded_files = st.file_uploader(
    "📤 Selecione um ou vários arquivos PDF de Notas Fiscais",
    type="pdf",
    accept_multiple_files=True,
    help="Suporta NFS-e, NF de serviço e qualquer layout de Nota Fiscal em PDF"
)

if uploaded_files:
    st.info(f"✅ {len(uploaded_files)} arquivo(s) PDF carregado(s)")

    if st.button("🔄 Processar Notas Fiscais", type="primary", use_container_width=True):
        with st.spinner("Processando PDFs... Por favor, aguarde."):
            progress_bar = st.progress(0)
            todas_linhas = []
            
            for idx, uploaded_file in enumerate(uploaded_files):
                try:
                    file_bytes = BytesIO(uploaded_file.getvalue())
                    cabecalho, lista_itens = extrair_dados_pdf(file_bytes)
                    
                    if not lista_itens:
                        lista_itens = [{}]
                    
                    for item in lista_itens:
                        linha = {**cabecalho, **item}
                        linha["Nome do arquivo PDF"] = uploaded_file.name
                        linha["Data/Hora do processamento"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                        todas_linhas.append(linha)
                
                except Exception as e:
                    st.error(f"❌ Erro ao processar {uploaded_file.name}: {str(e)}")
                    continue
                
                progress_bar.progress((idx + 1) / len(uploaded_files))
            
            df = pd.DataFrame(todas_linhas)
            
            ordem_colunas = [
                "CPF/CNPJ", "RAZÃO SOCIAL", "UF", "MUNICÍPIO E ENDEREÇO",
                "NÚMERO DE DOCUMENTO", "DATA DE EMISSÃO", "DATA DE ENTRADA", "SITUAÇÃO",
                "ACUMULADOR", "CFOP", "VALOR DE SERVIÇOS", "VALOR DESCONTO",
                "VALOR CONTÁBIL", "BASE DE CÁLCULO", "ALÍQUOTA ISS",
                "VALOR ISS NORMAL", "VALOR ISS RETIDO", "VALOR IRRF",
                "VALOR PIS", "VALOR COFINS", "VALOR CSLL", "VALOR CRF", "VALOR INSS",
                "CÓDIGO DO ITEM", "QUANTIDADE", "VALOR UNITÁRIO",
                "Nome do arquivo PDF", "Data/Hora do processamento"
            ]
            
            for col in ordem_colunas:
                if col not in df.columns:
                    df[col] = ""
            df = df[ordem_colunas]
            
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name="Notas_Fiscais")
                workbook = writer.book
                worksheet = writer.sheets["Notas_Fiscais"]
                for col in worksheet.columns:
                    max_length = 0
                    column_letter = col[0].column_letter
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
            
            output.seek(0)
            
            st.success(f"✅ **Processamento concluído!**  \n📄 {len(uploaded_files)} PDFs processados  \n📊 {len(df)} linhas geradas")
            st.subheader("🔍 Prévia dos dados (10 primeiras linhas)")
            st.dataframe(df.head(10), use_container_width=True)
            
            st.download_button(
                label="⬇️ Baixar Excel Completo",
                data=output,
                file_name=f"Notas_Fiscais_Extraidas_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )
            st.balloons()

st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: #64748b; font-size: 0.9rem;'>"
    "NF-PDF → Excel Automático • Desenvolvido para automação fiscal"
    "</p>",
    unsafe_allow_html=True
)
