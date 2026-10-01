import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import PyPDF2
import re
import io

st.set_page_config(page_title="Mapa de Cotação Automático", page_icon="🏗️")

st.title("🏗️ Gerador de Mapa de Cotação")
st.write("Anexe os orçamentos e confirme os dados antes de gerar a planilha final.")

# Função para extrair texto e valores do PDF
def extrair_dados_pdf(pdf_file):
    try:
        leitor = PyPDF2.PdfReader(pdf_file)
        texto = ""
        for pagina in leitor.pages:
            texto += pagina.extract_text() + "\n"
        
        linhas = [linha for linha in texto.split('\n') if linha.strip()]
        nome_sugerido = linhas[0][:30] if linhas else pdf_file.name
        
        padrao_moeda = r'(\d{1,3}(?:\.\d{3})*,\d{2})'
        valores = re.findall(padrao_moeda, texto)
        
        maior_valor = 0.0
        if valores:
            valores_num = [float(v.replace('.', '').replace(',', '.')) for v in valores]
            maior_valor = max(valores_num)
            
        return nome_sugerido, maior_valor
    except Exception:
        return pdf_file.name, 0.0

# Uploads e Input do Material
modelo_excel = st.file_uploader("1. Anexe o Modelo Excel (MPC - Modelo.xlsx)", type=["xlsx"])
fornecedores_pdfs = st.file_uploader("2. Anexe os PDFs dos Fornecedores (Máx 3)", type=["pdf"], accept_multiple_files=True)

# NOVO CAMPO: Onde você define o nome do material para preencher a planilha
nome_material = st.text_input("3. Qual é o material/serviço sendo cotado?", placeholder="Ex: TUBO ESGOTO PVC 300MM")

if modelo_excel and fornecedores_pdfs:
    st.subheader("Verifique os Dados Extraídos")
    
    dados_fornecedores = []
    
    for i, pdf in enumerate(fornecedores_pdfs[:3]):
        nome_ext, valor_ext = extrair_dados_pdf(pdf)
        
        st.markdown(f"**Fornecedor {i+1} (Arquivo: {pdf.name})**")
        col1, col2 = st.columns(2)
        
        with col1:
            nome_final = st.text_input(f"Nome do Fornecedor {i+1}", value=nome_ext, key=f"nome_{i}")
        with col2:
            valor_final = st.number_input(f"Valor Unitário (R$) {i+1}", value=float(valor_ext), format="%.2f", key=f"val_{i}")
            
        dados_fornecedores.append({"nome": nome_final.upper(), "valor": valor_final})
        st.divider()

    if st.button("Gerar Planilha Excel"):
        if not nome_material:
            st.error("Por favor, preencha o nome do material/serviço no passo 3.")
        else:
            try:
                wb = openpyxl.load_workbook(modelo_excel)
                ws = wb.active
                
                # 1. Substituir "Topografia" pelo nome do material (Linha 2, Coluna 5)
                ws.cell(row=2, column=5, value=nome_material.upper())
                
                # 2. Preencher a coluna "Material" na tabela de itens (Linha 8, Coluna 2)
                ws.cell(row=8, column=2, value=nome_material.upper())
                
                # Encontrar o menor preço
                valores = [f['valor'] for f in dados_fornecedores if f['valor'] > 0]
                menor_preco = min(valores) if valores else 0
                
                colunas_fornecedores = [7, 9, 11] # Colunas G, I, K
                
                # Preencher os dados na folha
                for i, fornecedor in enumerate(dados_fornecedores):
                    col = colunas_fornecedores[i]
                    
                    # Nome no cabeçalho
                    ws.cell(row=2, column=col, value=fornecedor['nome'])
                    
                    # Preço Unitário na linha 8
                    celula_preco = ws.cell(row=8, column=col)
                    celula_preco.value = fornecedor['valor']
                    celula_preco.number_format = 'R$ #,##0.00'
                    
                    # Destacar Vencedor a Verde
                    if fornecedor['valor'] == menor_preco and fornecedor['valor'] > 0:
                        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                        font_winner = Font(color="006100", bold=True)
                        celula_preco.fill = green_fill
                        celula_preco.font = font_winner
                        
                        # Pintar o total do vencedor
                        ws.cell(row=8, column=col+1).fill = green_fill
                        ws.cell(row=8, column=col+1).font = font_winner

                # Área de Anexos no final
                start_row = 24
                ws.merge_cells(start_row=start_row, start_column=2, end_row=start_row, end_column=12)
                header = ws.cell(row=start_row, column=2, value="ÁREA DE ANEXOS - ORÇAMENTOS ORIGINAIS")
                header.fill = PatternFill(start_color="002060", fill_type="solid")
                header.font = Font(color="FFFFFF", bold=True)
                header.alignment = Alignment(horizontal="center", vertical="center")
                
                ws.merge_cells(start_row=start_row+1, start_column=2, end_row=start_row+1, end_column=12)
                ws.cell(row=start_row+1, column=2, value="Para anexar: Inserir -> Texto/Objeto -> Criar do Arquivo -> Marcar 'Exibir como ícone'").alignment = Alignment(horizontal="center")
                
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                st.success("Mapa gerado com sucesso!")
                st.download_button(label="📥 Baixar Mapa Preenchido", data=output, file_name=f"Mapa_Cotacao_{nome_material[:15]}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
            except Exception as e:
                st.error(f"Erro ao processar a planilha: {e}")
