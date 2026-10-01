import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import PyPDF2
import re
import io

st.set_page_config(page_title="Mapa de Cotação Automático", page_icon="🏗️")

st.title("🏗️ Gerador de Mapa de Cotação")
st.write("Anexe o seu modelo (já com o cabeçalho preenchido se preferir) e os orçamentos.")

# Função de extração de valores do PDF
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
            
        return nome_sugerido.upper(), maior_valor
    except Exception:
        return pdf_file.name.upper(), 0.0

# 1. Uploads
modelo_excel = st.file_uploader("1. Anexe o Modelo Excel (Ex: MPC - Modelo.xlsx)", type=["xlsx"])
fornecedores_pdfs = st.file_uploader("2. Anexe os PDFs dos Fornecedores (Máx 3)", type=["pdf"], accept_multiple_files=True)

# 2. Dados Manuais Opcionais (Caso a planilha base esteja em branco)
st.markdown("### Dados da Cotação (Opcional)")
st.write("Se o seu Excel já estiver preenchido com o material e quantidade, pode deixar isto em branco.")
col_mat, col_qtd = st.columns([3, 1])
with col_mat:
    nome_material = st.text_input("Qual é o material/serviço sendo cotado?")
with col_qtd:
    quantidade_item = st.number_input("Quantidade", value=1, min_value=1)

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

    if st.button("Gerar Planilha Final (EXATA)"):
        try:
            wb = openpyxl.load_workbook(modelo_excel)
            ws = wb.active
            
            # Formatação de Contabilidade exata que o Excel usa
            formato_moeda = '_-"R$"* #,##0.00_-;\-"R$"* #,##0.00_-;_-"R$"* "-"??_-;_-@_-'
            
            # Preencher Nome do Material e Quantidade se o utilizador tiver digitado no site
            if nome_material:
                ws.cell(row=2, column=5, value=nome_material.upper()) # Cabeçalho
                ws.cell(row=8, column=2, value=nome_material.upper()) # Tabela Item
            if quantidade_item > 1:
                ws.cell(row=8, column=4, value=quantidade_item) # Quantidade Tabela
            
            # Encontrar o menor preço para destacar
            valores = [f['valor'] for f in dados_fornecedores if f['valor'] > 0]
            menor_preco = min(valores) if valores else 0
            
            # Colunas dos fornecedores no modelo
            # Forn 1: Nome=G2, Unit=G8, Total=H8
            # Forn 2: Nome=I2, Unit=I8, Total=J8
            # Forn 3: Nome=K2, Unit=K8, Total=L8
            colunas_fornecedores = [
                {'nome': 7, 'unit': 7, 'total': 8},
                {'nome': 9, 'unit': 9, 'total': 10},
                {'nome': 11, 'unit': 11, 'total': 12}
            ]
            
            for i, fornecedor in enumerate(dados_fornecedores):
                cols = colunas_fornecedores[i]
                
                # 1. Preenche Nome do Fornecedor na linha 2
                ws.cell(row=2, column=cols['nome'], value=fornecedor['nome'])
                
                # 2. Preenche Preço Unitário na linha 8 e aplica formatação de moeda
                celula_unit = ws.cell(row=8, column=cols['unit'])
                celula_unit.value = fornecedor['valor']
                celula_unit.number_format = formato_moeda
                
                # 3. CRUCIAL: Insere a fórmula de TOTAL multiplicando pela Quantidade (D8)
                letra_col_unit = openpyxl.utils.get_column_letter(cols['unit'])
                celula_total = ws.cell(row=8, column=cols['total'])
                celula_total.value = f"=D8*{letra_col_unit}8"
                celula_total.number_format = formato_moeda
                
                # 4. Destacar Vencedor a Verde (Exatamente como no seu exemplo)
                if fornecedor['valor'] == menor_preco and fornecedor['valor'] > 0:
                    # Verde Claro Padrão do Excel (Index 9 no preenchimento do Openpyxl)
                    green_fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
                    font_winner = Font(color="006100", bold=True)
                    
                    celula_unit.fill = green_fill
                    celula_unit.font = font_winner
                    
                    celula_total.fill = green_fill
                    celula_total.font = font_winner

            # Área de Anexos adicionada 15 linhas abaixo do fim dos itens para não estragar nada
            linha_anexos = 24
            ws.merge_cells(start_row=linha_anexos, start_column=2, end_row=linha_anexos, end_column=12)
            header = ws.cell(row=linha_anexos, column=2, value="ÁREA DE ANEXOS - ORÇAMENTOS ORIGINAIS")
            header.fill = PatternFill(start_color="002060", fill_type="solid")
            header.font = Font(color="FFFFFF", bold=True)
            header.alignment = Alignment(horizontal="center", vertical="center")
            
            ws.merge_cells(start_row=linha_anexos+1, start_column=2, end_row=linha_anexos+1, end_column=12)
            ws.cell(row=linha_anexos+1, column=2, value="Para anexar: Inserir -> Texto/Objeto -> Criar do Arquivo -> Marcar 'Exibir como ícone'").alignment = Alignment(horizontal="center")
            
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.success("Planilha EXACTA gerada com sucesso!")
            st.download_button(label="📥 Baixar Planilha Final", data=output, file_name="MPC_Finalizado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
        except Exception as e:
            st.error(f"Erro ao processar a planilha: {e}")
