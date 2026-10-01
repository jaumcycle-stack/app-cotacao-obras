import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import io

# Configuração da página para ecrã inteiro
st.set_page_config(page_title="Mapa de Cotação Definitivo", page_icon="🏗️", layout="wide")

st.title("🏗️ Gerador de Mapa de Cotação - Completo")
st.write("Preencha todos os campos do orçamento. O aplicativo respeitará 100% da estrutura da sua planilha.")

# 1. Upload do Modelo
st.subheader("1. Modelo Excel")
modelo_excel = st.file_uploader("Anexe o Modelo (MPC - Modelo.xlsx)", type=["xlsx"])

st.divider()

# 2. Dados Globais da Cotação
st.subheader("2. Dados do Material Cotado")
col_mat, col_unid, col_qtd = st.columns([3, 1, 1])
with col_mat:
    nome_material = st.text_input("Nome do Material (Ex: TUBOS DE ESGOTO 300MM)")
with col_unid:
    unidade = st.text_input("Unid. Medida (Ex: Unid., VB, Pct, M)", value="Unid.")
with col_qtd:
    quantidade = st.number_input("Quantidade", value=1.0, min_value=0.01)

st.divider()

# 3. Dados dos Fornecedores (Lado a Lado)
st.subheader("3. Dados dos Fornecedores")
fornecedores = []
cols_forn = st.columns(3)

for i in range(3):
    with cols_forn[i]:
        st.markdown(f"### Fornecedor {i+1}")
        nome = st.text_input("Nome da Empresa", key=f"nome_{i}")
        valor = st.number_input("Valor Unitário (R$)", value=0.0, format="%.2f", key=f"val_{i}")
        frete = st.text_input("Frete (Ex: CIF, FOB, Incluso)", key=f"frete_{i}")
        pagamento = st.text_input("Condição de Pagto (Ex: 15 Dias)", key=f"pagto_{i}")
        prazo = st.text_input("Prazo de Entrega (Ex: 10 Dias)", key=f"prazo_{i}")
        
        fornecedores.append({
            "nome": nome,
            "valor": valor,
            "frete": frete,
            "pagamento": pagamento,
            "prazo": prazo
        })

st.divider()

# 4. Geração do Excel
if st.button("Gerar Planilha Final", type="primary", use_container_width=True):
    if not modelo_excel:
        st.error("ERRO: Por favor, anexe o modelo Excel no Passo 1.")
    elif not nome_material:
        st.error("ERRO: Por favor, preencha o Nome do Material no Passo 2.")
    else:
        try:
            wb = openpyxl.load_workbook(modelo_excel)
            ws = wb.active
            
            formato_moeda = '_-"R$"* #,##0.00_-;\-"R$"* #,##0.00_-;_-"R$"* "-"??_-;_-@_-'
            
            # 1. Preencher Cabeçalho e Dados do Item
            ws.cell(row=2, column=5, value=nome_material.upper())
            ws.cell(row=8, column=2, value=nome_material.upper())
            ws.cell(row=8, column=3, value=unidade)
            ws.cell(row=8, column=4, value=quantidade)
            
            # Descobrir menor preço válido
            valores_validos = [f['valor'] for f in fornecedores if f['valor'] > 0]
            menor_preco = min(valores_validos) if valores_validos else 0
            
            # Mapeamento: F1 (G/H), F2 (I/J), F3 (K/L)
            mapa_colunas = [{'base': 7}, {'base': 9}, {'base': 11}]
            
            for i, forn in enumerate(fornecedores):
                col_unit = mapa_colunas[i]['base']
                col_tot = col_unit + 1
                letra_unit = openpyxl.utils.get_column_letter(col_unit)
                letra_tot = openpyxl.utils.get_column_letter(col_tot)
                
                # Nome Fornecedor
                if forn['nome']:
                    ws.cell(row=2, column=col_unit, value=forn['nome'].upper())
                
                # Preços e Fórmulas
                if forn['valor'] > 0:
                    c_unit = ws.cell(row=8, column=col_unit, value=forn['valor'])
                    c_unit.number_format = formato_moeda
                    
                    # Fórmula de Total (Quantidade * Unitário)
                    c_tot = ws.cell(row=8, column=col_tot, value=f"=D8*{letra_unit}8")
                    c_tot.number_format = formato_moeda
                    
                    # Subtotal (Linha 9) e Total Final (Linha 17)
                    ws.cell(row=9, column=col_tot, value=f"=SUM({letra_tot}8:{letra_tot}8)").number_format = formato_moeda
                    ws.cell(row=17, column=col_tot, value=f"={letra_tot}9").number_format = formato_moeda
                    
                    # Informações Complementares
                    if forn['frete']: ws.cell(row=13, column=col_unit, value=forn['frete'].upper())
                    if forn['pagamento']: ws.cell(row=14, column=col_unit, value=forn['pagamento'].upper())
                    if forn['prazo']: ws.cell(row=15, column=col_unit, value=forn['prazo'].upper())
                    
                    # Destacar Vencedor a Verde
                    if forn['valor'] == menor_preco:
                        green_fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
                        font_winner = Font(color="006100", bold=True)
                        c_unit.fill = green_fill
                        c_unit.font = font_winner
                        c_tot.fill = green_fill
                        c_tot.font = font_winner
            
            # 5. Criar Área de Anexos no final da folha
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
            
            st.success("✅ Planilha gerada com sucesso e com todos os dados!")
            st.download_button(
                label="📥 Baixar Planilha Final (Perfeita)", 
                data=output, 
                file_name="MPC_Completo.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Erro ao processar a planilha: {e}")
