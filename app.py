import streamlit as st
import pandas as pd
import io
from io import BytesIO
import numpy as np
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
import unicodedata
import datetime
import re

# Configuração da página
st.set_page_config(
    page_title="Sistema de Gestão Pedagógica",
    page_icon="📊",
    layout="wide"
)

# Título principal
st.title("🏫 Sistema de Gestão Pedagógica")
st.markdown("---")

# Menu de seleção de função
st.sidebar.title("🔧 Menu de Ferramentas")
funcao = st.sidebar.radio(
    "Selecione a funcionalidade:",
    ["Compilar Planilhas", "Reestruturar Relatório", "Busca Ativa de Estudantes", "Risco de Reprovação Presencial"]
)

# ==================================================
# FUNÇÃO 1: COMPILAR PLANILHAS
# ==================================================
if funcao == "Compilar Planilhas":
    st.header("📂 Compilar Múltiplas Planilhas")
    st.info("Faça upload de várias planilhas para compilar em um único arquivo.")
    
    # Upload de múltiplos arquivos
    uploaded_files = st.file_uploader(
        "Selecione todas as planilhas que deseja compilar:",
        type=["xlsx"],
        accept_multiple_files=True
    )
    
    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} arquivo(s) carregado(s) com sucesso!")
        
        # Função para processar cada planilha
        def processar_planilha(arquivo):
            df = pd.read_excel(arquivo, header=None)
            df = df.iloc[1:].reset_index(drop=True)
            df.columns = df.iloc[0]
            df = df.iloc[1:].reset_index(drop=True)
            return df
        
        dfs = []
        for uploaded_file in uploaded_files:
            df = processar_planilha(uploaded_file)
            dfs.append(df)
        
        if dfs:
            df_compilado = pd.concat(dfs, ignore_index=True)
            df_compilado = df_compilado.loc[:,~df_compilado.columns.duplicated()]
            
            st.subheader("📊 Preview do Arquivo Compilado")
            st.dataframe(df_compilado.head())
            
            towrite = BytesIO()
            df_compilado.to_excel(towrite, index=False)
            towrite.seek(0)
            
            st.download_button(
                label="⬇️ Baixar Arquivo Compilado",
                data=towrite,
                file_name="planilhas_compiladas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.success(f"✅ Compilação concluída! Total de linhas: {len(df_compilado)}")

# ==================================================
# FUNÇÃO 2: REESTRUTURAR RELATÓRIO
# ==================================================
elif funcao == "Reestruturar Relatório":
    st.header("🔄 Reestruturar Relatório Pedagógico")
    st.info("Processa o relatório compilado para formato de análise pedagógica.")
    
    uploaded_file = st.file_uploader("Carregue o relatório compilado", type=["xlsx"])
    
    if uploaded_file:
        preview = pd.read_excel(uploaded_file, header=None, nrows=10)
        header_linha = None
        for i, row in preview.iterrows():
            if all(col in row.values for col in ['DR', 'Polo', 'Nome']):
                header_linha = i
                break
        
        if header_linha is None:
            st.error("Não foi possível detectar o cabeçalho automaticamente.")
        else:
            df = pd.read_excel(uploaded_file, header=header_linha)
            df = df.dropna(how='all')
            
            colunas_necessarias = ['Nome', 'Atividades(tentativas/quantidade de tentativas)', 'Menção Atual']
            colunas_faltantes = [col for col in colunas_necessarias if col not in df.columns]
            
            if colunas_faltantes:
                st.error(f"Colunas faltantes: {colunas_faltantes}")
            else:
                df['Aluno_ID'] = df['Polo'] + ' - ' + df['Nome']
                df['Atividade'] = df['Atividades(tentativas/quantidade de tentativas)'].str.split('(').str[0].str.strip()
                
                def normalizar_nome(nome):
                    if pd.isna(nome):
                        return nome
                    nome = ''.join(c for c in unicodedata.normalize('NFD', str(nome)) 
                                 if unicodedata.category(c) != 'Mn')
                    nome = nome.lower().strip()
                    return nome
                
                df['Atividade_Normalizada'] = df['Atividade'].apply(normalizar_nome)
                
                tentativas = df['Atividades(tentativas/quantidade de tentativas)'].str.extract(r'\((\d+)/(\d+)\)')
                if not tentativas.empty:
                    df['Tentativas_Realizadas'] = tentativas[0].fillna(0).astype(int)
                    df['Tentativas_Total'] = tentativas[1].fillna(0).astype(int)
                
                pivot_mencoes = df.pivot_table(
                    index='Aluno_ID',
                    columns='Atividade_Normalizada',
                    values='Menção Atual',
                    aggfunc='first',
                    fill_value='--'
                ).reset_index()
                
                if 'Tentativas_Realizadas' in df.columns:
                    pivot_tentativas = df.pivot_table(
                        index='Aluno_ID',
                        columns='Atividade_Normalizada',
                        values='Tentativas_Realizadas',
                        aggfunc='first',
                        fill_value=0
                    ).reset_index()
                    pivot_tentativas.columns = ['Aluno_ID'] + [f'{col}_Tentativas' for col in pivot_tentativas.columns if col != 'Aluno_ID']
                
                colunas_aluno = ['Aluno_ID', 'DR', 'Polo', 'Nome', 'Etapa', 'Sala', 'Área de conhecimento',
                               'Data último acesso', 'Brasileiro(a)', 'Aluno AEE']
                colunas_aluno = [col for col in colunas_aluno if col in df.columns]
                info_alunos = df[colunas_aluno].drop_duplicates(subset=['Aluno_ID'])
                
                resultado = info_alunos.merge(pivot_mencoes, on='Aluno_ID', how='left')
                if 'Tentativas_Realizadas' in df.columns:
                    resultado = resultado.merge(pivot_tentativas, on='Aluno_ID', how='left')
                
                colunas_ordenadas = ['Aluno_ID', 'DR', 'Polo', 'Nome', 'Etapa', 'Sala', 'Área de conhecimento',
                                   'Data último acesso', 'Brasileiro(a)', 'Aluno AEE']
                colunas_ordenadas = [col for col in colunas_ordenadas if col in resultado.columns]
                colunas_atividades = [col for col in resultado.columns if col not in colunas_ordenadas and col != 'Aluno_ID']
                colunas_ordenadas.extend(colonas := colunas_atividades)
                resultado = resultado[colunas_ordenadas]
                
                towrite = BytesIO()
                resultado.to_excel(towrite, index=False)
                towrite.seek(0)
                
                st.subheader("📋 Resultado Processado")
                st.dataframe(resultado.head(3))
                
                st.download_button(
                    label="⬇️ Baixar Relatório Processado",
                    data=towrite,
                    file_name="relatorio_estruturado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.success(f"✅ Processamento concluído! {len(resultado)} alunos processados.")

# ==================================================
# FUNÇÃO 3: BUSCA ATIVA
# ==================================================
elif funcao == "Busca Ativa de Estudantes":
    st.header("🔍 Busca Ativa de Estudantes com Pendências")
    st.info("Identifica alunos com resultados pendentes por avaliativa.")
    
    uploaded_file = st.file_uploader("Carregue o relatório processado", type=["xlsx"])
    
    if uploaded_file:
        xls = pd.ExcelFile(uploaded_file)
        
        if len(xls.sheet_names) > 1:
            sheet_name = st.selectbox("Selecione a aba para processar", xls.sheet_names)
        else:
            sheet_name = xls.sheet_names[0]
        
        if sheet_name:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            
            # Avaliativas + opção "Todos"
            avaliativa = st.selectbox("Selecione a Avaliativa", ["Todos", 1, 2, 3, 4])
            
            if avaliativa != "Todos":
                colunas_avaliativa = []
                colunas_ignorar = []
                
                for col in df.columns:
                    if f"avaliativa {avaliativa}" in col.lower():
                        if "tentativas" in col.lower():
                            colunas_ignorar.append(col)
                        else:
                            colunas_avaliativa.append(col)
                
                if colunas_avaliativa:
                    st.success(f"✅ {len(colunas_avaliativa)} coluna(s) encontrada(s) para a Avaliativa {avaliativa}")
                    
                    colunas_adicionais = []
                    colunas_padrao = ["DR", "Polo", "Nome"]
                    for coluna in ["Etapa", "Sala", "Data último acesso"]:
                        if coluna in df.columns:
                            colunas_adicionais.append(coluna)
                    
                    todas_colunas = colunas_padrao + colunas_adicionais + colunas_avaliativa
                    
                    mask = df[colunas_avaliativa].apply(lambda x: x.astype(str).str.contains("--")).any(axis=1)
                    alunos_com_pendencia = df[mask][todas_colunas].copy()
                    
                    def identificar_areas_pendentes(row):
                        areas_pendentes = []
                        for col in colunas_avaliativa:
                            if str(row[col]).strip() == "--":
                                area = col.replace(f"Avaliativa {avaliativa}", "").strip()
                                if area.startswith(('-', '–', '—', ':')):
                                    area = area[1:].strip()
                                if area:
                                    areas_pendentes.append(area)
                        return ", ".join(areas_pendentes) if areas_pendentes else "Nenhuma"
                    
                    alunos_com_pendencia["Áreas com Pendência"] = alunos_com_pendencia.apply(identificar_areas_pendentes, axis=1)
                    alunos_com_pendencia = alunos_com_pendencia[alunos_com_pendencia["Áreas com Pendência"] != "Nenhuma"]
                
                else:
                    st.warning(f"❌ Nenhuma coluna encontrada para a Avaliativa {avaliativa}")
            
            else:
                colunas_atividades = [c for c in df.columns if "avaliativa" in c.lower() and "tentativas" not in c.lower()]
                
                if colunas_atividades:
                    st.success(f"✅ {len(colunas_atividades)} colunas de avaliativas consideradas")
                    
                    colunas_padrao = ["DR", "Polo", "Nome"]
                    colunas_adicionais = []
                    for coluna in ["Etapa", "Sala", "Data último acesso"]:
                        if coluna in df.columns:
                            colunas_adicionais.append(coluna)
                    
                    todas_colunas = colunas_padrao + colunas_adicionais + colunas_atividades
                    
                    mask = df[colunas_atividades].apply(lambda x: x.astype(str).str.contains("--")).all(axis=1)
                    alunos_com_pendencia = df[mask][todas_colunas].copy()
                    
                    alunos_com_pendencia["Áreas com Pendência"] = "Todas"
                
                else:
                    st.warning("❌ Nenhuma coluna de avaliativas encontrada.")
            
            # Exibição final
            if 'alunos_com_pendencia' in locals() and not alunos_com_pendencia.empty:
                st.subheader(f"🎯 Estudantes com Pendências - Avaliativa {avaliativa}")
                
                cols_to_show = ["DR", "Polo", "Nome"]
                for col in ["Etapa", "Sala", "Data último acesso"]:
                    if col in alunos_com_pendencia.columns:
                        cols_to_show.append(col)
                cols_to_show.append("Áreas com Pendência")
                
                st.dataframe(alunos_com_pendencia[cols_to_show])
                
                towrite = BytesIO()
                with pd.ExcelWriter(towrite, engine='openpyxl') as writer:
                    alunos_com_pendencia.to_excel(writer, index=False, sheet_name="Pendências")
                towrite.seek(0)
                
                st.download_button(
                    label="⬇️ Baixar Relatório de Pendências",
                    data=towrite,
                    file_name=f"pendencias_avaliativa_{avaliativa}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.subheader("📈 Estatísticas")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total de Alunos com Pendências", len(alunos_com_pendencia))
                with col2:
                    if "Etapa" in df.columns:
                        st.metric("Etapas Envolvidas", alunos_com_pendencia["Etapa"].nunique())
                with col3:
                    if "Sala" in df.columns:
                        st.metric("Salas Envolvidas", alunos_com_pendencia["Sala"].nunique())
                
                if avaliativa == "Todos":
                    st.bar_chart({"Sem Nenhuma Entrega": [len(alunos_com_pendencia)]})
                else:
                    areas_count = alunos_com_pendencia["Áreas com Pendência"].str.split(", ").explode().value_counts()
                    st.bar_chart(areas_count)
            
            else:
                st.info("🎉 Nenhum aluno com pendência encontrado!")

# ==================================================
# NOVA SESSÃO: RISCO DE REPROVAÇÃO PRESENCIAL (ATUALIZADA E OTIMIZADA)
# ==================================================
elif funcao == "Risco de Reprovação Presencial":
    import pandas as pd
    import numpy as np
    import re
    import datetime
    from io import BytesIO

    st.header("⚠️ Identificar Estudantes em Risco de Reprovação Presencial")

    st.info(
        "Carregue uma planilha (.xlsx). O sistema detectará automaticamente a coluna com CH "
        "(horas realizadas / horas totais) — normalmente na coluna E ou em uma coluna cujo cabeçalho contenha 'CH'. "
        "O cálculo usa o denominador real presente no arquivo, respeitando variações individuais."
    )

    # ==============================
    # PARÂMETROS DO USUÁRIO
    # ==============================
    meses = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]
    hoje = datetime.datetime.now()
    mes_default_index = hoje.month - 1

    mes_selecionado = st.selectbox("Selecione o mês atual", meses, index=mes_default_index)
    carga_ideal = st.number_input(
        "Carga horária ideal (padrão para quem não possui denominador no arquivo)",
        min_value=1, value=80, step=1
    )
    carga_ocorrida = st.number_input(
        "Carga horária já ocorrida até o mês selecionado",
        min_value=0, value=36, step=1
    )

    horas_restantes = max(carga_ideal - carga_ocorrida, 0)
    st.write(f"Horas restantes possíveis no semestre (baseado na carga ideal): **{horas_restantes}h**")

    # ==============================
    # UPLOAD DE ARQUIVO
    # ==============================
    uploaded_file = st.file_uploader(
        "Carregue a planilha (.xlsx) com a coluna CH (exemplo: '28/84')",
        type=["xlsx"]
    )

    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
        except Exception as e:
            st.error(f"Erro ao ler o arquivo: {e}")
            st.stop()

        # ==============================
        # DETECÇÃO AUTOMÁTICA DA COLUNA CH
        # ==============================
        ch_col = None
        for col in df.columns:
            if isinstance(col, str) and 'ch' in col.lower():
                ch_col = col
                break

        if ch_col is None:
            if df.shape[1] >= 5:
                ch_col = df.columns[4]
                st.info(f"Coluna 'CH' não identificada pelo nome. Utilizando a 5ª coluna (E): '{ch_col}'")
            else:
                st.error("Não foi possível localizar a coluna CH. Verifique se o arquivo possui uma coluna com valores como '28/84'.")
                st.stop()

        # ==============================
        # FUNÇÃO PARA INTERPRETAR OS VALORES DE CH
        # ==============================
        def parse_ch(val):
            """Aceita formatos como '28/84', '36 / 76', ou apenas '28'."""
            if pd.isna(val):
                return (np.nan, np.nan)
            s = str(val).strip()
            m = re.search(r'(\d+(?:[.,]\d+)?)\s*/\s*(\d+(?:[.,]\d+)?)', s)
            if m:
                num = float(m.group(1).replace(',', '.'))
                den = float(m.group(2).replace(',', '.'))
                return (num, den)
            parts = s.split('/')
            if len(parts) == 2:
                try:
                    num = float(parts[0].replace(',', '.'))
                    den = float(parts[1].replace(',', '.'))
                    return (num, den)
                except:
                    return (np.nan, np.nan)
            try:
                num = float(s.replace(',', '.'))
                return (num, np.nan)
            except:
                return (np.nan, np.nan)

        # Aplica a função de parsing
        parsed = df[ch_col].apply(parse_ch)
        df['Horas_Realizadas'] = parsed.apply(lambda x: x[0])
        df['Horas_Totais_Arquivo'] = parsed.apply(lambda x: x[1])

        # Substitui denominador ausente pelo valor padrão informado
        df['Horas_Totais_Usadas'] = df['Horas_Totais_Arquivo'].fillna(carga_ideal)

        # Evita divisões por zero
        df.loc[df['Horas_Totais_Usadas'] <= 0, 'Horas_Totais_Usadas'] = carga_ideal

        # ==============================
        # CÁLCULOS
        # ==============================
        df['Horas_Restantes_Possiveis'] = df['Horas_Totais_Usadas'] - carga_ocorrida
        df.loc[df['Horas_Restantes_Possiveis'] < 0, 'Horas_Restantes_Possiveis'] = 0

        df['Percentual_Atual'] = (df['Horas_Realizadas'] / df['Horas_Totais_Usadas']) * 100
        df['Max_Horas_Possiveis'] = df['Horas_Realizadas'].fillna(0) + df['Horas_Restantes_Possiveis']
        df['Percentual_Final_Possivel'] = (df['Max_Horas_Possiveis'] / df['Horas_Totais_Usadas']) * 100

        # Indicador de risco (<75%)
        df['Estudante em risco de reprovação presencial'] = df['Percentual_Final_Possivel'] < 75

        # ==============================
        # RESUMO
        # ==============================
        total_alunos = len(df)
        total_risco = int(df['Estudante em risco de reprovação presencial'].sum())
        pct_risco = (total_risco / total_alunos * 100) if total_alunos > 0 else 0

        st.subheader("📋 Resultado — Verificação de Risco com Denominador Real")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total de estudantes (linhas consideradas)", total_alunos)
        col2.metric("Estudantes em risco (final possível < 75%)", total_risco)
        col3.metric("Percentual em risco", f"{pct_risco:.1f}%")

        st.write(f"**Mês:** {mes_selecionado} — **Carga ocorrida:** {carga_ocorrida}h — "
                 f"**Denominador individual conforme arquivo.**")

        # ==============================
        # EXIBIÇÃO DA TABELA
        # ==============================
        cols_exibir = []
        for c in ["DR", "Polo", "Nome", "Etapa", "Sala", "Data último acesso"]:
            if c in df.columns:
                cols_exibir.append(c)
        cols_exibir += [
            ch_col, 'Horas_Realizadas', 'Horas_Totais_Arquivo', 'Horas_Totais_Usadas',
            'Percentual_Atual', 'Max_Horas_Possiveis', 'Percentual_Final_Possivel',
            'Estudante em risco de reprovação presencial'
        ]
        cols_exibir = [c for c in cols_exibir if c in df.columns]

        # Formata percentuais
        df['Percentual_Atual'] = df['Percentual_Atual'].round(1)
        df['Percentual_Final_Possivel'] = df['Percentual_Final_Possivel'].round(1)

        st.dataframe(df[cols_exibir].head(200))

        # ==============================
        # EXPORTAÇÃO PARA EXCEL
        # ==============================
        towrite = BytesIO()
        with pd.ExcelWriter(towrite, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Risco_Reprovacao")
        towrite.seek(0)

        st.download_button(
            label="⬇️ Baixar Relatório com Indicador de Risco (ajustado)",
            data=towrite,
            file_name="relatorio_risco_reprovacao_presencial.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.success("✅ Análise concluída com sucesso, considerando o denominador real de cada estudante.")

