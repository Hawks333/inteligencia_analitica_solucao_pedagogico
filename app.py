# ==================================================
# NOVA SESSÃO: RISCO DE REPROVAÇÃO PRESENCIAL (ATUALIZADA)
# ==================================================
elif funcao == "Risco de Reprovação Presencial":
    st.header("⚠️ Identificar Estudantes em Risco de Reprovação Presencial")
    st.info(
        "Carregue uma planilha (xlsx). O sistema detectará a coluna com CH (horas realizadas / horas totais) "
        "— tipicamente na coluna E ou em uma coluna cujo cabeçalho contenha 'CH'. "
        "O cálculo agora utiliza o denominador real presente no arquivo, respeitando eventuais variações individuais."
    )

    # Parâmetros do usuário
    meses = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
    hoje = datetime.datetime.now()
    mes_default_index = hoje.month - 1
    mes_selecionado = st.selectbox("Selecione o mês atual", meses, index=mes_default_index)
    carga_ideal = st.number_input("Carga horária ideal (padrão para quem não possui denominador no arquivo)", min_value=1, value=80, step=1)
    carga_ocorrida = st.number_input("Carga horária já ocorrida até o mês selecionado", min_value=0, value=36, step=1)
    st.write(f"Horas restantes possíveis no semestre (baseado na carga ideal): **{max(carga_ideal - carga_ocorrida, 0)}** horas")

    uploaded_file = st.file_uploader("Carregue a planilha (xlsx) com a coluna CH (ex: '28/84')", type=["xlsx"])
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
        except Exception as e:
            st.error(f"Erro ao ler o arquivo: {e}")
            st.stop()

        # Detectar coluna CH automaticamente
        ch_col = None
        for col in df.columns:
            if isinstance(col, str) and 'ch' in col.lower():
                ch_col = col
                break
        if ch_col is None and df.shape[1] >= 5:
            ch_col = df.columns[4]
            st.info(f"Coluna 'CH' não identificada pelo nome. Utilizando a 5ª coluna (E): '{ch_col}'")
        elif ch_col is None:
            st.error("Não foi possível localizar a coluna CH. Verifique o arquivo.")
            st.stop()

        # Função para extrair horas realizadas e total (aceita formatos variados: '28/84', '36 / 76', etc.)
        def parse_ch(val):
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

        parsed = df[ch_col].apply(parse_ch)
        df['Horas_Realizadas'] = parsed.apply(lambda x: x[0])
        df['Horas_Totais_Arquivo'] = parsed.apply(lambda x: x[1])

        # Substitui denominador ausente pelo valor padrão informado
        df['Horas_Totais_Usadas'] = df['Horas_Totais_Arquivo'].fillna(carga_ideal)

        # Cálculo das horas restantes para cada estudante, respeitando o denominador individual
        df['Horas_Restantes_Possiveis'] = df['Horas_Totais_Usadas'] - carga_ocorrida
        df.loc[df['Horas_Restantes_Possiveis'] < 0, 'Horas_Restantes_Possiveis'] = 0

        # Percentual atual e final possível
        df['Percentual_Atual'] = (df['Horas_Realizadas'] / df['Horas_Totais_Usadas']) * 100
        df['Max_Horas_Possiveis'] = df['Horas_Realizadas'].fillna(0) + df['Horas_Restantes_Possiveis']
        df['Percentual_Final_Possivel'] = (df['Max_Horas_Possiveis'] / df['Horas_Totais_Usadas']) * 100

        # Indicador de risco (<75%)
        df['Estudante em risco de reprovação presencial'] = df['Percentual_Final_Possivel'] < 75

        # Resumo e exibição
        total_alunos = len(df)
        total_risco = int(df['Estudante em risco de reprovação presencial'].sum())
        pct_risco = (total_risco / total_alunos * 100) if total_alunos > 0 else 0

        st.subheader("📋 Resultado — verificação de risco com denominador do arquivo")
        st.metric("Total de estudantes (linhas consideradas)", total_alunos)
        st.metric("Estudantes em risco (final possível < 75%)", f"{total_risco} ({pct_risco:.1f}%)")
        st.write(f"Carga já ocorrida: **{carga_ocorrida}h** — Denominador individual conforme arquivo — Mês: **{mes_selecionado}**")

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
        st.dataframe(df[cols_exibir].head(200))

        # Exportar com as novas colunas
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

        st.success("✅ Análise concluída com o denominador real de cada estudante.")
