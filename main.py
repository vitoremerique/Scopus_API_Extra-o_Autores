import pandas as pd
import os
import sys
import pybliometrics
from pybliometrics.scopus import AbstractRetrieval, AuthorRetrieval

# --- 1. Configuração Inicial e Verificação de API ---
def verificar_configuracao():
    """
    Verifica se o pybliometrics está configurado.
    Se não estiver, inicia o processo de configuração para pedir a API Key.
    """
    try:
        # Tenta acessar um valor de configuração para ver se existe
        if not pybliometrics.scopus.config.has_section('Authentication'):
             raise Exception("Config not found")
    except Exception:
        print("\n" + "="*60)
        print("⚠️  CONFIGURAÇÃO NECESSÁRIA (Primeira execução) ⚠️")
        print("O arquivo de configuração do Scopus não foi encontrado.")
        print("Tenha sua API Key pronta (pegue em https://dev.elsevier.com/)")
        print("="*60 + "\n")
        # Inicia o prompt interativo para inserir a chave
        pybliometrics.scopus.init()

# --- 2. Funções Auxiliares ---

def limpar_doi(doi_bruto):
    """
    Remove prefixos de URL e espaços em branco do DOI.
    Ex: 'https://doi.org/10.1016/xyz' -> '10.1016/xyz'
    """
    if not isinstance(doi_bruto, str):
        return str(doi_bruto)
    
    doi = doi_bruto.strip()
    # Remove prefixos comuns que causam erro na API
    doi = doi.replace("https://doi.org/", "")
    doi = doi.replace("http://doi.org/", "")
    doi = doi.replace("http://dx.doi.org/", "")
    doi = doi.replace("dx.doi.org/", "")
    return doi

def get_paper_authors_stats(doi, id_planilha=None):
    """
    Busca metadados do artigo e estatísticas dos autores.
    """
    doi_limpo = limpar_doi(doi)
    print(f"--- Processando: {doi_limpo} ---")
    
    try:
        # A. Busca os metadados do artigo
        try:
            paper = AbstractRetrieval(doi_limpo, view='FULL')
        except Exception as e:
            # Tratamento genérico compatível com todas as versões
            erro_str = str(e)
            if "404" in erro_str or "NOT_FOUND" in erro_str:
                print(f"   [!] Artigo não encontrado no Scopus (404).")
                return []
            elif "401" in erro_str or "UNAUTHORIZED" in erro_str:
                print(f"   [!] Erro de Autenticação (401). Verifique sua API Key ou conexão.")
                return []
            else:
                print(f"   [!] Erro ao buscar artigo: {e}")
                return []

        titulo = paper.title
        
        if not paper.authors:
            print("   [i] Nenhum autor listado nos metadados.")
            return []

        authors_list = []

        # B. Itera sobre cada autor
        for auth in paper.authors:
            author_id = auth.auid
            author_name = f"{auth.given_name} {auth.surname}"
            
            author_data = {
                "ID_Planilha": id_planilha,
                "DOI_Original": doi, 
                "DOI_Limpo": doi_limpo,
                "Titulo_Artigo": titulo,
                "Nome": author_name,
                "Scopus_ID": author_id,
                "Total_Papers": 0
            }

            # C. Busca perfil detalhado do autor
            try:
                profile = AuthorRetrieval(author_id)
                author_data["Total_Papers"] = profile.document_count
            except Exception as e:
                # Tratamento de erro genérico para o perfil do autor
                if "404" in str(e):
                    author_data["Total_Papers"] = "N/A (Perfil não achado)"
                else:
                    author_data["Total_Papers"] = "Erro na busca"

            authors_list.append(author_data)

        print(f"   -> Sucesso: {len(authors_list)} autores processados.")
        return authors_list

    except Exception as e:
        print(f"Erro fatal neste DOI: {e}")
        return []

def ler_dois_de_xlsx(arquivo_xlsx, coluna_doi="doi", coluna_id="id"):
    """
    Lê o Excel e retorna lista de tuplas (id_original, doi).
    Se a coluna de ID não existir, usa o índice da linha como ID.
    """
    try:
        df = pd.read_excel(arquivo_xlsx)
        
        # Normaliza nomes das colunas (tudo minúsculo, sem espaços)
        original_columns = list(df.columns)
        df.columns = [str(c).lower().strip() for c in df.columns]
        coluna_doi = coluna_doi.lower().strip()
        coluna_id = coluna_id.lower().strip()

        if coluna_doi not in df.columns:
            print(f"Erro: Coluna '{coluna_doi}' não encontrada.")
            print(f"Colunas disponíveis: {list(df.columns)}")
            return []

        # Prepara série de IDs (usa coluna se existir, senão índice)
        if coluna_id in df.columns:
            ids = df[coluna_id].fillna('').astype(str).tolist()
        else:
            ids = [str(i+1) for i in range(len(df))]

        dois = df[coluna_doi].fillna('').astype(str).tolist()

        # Junta em tuplas (id, doi), filtrando DOIs vazios
        resultado = [(ids[i].strip(), dois[i].strip()) for i in range(len(dois)) if str(dois[i]).strip()]
        print(f"Arquivo carregado. Total de linhas com DOI: {len(resultado)}\n")
        return resultado
        
    except FileNotFoundError:
        print(f"Erro: Arquivo '{arquivo_xlsx}' não encontrado.")
        return []
    except Exception as e:
        print(f"Erro ao ler Excel: {e}")
        return []


def formatar_autores_por_artigo(df, chave='DOI_Limpo'):
    """
    Cria a coluna 'Autores_Formatados' no DataFrame `df` agrupando autores por `chave`.
    Formato por autor: 'Sobrenome, Nome, ID:<scopus_id>, (<Total_Papers>)'
    Autores separados por ' and '.
    """
    if chave not in df.columns:
        # Tenta usar DOI_Original como chave alternativa
        if 'DOI_Original' in df.columns:
            chave = 'DOI_Original'
        else:
            # Sem chave aplicável, cria coluna vazia e retorna
            df['Autores_Formatados'] = ''
            return df

    def formatar_grupo(g):
        partes = []
        for _, row in g.iterrows():
            nome = str(row.get('Nome', '')).strip()
            scopus_id = str(row.get('Scopus_ID', '')).strip()
            total = row.get('Total_Papers', '')

            # Quebra o nome em Given(s) + Surname (assume surname = último token)
            tokens = nome.split()
            if len(tokens) == 0:
                given = ''
                surname = ''
            elif len(tokens) == 1:
                given = tokens[0]
                surname = ''
            else:
                surname = tokens[-1]
                given = ' '.join(tokens[:-1])

            parte = f"{surname}, {given}, ID:{scopus_id}, ({total})"
            partes.append(parte)

        return ' and '.join(partes)

    agrupados = df.groupby(chave, sort=False).apply(formatar_grupo)
    mapping = agrupados.to_dict()
    df['Autores_Formatados'] = df[chave].map(mapping)
    return df

# --- 3. Execução Principal ---

if __name__ == "__main__":
    # Passo 1: Verifica se tem a chave de API
    verificar_configuracao()

    # Configurações de arquivo
    ARQUIVO_ENTRADA = "Scopus Teste.xlsx"
    ARQUIVO_SAIDA = "autores_scopus_completo.csv"
    
    # Lista para armazenar todos os dados
    todos_autores = []

    # Passo 2: Carregar DOIs
    if os.path.exists(ARQUIVO_ENTRADA):
        print(f"Lendo DOIs de '{ARQUIVO_ENTRADA}'...")
        lista_dois = ler_dois_de_xlsx(ARQUIVO_ENTRADA)
        
        # Carrega DOIs já processados (se houver) para evitar requisições duplicadas
        processed_dois = set()
        if os.path.exists(ARQUIVO_SAIDA):
            try:
                df_exist = pd.read_csv(ARQUIVO_SAIDA, dtype=str, encoding='utf-8-sig')
                cols = [c.lower() for c in df_exist.columns]
                # Procura possíveis nomes de coluna que contenham DOI limpo
                if 'doi_limpo' in cols:
                    processed_dois.update(df_exist.iloc[:, cols.index('doi_limpo')].dropna().astype(str).str.strip().tolist())
                elif 'doi' in cols:
                    processed_dois.update(df_exist.iloc[:, cols.index('doi')].dropna().astype(str).str.strip().tolist())
            except Exception:
                # Se não conseguir ler, continua sem lista de processados
                processed_dois = set()

        # Passo 3: Loop principal (pula DOIs já processados)
        for i, item in enumerate(lista_dois, 1):
            id_planilha, doi = item
            doi_limpo = limpar_doi(doi)
            print(f"\n[{i}/{len(lista_dois)}] Iniciando: {doi_limpo} (ID_planilha={id_planilha})...")

            # Verifica se já existe no CSV de saída
            if doi_limpo in processed_dois:
                print(f"   [i] Pulando: DOI já processado e presente em '{ARQUIVO_SAIDA}'.")
                continue

            # Verifica se já foi adicionado na sessão atual (evita duplicados dentro do mesmo run)
            if any((rec.get('DOI_Limpo') == doi_limpo or rec.get('DOI_Limpo') == doi_limpo) for rec in todos_autores):
                print(f"   [i] Pulando: DOI duplicado na execução atual.")
                continue

            # Chama a função principal
            dados_artigo = get_paper_authors_stats(doi, id_planilha=id_planilha)
            
            if dados_artigo:
                todos_autores.extend(dados_artigo)
        
        # Passo 4: Salvar resultados
        if todos_autores:
            df_final = pd.DataFrame(todos_autores)
            
            print("\n" + "="*50)
            print(f"Processamento concluído! Total de registros: {len(df_final)}")
            print("="*50)
            
            # Mostra prévia das colunas que existem
            colunas_preview = ['ID_Planilha', 'DOI_Limpo', 'Nome', 'Scopus_ID', 'Total_Papers']
            cols_to_show = [c for c in colunas_preview if c in df_final.columns]
            print(df_final[cols_to_show].head())
            
            try:
                # Se o arquivo já existe, lê-o, concatena com os novos registros,
                # ordena por ID_Planilha (numérico quando possível) e salva sobrescrevendo.
                if os.path.exists(ARQUIVO_SAIDA):
                    df_exist = pd.read_csv(ARQUIVO_SAIDA, dtype=str, encoding='utf-8-sig')
                    df_comb = pd.concat([df_exist, df_final], ignore_index=True, sort=False)

                    # Gera a coluna formatada para o conjunto combinado
                    df_comb = formatar_autores_por_artigo(df_comb, chave='DOI_Limpo')

                    if 'ID_Planilha' in df_comb.columns:
                        # Tenta ordenar numericamente quando possível, senão por string
                        try:
                            df_comb['__id_num'] = pd.to_numeric(df_comb['ID_Planilha'], errors='coerce')
                            df_comb = df_comb.sort_values(by=['__id_num', 'DOI_Limpo'], na_position='last')
                            df_comb = df_comb.drop(columns=['__id_num'])
                        except Exception:
                            df_comb = df_comb.sort_values(by=['ID_Planilha', 'DOI_Limpo'], na_position='last')
                    else:
                        df_comb = df_comb.sort_values(by='DOI_Limpo', na_position='last')

                    # Salva o arquivo combinado ordenado (sobrescreve)
                    df_comb.to_csv(ARQUIVO_SAIDA, index=False, encoding='utf-8-sig')
                    print(f"\n✅ Dados combinados e salvos com sucesso em: '{ARQUIVO_SAIDA}'")

                    # Gera arquivo resumido (uma linha por artigo/DOI)
                    try:
                        resumo_cols = []
                        if 'ID_Planilha' in df_comb.columns:
                            resumo_cols.append('ID_Planilha')
                        if 'DOI_Limpo' in df_comb.columns:
                            resumo_cols.append('DOI_Limpo')
                        elif 'DOI_Original' in df_comb.columns:
                            resumo_cols.append('DOI_Original')
                        if 'Titulo_Artigo' in df_comb.columns:
                            resumo_cols.append('Titulo_Artigo')
                        resumo_cols.append('Autores_Formatados')

                        df_resumo = df_comb.drop_duplicates(subset=['DOI_Limpo'] if 'DOI_Limpo' in df_comb.columns else ['DOI_Original'])
                        df_resumo = df_resumo.loc[:, [c for c in resumo_cols if c in df_resumo.columns]]
                        # Ordena resumo por ID quando existir
                        if 'ID_Planilha' in df_resumo.columns:
                            try:
                                df_resumo['__id_num'] = pd.to_numeric(df_resumo['ID_Planilha'], errors='coerce')
                                df_resumo = df_resumo.sort_values(by='__id_num', na_position='last')
                                df_resumo = df_resumo.drop(columns=['__id_num'])
                            except Exception:
                                df_resumo = df_resumo.sort_values(by='ID_Planilha', na_position='last')

                        resumo_file = 'autores_scopus_resumido.csv'
                        df_resumo.to_csv(resumo_file, index=False, encoding='utf-8-sig')
                        print(f"\n✅ Arquivo resumido salvo em: '{resumo_file}'")
                    except Exception as e:
                        print(f"\n⚠️ Não foi possível gerar o resumo: {e}")
                else:
                    # Arquivo não existe: gera a coluna formatada e salva normalmente
                    df_final = formatar_autores_por_artigo(df_final, chave='DOI_Limpo')

                    # Ordena df_final por ID_Planilha antes de criar
                    if 'ID_Planilha' in df_final.columns:
                        try:
                            df_final['__id_num'] = pd.to_numeric(df_final['ID_Planilha'], errors='coerce')
                            df_final = df_final.sort_values(by='__id_num', na_position='last')
                            df_final = df_final.drop(columns=['__id_num'])
                        except Exception:
                            df_final = df_final.sort_values(by='ID_Planilha', na_position='last')

                    df_final.to_csv(ARQUIVO_SAIDA, index=False, encoding='utf-8-sig')
                    print(f"\n✅ Dados salvos com sucesso em: '{ARQUIVO_SAIDA}'")

                    # Gera resumo para o caso de criação do arquivo
                    try:
                        df_resumo = df_final.drop_duplicates(subset=['DOI_Limpo'] if 'DOI_Limpo' in df_final.columns else ['DOI_Original'])
                        resumo_cols = []
                        if 'ID_Planilha' in df_resumo.columns:
                            resumo_cols.append('ID_Planilha')
                        if 'DOI_Limpo' in df_resumo.columns:
                            resumo_cols.append('DOI_Limpo')
                        elif 'DOI_Original' in df_resumo.columns:
                            resumo_cols.append('DOI_Original')
                        if 'Titulo_Artigo' in df_resumo.columns:
                            resumo_cols.append('Titulo_Artigo')
                        resumo_cols.append('Autores_Formatados')
                        df_resumo = df_resumo.loc[:, [c for c in resumo_cols if c in df_resumo.columns]]
                        resumo_file = 'autores_scopus_resumido.csv'
                        df_resumo.to_csv(resumo_file, index=False, encoding='utf-8-sig')
                        print(f"\n✅ Arquivo resumido salvo em: '{resumo_file}'")
                    except Exception as e:
                        print(f"\n⚠️ Não foi possível gerar o resumo: {e}")
            except PermissionError:
                print(f"\n❌ Erro: Não foi possível salvar em '{ARQUIVO_SAIDA}'. O arquivo está aberto? Feche-o e tente novamente.")
        else:
            print("\nNenhum dado foi coletado. Verifique os DOIs ou sua conexão.")
            
    else:
        print(f"Arquivo de entrada '{ARQUIVO_ENTRADA}' não encontrado na pasta.")
        print("Por favor, coloque o arquivo Excel na mesma pasta deste script.")