"""
Script para classificação de comentários em categorias usando OpenAI GPT-4-Turbo
Autor: Análise de Comentários Grupo NC
Data: 2025-10-17
"""

import os
import sys
import json
from pathlib import Path
from dotenv import load_dotenv
from openai import OpenAI
import openpyxl
from openpyxl import load_workbook
import pandas as pd
from time import sleep

# Configurações
ARQUIVO_EXCEL = "Teste Avaliação de Comentários.xlsx"
SHEET_NAME = "Comentários de Produtos"
MODELO_OPENAI = "gpt-5-nano"  # Modelo mais economico e rapido

# Configuração de log
# True = mostra cada comentário sendo processado (detalhado)
# False = mostra apenas progresso a cada 50 comentários (resumido)
LOG_DETALHADO = True

# Categorias disponíveis
CATEGORIAS = ["Design", "Qualidade", "Preço", "Durabilidade", "Logística"]


def carregar_configuracao():
    """Carrega variáveis de ambiente e valida configurações"""
    print("=" * 70)
    print("CLASSIFICACAO DE COMENTARIOS EM CATEGORIAS")
    print("=" * 70)
    print("\nCarregando configuracoes...")
    
    # Carregar .env
    load_dotenv()
    
    # Validar API Key
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key or api_key == "your_api_key_here":
        print("ERRO: OPENAI_API_KEY nao encontrada no arquivo .env")
        print("Por favor, adicione sua chave da OpenAI no arquivo .env")
        sys.exit(1)
    
    # Validar arquivo Excel
    if not Path(ARQUIVO_EXCEL).exists():
        print(f"ERRO: Arquivo '{ARQUIVO_EXCEL}' nao encontrado")
        sys.exit(1)
    
    print("Configuracoes carregadas com sucesso\n")
    return api_key


def ler_comentarios_excel():
    """Lê comentários do Excel"""
    print(f"Lendo dados de '{SHEET_NAME}'...")
    
    try:
        # Ler Excel
        df = pd.read_excel(ARQUIVO_EXCEL, sheet_name=SHEET_NAME)
        
        # Validar colunas
        if 'Produto' not in df.columns or 'Comentário' not in df.columns:
            print("ERRO: Colunas 'Produto' e 'Comentario' nao encontradas")
            sys.exit(1)
        
        # Remover linhas com valores nulos
        df = df.dropna(subset=['Produto', 'Comentário'])
        
        total_comentarios = len(df)
        print(f"Total de comentarios a classificar: {total_comentarios}\n")
        
        return df
    
    except Exception as e:
        print(f"ERRO ao ler Excel: {e}")
        sys.exit(1)


def classificar_comentario_com_openai(cliente_openai, comentario, max_retries=3):
    """Classifica um comentário usando OpenAI"""
    
    categorias_str = ", ".join(CATEGORIAS)
    
    prompt = f"""Analise o seguinte comentário sobre um produto e classifique-o em UMA OU MAIS das seguintes categorias:

CATEGORIAS DISPONÍVEIS:
- Design: Aparência, estética, beleza, estilo visual do produto
- Qualidade: Qualidade geral, acabamento, materiais, construção
- Preço: Valor, custo-benefício, preço alto/baixo, vale a pena
- Durabilidade: Resistência, durabilidade ao longo do tempo, quebra fácil/difícil
- Logística: Entrega, envio, embalagem, transporte, prazo

COMENTÁRIO:
"{comentario}"

INSTRUÇÕES:
1. Um comentário pode ter MÚLTIPLAS categorias se mencionar vários aspectos
2. Retorne APENAS as categorias relevantes mencionadas no comentário
3. Se o comentário não se encaixar claramente em nenhuma categoria, retorne a mais próxima
4. Responda APENAS em formato JSON

Responda no formato JSON:
{{
  "categorias": ["Categoria1", "Categoria2"]
}}"""

    for tentativa in range(max_retries):
        try:
            # Chamar API OpenAI
            response = cliente_openai.chat.completions.create(
                model=MODELO_OPENAI,
                messages=[
                    {"role": "system", "content": "Você é um assistente especializado em classificação de comentários de produtos. Sempre responda em português do Brasil e em formato JSON válido."},
                    {"role": "user", "content": prompt}
                ]
                # temperature removido - GPT-5 usa valor padrao (1)
            )
            
            # Extrair resposta
            resposta_texto = response.choices[0].message.content.strip()
            
            # Tentar parsear JSON
            # Remover possíveis markdown code blocks
            if resposta_texto.startswith("```"):
                resposta_texto = resposta_texto.split("```")[1]
                if resposta_texto.startswith("json"):
                    resposta_texto = resposta_texto[4:]
                resposta_texto = resposta_texto.strip()
            
            resultado = json.loads(resposta_texto)
            
            # Validar categorias
            categorias_encontradas = resultado.get("categorias", [])
            
            # Filtrar apenas categorias válidas
            categorias_validas = [cat for cat in categorias_encontradas if cat in CATEGORIAS]
            
            if not categorias_validas:
                # Se não houver categorias válidas, tentar novamente
                if tentativa < max_retries - 1:
                    sleep(0.5)
                    continue
                else:
                    return ["Qualidade"]  # Categoria padrão
            
            return categorias_validas
        
        except json.JSONDecodeError as e:
            if tentativa < max_retries - 1:
                sleep(0.5)
                continue
            else:
                print(f"   Erro ao parsear JSON (tentativa {tentativa + 1})")
                return ["Qualidade"]  # Categoria padrão
        
        except Exception as e:
            if tentativa < max_retries - 1:
                sleep(0.5)
                continue
            else:
                print(f"   Erro na API (tentativa {tentativa + 1}): {e}")
                return ["Qualidade"]  # Categoria padrão
    
    return ["Qualidade"]  # Fallback


def classificar_todos_comentarios(api_key, df):
    """Classifica todos os comentários usando OpenAI"""
    print(f"Iniciando classificacao com OpenAI ({MODELO_OPENAI})...")
    print(f"Categorias: {', '.join(CATEGORIAS)}\n")
    
    if LOG_DETALHADO:
        print("=" * 70)
        print("LOG DETALHADO - Cada comentario sera exibido")
        print("=" * 70)
    
    # Inicializar cliente OpenAI
    cliente = OpenAI(api_key=api_key)
    
    # Lista para armazenar categorias
    lista_categorias = []
    
    total = len(df)
    
    for idx, row in df.iterrows():
        comentario = row['Comentário']
        produto = row['Produto']
        
        # Log detalhado ou resumido
        if LOG_DETALHADO:
            # Truncar comentário para exibição (max 60 caracteres)
            comentario_preview = comentario[:60] + "..." if len(comentario) > 60 else comentario
            
            # Mostrar log de cada comentário
            print(f"[{idx + 1}/{total}] {produto} | {comentario_preview}")
        else:
            # Mostrar progresso a cada 50 comentários
            if (idx + 1) % 50 == 0 or idx == 0:
                print(f"Progresso: {idx + 1}/{total} comentarios processados...")
        
        # Classificar
        categorias = classificar_comentario_com_openai(cliente, comentario)
        
        # Converter lista para string separada por vírgula
        categorias_str = ", ".join(categorias)
        lista_categorias.append(categorias_str)
        
        # Mostrar resultado no log detalhado
        if LOG_DETALHADO:
            print(f"         → Categorias: {categorias_str}")
            print()
        
        # Pequeno delay para evitar rate limits (ajuste conforme necessário)
        if idx < total - 1 and (idx + 1) % 10 == 0:
            sleep(0.3)
    
    if LOG_DETALHADO:
        print("=" * 70)
    print(f"\nClassificacao concluida para {total} comentarios!\n")
    
    # Adicionar coluna de categorias ao DataFrame
    df['Categorias'] = lista_categorias
    
    return df


def salvar_resultados(df):
    """Salva os resultados de volta no Excel"""
    print("Salvando resultados no Excel...")
    
    try:
        # Carregar workbook
        wb = load_workbook(ARQUIVO_EXCEL)
        
        # Verificar se a sheet existe
        if SHEET_NAME not in wb.sheetnames:
            print(f"ERRO: Sheet '{SHEET_NAME}' nao encontrada")
            sys.exit(1)
        
        ws = wb[SHEET_NAME]
        
        # Encontrar a próxima coluna disponível ou a coluna "Categorias" se já existir
        ultima_coluna = ws.max_column
        coluna_categorias = None
        
        # Verificar se já existe coluna "Categorias"
        for col in range(1, ultima_coluna + 1):
            if ws.cell(row=1, column=col).value == "Categorias":
                coluna_categorias = col
                break
        
        # Se não existir, criar nova coluna
        if coluna_categorias is None:
            coluna_categorias = ultima_coluna + 1
            ws.cell(row=1, column=coluna_categorias, value="Categorias")
            ws.cell(row=1, column=coluna_categorias).font = openpyxl.styles.Font(bold=True)
        
        # Escrever categorias
        for idx, categorias in enumerate(df['Categorias'], start=2):  # start=2 para pular o cabeçalho
            ws.cell(row=idx, column=coluna_categorias, value=categorias)
        
        # Ajustar largura da coluna
        ws.column_dimensions[openpyxl.utils.get_column_letter(coluna_categorias)].width = 40
        
        # Salvar arquivo
        wb.save(ARQUIVO_EXCEL)
        print(f"Resultados salvos com sucesso em '{ARQUIVO_EXCEL}'!")
        print(f"Coluna 'Categorias' adicionada a sheet '{SHEET_NAME}'")
        
    except Exception as e:
        print(f"ERRO ao salvar Excel: {e}")
        sys.exit(1)


def mostrar_estatisticas(df):
    """Mostra estatísticas das classificações"""
    print("\n" + "=" * 70)
    print("ESTATISTICAS DAS CLASSIFICACOES")
    print("=" * 70 + "\n")
    
    # Contar ocorrências de cada categoria
    contagem_categorias = {cat: 0 for cat in CATEGORIAS}
    
    for categorias_str in df['Categorias']:
        categorias_lista = [cat.strip() for cat in categorias_str.split(',')]
        for cat in categorias_lista:
            if cat in contagem_categorias:
                contagem_categorias[cat] += 1
    
    total_comentarios = len(df)
    
    print("Distribuicao por categoria:\n")
    for categoria in CATEGORIAS:
        count = contagem_categorias[categoria]
        percentual = (count / total_comentarios) * 100
        print(f"  {categoria:15} : {count:4} comentarios ({percentual:.1f}%)")
    
    # Comentários com múltiplas categorias
    multiplas_categorias = sum(1 for cat_str in df['Categorias'] if ',' in cat_str)
    percentual_multiplas = (multiplas_categorias / total_comentarios) * 100
    
    print(f"\nComentarios com multiplas categorias: {multiplas_categorias} ({percentual_multiplas:.1f}%)")


def main():
    """Função principal"""
    # 1. Carregar configurações
    api_key = carregar_configuracao()
    
    # 2. Ler comentários do Excel
    df = ler_comentarios_excel()
    
    # 3. Classificar comentários com OpenAI
    df_classificado = classificar_todos_comentarios(api_key, df)
    
    # 4. Salvar resultados
    salvar_resultados(df_classificado)
    
    # 5. Mostrar estatísticas
    mostrar_estatisticas(df_classificado)
    
    print("\n" + "=" * 70)
    print("PROCESSO CONCLUIDO COM SUCESSO!")
    print("=" * 70)
    print(f"\nAbra o arquivo '{ARQUIVO_EXCEL}'")
    print(f"Veja a nova coluna 'Categorias' na sheet '{SHEET_NAME}'")


if __name__ == "__main__":
    main()

