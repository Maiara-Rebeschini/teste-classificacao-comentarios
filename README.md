# 📊 Análise de Comentários de Produtos - Grupo NC

Sistema automatizado para análise de comentários de produtos usando OpenAI GPT-4-Turbo.

## 🚀 Configuração Inicial

### 1. Pré-requisitos
- Python 3.8 ou superior
- Conta OpenAI com API Key
- Windows PowerShell

### 2. Instalação

As dependências já foram instaladas no ambiente virtual. Se precisar reinstalar:

- Crie um ambiente virtual e instale as dependências.

```powershell
.\venv\Scripts\pip.exe install -r requirements.txt
```

### 3. Configurar API Key da OpenAI

**IMPORTANTE:** Edite o arquivo `.env` e adicione sua chave de API:

```env
OPENAI_API_KEY=sk-sua-chave-aqui
```

Para obter sua API Key:
1. Acesse https://platform.openai.com/api-keys
2. Crie uma nova chave
3. Copie e cole no arquivo `.env`

## 📖 Como Usar

### Script 1: Análise de Comentários (Resumo Positivos/Negativos)

```powershell
.\venv\Scripts\python.exe analise_comentarios.py
```
Ou clique em: `executar.bat`

**O que faz:**
1. **Lê** a sheet "Comentários de Produtos" do arquivo Excel
2. **Agrupa** todos os comentários por produto (Produto 01 até Produto 50)
3. **Analisa** cada produto usando OpenAI GPT-4-Turbo para identificar:
   - Resumo dos pontos positivos (máx 50 caracteres)
   - Resumo dos pontos negativos (máx 50 caracteres)
4. **Cria** uma nova sheet "Resumo de Análise" no mesmo arquivo Excel

### Script 2: Classificação de Comentários em Categorias

```powershell
.\venv\Scripts\python.exe classificar_comentarios.py
```
Ou clique em: `classificar.bat`

**O que faz:**
1. **Lê** cada comentário da sheet "Comentários de Produtos"
2. **Classifica** cada comentário em uma ou mais categorias:
   - Design (aparência, estética, beleza)
   - Qualidade (qualidade geral, acabamento, materiais)
   - Preço (valor, custo-benefício)
   - Durabilidade (resistência, durabilidade)
   - Logística (entrega, envio, embalagem)
3. **Adiciona** coluna "Categorias" na mesma sheet "Comentários de Produtos"

**Consulte:** `INSTRUCOES_CLASSIFICACAO.txt` para mais detalhes

### Estrutura do Arquivo Excel

**Sheet de Entrada:** "Comentários de Produtos"
- Coluna A: Produto (Produto 01, Produto 02, etc.)
- Coluna B: Comentário (texto das avaliações)

**Sheet de Saída:** "Resumo de Análise" (criada automaticamente)
- Coluna A: Produto
- Coluna B: Resumo dos comentários positivos
- Coluna C: Resumo dos comentários negativos

## 💰 Estimativa de Custos

### Script 1 - Análise de Comentários:
- **Modelo:** GPT-4-Turbo (alta qualidade)

### Script 2 - Classificação em Categorias:
- **Modelo:** GPT-5-nano (mais recente e econômico) ⭐

## 🔧 Estrutura do Projeto

```
Analise de Comentarios Grupo NC/
├── .env                                    # Configuração da API Key (VOCÊ DEVE EDITAR)
├── .gitignore                             # Arquivos ignorados pelo Git
├── requirements.txt                        # Dependências Python
├── venv/                                  # Ambiente virtual
│
├── analise_comentarios.py                 # Script 1: Análise/Resumo
├── executar.bat                           # Atalho para Script 1
│
├── classificar_comentarios.py             # Script 2: Classificação em Categorias
├── classificar.bat                        # Atalho para Script 2
│
├── README.md                              # Documentação completa
└── Teste Avaliação de Comentários.xlsx    # Arquivo Excel
```

## ⚠️ Troubleshooting

### Erro: "OPENAI_API_KEY não encontrada"
- Verifique se editou o arquivo `.env` com sua chave válida

### Erro: "Arquivo não encontrado"
- Certifique-se de que o arquivo "Teste Avaliação de Comentários.xlsx" está na pasta do projeto

### Erro: "Sheet não encontrada"
- Verifique se existe uma sheet chamada "Comentários de Produtos" no Excel

### PowerShell ExecutionPolicy Error
Se não conseguir ativar o venv, execute:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## 📝 Logs e Progresso

O script mostra em tempo real:
- Quantidade de comentários lidos
- Progresso da análise (Produto 15/50...)
- Resumos gerados para cada produto
- Status de conclusão
  
<img width="566" height="342" alt="image" src="https://github.com/user-attachments/assets/35fe959f-2727-46b6-91d8-178ea37222bc" />


## 🎯 Próximos Passos

1. ✅ Configure a API Key no arquivo `.env`
2. ✅ Execute o script
3. ✅ Abra o Excel e veja a nova sheet "Resumo de Análise"

## 📞 Suporte

Em caso de dúvidas ou problemas, verifique:
- Documentação da OpenAI: https://platform.openai.com/docs
- Documentação do openpyxl: https://openpyxl.readthedocs.io/


