# Guia Completo de Execução - Ferramenta de Validação de Dados

## 📋 Pré-requisitos

### 1. Python
- Python 3.8 ou superior instalado
- Verificar instalação: `python --version` ou `python3 --version`

### 2. Dependências
Instalar as bibliotecas necessárias:
```bash
pip install pandas openpyxl numpy
```

Ou usando o arquivo requirements.txt:
```bash
pip install -r requirements.txt
```

## 🚀 Passo a Passo para Executar

### Passo 1: Preparar os Arquivos de Dados

#### 1.1 Arquivo Excel (dados_producao.xlsx)
- **Localização**: `data/dados_producao.xlsx`
- **Requisitos**:
  - Deve ter uma aba chamada "Producao"
  - Deve conter a coluna chave "ID_CALCULO"
  - Deve ter as colunas que serão validadas: SUBTOTAL, VALOR_DESCONTO, VALOR_LIQUIDO, etc.

⚠️ **IMPORTANTE**: Se o Excel contém fórmulas mas os valores aparecem como vazios:
1. Abra o arquivo no Excel
2. Pressione `Ctrl + Alt + F9` para recalcular todas as fórmulas
3. Salve o arquivo
4. Faça uma cópia do arquivo original (com fórmulas) como `dados_producao_original.xlsx`

#### 1.2 Arquivo CSV (dados_sql_simulados.csv)
- **Localização**: `data/dados_sql_simulados.csv`
- **Formato esperado**:
  - Separador: `;` (ponto e vírgula)
  - Decimal: `.` (ponto)
  - Encoding: UTF-8
  - Deve conter a mesma coluna chave "ID_CALCULO"

Exemplo de conteúdo CSV:
```csv
ID_CALCULO;PRODUTO;CATEGORIA;SUBTOTAL;VALOR_DESCONTO;VALOR_LIQUIDO;VALOR_IMPOSTO;VALOR_TOTAL;STATUS_CALCULO
CALC_0001;Produto A;ELETRONICOS;4537.6;907.52;3630.08;181.5;3811.58;CALCULADO
CALC_0002;Produto A;LIVROS;4693.5;469.35;4224.15;211.21;4435.36;CALCULADO
```

### Passo 2: Configurar o Arquivo JSON

Criar ou editar o arquivo `config_exemplo_final.json`:

```json
{
  "arquivo_fonte_1": {
    "caminho": "./data/dados_producao.xlsx",
    "aba_planilha": "Producao",
    "coluna_chave": "ID_CALCULO"
  },
  "arquivo_fonte_2": {
    "caminho": "./data/dados_sql_simulados.csv",
    "aba_planilha": null,
    "coluna_chave": "ID_CALCULO"
  },
  "arquivo_formulas": {
    "caminho": "./data/dados_producao_original.xlsx",
    "tipo": "auto_discovery",
    "linha_amostra": 2,
    "key_column": "ID_CALCULO"
  },
  "colunas_para_validar": [
    "SUBTOTAL",
    "VALOR_DESCONTO", 
    "VALOR_LIQUIDO",
    "VALOR_IMPOSTO",
    "VALOR_TOTAL",
    "STATUS_CALCULO"
  ],
  "regras_de_tolerancia": {
    "SUBTOTAL": {
      "tipo": "decimal",
      "casas_decimais": 2
    },
    "VALOR_DESCONTO": {
      "tipo": "decimal", 
      "casas_decimais": 2
    },
    "VALOR_LIQUIDO": {
      "tipo": "decimal",
      "casas_decimais": 2
    },
    "VALOR_IMPOSTO": {
      "tipo": "decimal",
      "casas_decimais": 2
    },
    "VALOR_TOTAL": {
      "tipo": "decimal",
      "casas_decimais": 2
    },
    "default": {
      "tipo": "exata"
    }
  },
  "arquivo_saida": {
    "caminho": "./output/relatorio_validacao.xlsx"
  }
}
```

### Passo 3: Criar Estrutura de Pastas

Certifique-se de que as seguintes pastas existem:
```bash
mkdir -p data output
```

### Passo 4: Executar a Validação

Execute o comando principal:
```bash
python validate.py config_exemplo_final.json
```

Ou com modo verbose (mais detalhes):
```bash
python validate.py config_exemplo_final.json -v
```

### Passo 5: Verificar o Resultado

O relatório será gerado em: `output/relatorio_validacao.xlsx`

## 📊 Entendendo o Relatório de Saída

O arquivo Excel gerado contém 8 abas:

### 1. **Resumo Executivo**
- Visão geral da validação
- Taxa de sucesso
- Principais problemas encontrados
- Recomendações

### 2. **Sumário**
- Total de comparações
- Total de concordâncias/divergências
- Taxa de concordância
- Divergências por coluna

### 3. **Detalhes da Validação**
Contém todas as comparações com:
- **Chave**: ID do registro
- **Valor_Arquivo_1**: Valor do Excel
- **Valor_Arquivo_2**: Valor do CSV/SQL
- **Resultado_Validacao**: VERDADEIRO ou FALSO
- **Formula_Original**: Fórmula do Excel (ex: =E2*F2)
- **Formula_Traduzida**: Fórmula com nomes de colunas (ex: =QUANTIDADE*VALOR_UNITARIO)
- **Divergencias_Na_Formula**: Possíveis causas das divergências

### 4. **Linhas sem Correspondência**
- Registros que existem apenas em um dos arquivos

### 5. **Fórmulas Extraídas**
- Lista completa de todas as fórmulas encontradas
- Mapeamento de células para colunas

### 6. **Mapa de Dependências**
- Visualização de quais colunas dependem de outras
- Análise de impacto

### 7. **Alertas e Recomendações**
- Alertas de alta, média e baixa severidade
- Recomendações específicas

### 8. **Impacto em Cascata**
- Análise de como erros se propagam
- Colunas afetadas por divergências

## 🔧 Solução de Problemas Comuns

### Problema 1: "Valor_Arquivo_1" aparece vazio/nulo
**Causa**: Excel com fórmulas não calculadas
**Solução**: 
1. Abrir o Excel e recalcular (Ctrl+Alt+F9)
2. Salvar o arquivo
3. Executar novamente

### Problema 2: Fórmulas não aparecem no relatório
**Causa**: Arquivo de fórmulas incorreto
**Solução**: 
1. Verificar se `arquivo_formulas.caminho` aponta para arquivo com fórmulas
2. Confirmar que `tipo: "auto_discovery"` está configurado

### Problema 3: Erro ao ler CSV
**Causa**: Formato incorreto do CSV
**Solução**: 
1. Verificar separador (deve ser `;`)
2. Verificar decimal (deve ser `.`)
3. Salvar como UTF-8

### Problema 4: Colunas não encontradas
**Causa**: Nomes de colunas diferentes
**Solução**: 
1. Verificar exatamente os nomes das colunas em ambos arquivos
2. Ajustar `colunas_para_validar` no config.json

## 📁 Estrutura de Arquivos Necessária

```
validator_tool/
├── validate.py              # Script principal
├── config_exemplo_final.json # Configuração
├── requirements.txt         # Dependências
├── src/                     # Código fonte
│   ├── config_loader.py
│   ├── data_loader.py
│   ├── data_aligner.py
│   ├── validator.py
│   ├── formula_extractor.py
│   ├── formula_auto_discovery.py
│   ├── impact_analyzer.py
│   ├── cross_sheet_analyzer.py
│   └── report_generator.py
├── data/                    # Seus arquivos de dados
│   ├── dados_producao.xlsx
│   ├── dados_producao_original.xlsx
│   └── dados_sql_simulados.csv
└── output/                  # Relatórios gerados
    └── relatorio_validacao.xlsx
```

## 💡 Dicas Importantes

1. **Sempre mantenha backup** dos arquivos originais
2. **Teste com poucos dados primeiro** para verificar se está funcionando
3. **Use modo verbose (-v)** para debug quando houver problemas
4. **Verifique o log** em `validator.log` para mais detalhes
5. **Garanta que as colunas chave** existem e têm valores únicos

## 🎯 Exemplo Completo

```bash
# 1. Clonar ou baixar o projeto
cd validator_tool

# 2. Instalar dependências
pip install -r requirements.txt

# 3. Preparar seus arquivos em data/
# - dados_producao.xlsx (com valores calculados)
# - dados_producao_original.xlsx (com fórmulas)
# - dados_sql_simulados.csv

# 4. Executar validação
python validate.py config_exemplo_final.json

# 5. Verificar resultado
# Abrir output/relatorio_validacao.xlsx no Excel
```

## ❓ Precisa de Ajuda?

1. Verifique o arquivo `validator.log` para mensagens de erro detalhadas
2. Execute com `-v` para mais informações: `python validate.py config.json -v`
3. Confirme que todos os arquivos estão nos locais corretos
4. Verifique se os nomes das colunas estão exatamente iguais à configuração