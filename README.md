# Gera-o-de-Planilhas-Excel-com-Dados-e-Totaliza-o
### How-To: Utilização do Script para Geração de Planilhas Excel com Dados e Totalização

Este guia explica como utilizar o script para gerar planilhas Excel a partir de um arquivo `.ods` com dados, incorporando funções de personalização, filtragem, totalização e interface gráfica.

---

#### Pré-requisitos
1. **Bibliotecas Necessárias**:
   - `pandas`
   - `tkinter`
   - `ttkbootstrap`
   - `openpyxl`
   - `logging`

   Essas bibliotecas podem ser instaladas via `pip`:
   ```bash
   pip install pandas ttkbootstrap openpyxl
   ```

2. **Arquivos de Configuração**:
   - **total.ods**: Arquivo de entrada contendo dados a serem processados.
   - **modelo.xlsx**: Modelo de planilha que serve como base para as novas abas geradas.
   - **area.txt**: Arquivo de texto contendo o nome da área que será usada para nomear o arquivo final.
   - **lista_nomes.txt**: Lista de nomes a serem filtrados no arquivo `.ods`.

3. **Estrutura do Diretório**:
   - Coloque todos os arquivos na mesma pasta que o script.

---

#### Funcionamento do Script
O script executa as seguintes tarefas principais:

1. **Interface Gráfica (GUI)**:
   - A interface gráfica criada com `ttkbootstrap` permite iniciar o processo de geração de planilhas Excel com um clique no botão **"Criar Abas no Excel"**.
   - Uma barra de progresso é exibida para acompanhar o progresso da execução.

2. **Funções Principais e Suas Descrições**:
   - **Configuração do Log**: Registra eventos no arquivo `log_script.log` para facilitar a análise de erros e o acompanhamento da execução.

3. **Carregamento e Validação dos Arquivos de Entrada**:
   - `verificar_arquivo_ods()`: Verifica se o arquivo `total.ods` está presente.
   - `carregar_modelo_excel()`: Carrega o modelo `modelo.xlsx`, que é usado para criar novas abas no Excel.
   - `carregar_area()`: Lê o conteúdo de `area.txt` e carrega o valor para nomear o arquivo de saída.
   - `carregar_lista_nomes()`: Carrega e processa a lista de nomes a partir de `lista_nomes.txt`, removendo linhas em branco.

4. **Processamento dos Dados**:
   - `identificar_duplicidade_nomes(lista_nomes)`: Identifica nomes duplicados na lista para evitar conflitos de nomes ao criar abas.
   - `obter_nome_completo(nome_lista, lista_nomes)`: Retorna o nome completo para ser inserido na célula `B2` de cada aba.

5. **Geração da Planilha**:
   - `criar_abas_excel(progresso)`: Função principal que executa as seguintes etapas:
     - **Carregamento do ODS e Filtro de Nomes**:
       - Carrega o `DataFrame` a partir do `total.ods` e filtra somente os nomes listados em `lista_nomes.txt`.
       - Remove `(Titular)` dos nomes completos.
     - **Criação de Abas e Dados**:
       - Cria uma aba para cada nome na lista, nomeando-a com o primeiro nome ou com uma combinação `PrimeiroNome-SegundoNome` em caso de duplicidade.
       - Adiciona o nome completo na célula `B2` e o valor da área na célula `B1`.
     - **Inserção dos Dados e Controle de Coluna**:
       - Para cada linha de dados, o script adiciona as informações de acordo com as colunas do arquivo `modelo.xlsx`.
       - Se o campo `"Andamento"` for `"Concluído"`, o valor de `"Alocação padrão"` é adicionado na coluna `H`.
       - Insere automaticamente o valor de **Total** nas colunas `F`, `G` e `H` no fim dos dados:
         - Se a aba está vazia, insere "Total" em `F11`, `G11`, e `H11` com valores `"0,00%"`.
         - Caso haja dados, calcula a soma das colunas `G` e `H` com uma fórmula (`=SUM(G9:G{ultima_linha-1})`).

6. **Salvar e Finalizar**:
   - A planilha final é salva com o nome especificado em `area.txt` na pasta `Dados`.

---

#### Passo a Passo para Utilização do Script

1. **Configuração Inicial**:
   - Coloque os arquivos `total.ods`, `modelo.xlsx`, `area.txt`, e `lista_nomes.txt` no mesmo diretório que o script.

2. **Executando o Script**:
   - Inicie o script através de um terminal ou editor Python.
   - Clique no botão **"Criar Abas no Excel"** para iniciar o processo.

3. **Saída**:
   - O arquivo final será salvo na pasta `Dados` com o nome especificado em `area.txt`.

4. **Verificação de Logs**:
   - Consulte o arquivo `log_script.log` para ver o registro detalhado das operações e identificar possíveis erros.

---

#### Exemplo de Resultados
Se `total.ods` contiver dados como:
| Nome Completo       | Andamento | Alocação padrão |
|---------------------|-----------|-----------------|
| João da Silva       | Concluído | 50%            |
| Maria dos Santos    | Em Progresso | 30%       |

A planilha gerada terá abas nomeadas conforme os nomes na lista, e exibirá a totalização conforme o formato descrito.

---

