# Inserir Notas Fiscais No Thuban

Este projeto é um script de automação para criação e contabilização de Notas Fiscais (NFs) utilizando o navegador Microsoft Edge e automação de interação com a interface do usuário através das bibliotecas `Selenium` e `PyAutoGUI`. O script também utiliza `Pandas` para leitura de arquivos Excel e automação de preenchimento de dados em um sistema web.

## Funcionalidades

- **Criação de Notas Fiscais**: Automatiza o preenchimento de formulários no sistema Thuban Web para criação de NFs, preenchendo campos com dados extraídos de uma planilha Excel.
- **Contabilização de Notas Fiscais**: Automatiza o processo de contabilização de NFs, enviando as informações para o sistema Thuban Web.
- **Automação com Selenium e PyAutoGUI**: Combina a automação de ações do navegador (com Selenium) e simulação de teclado e mouse (com PyAutoGUI) para preencher formulários web.
- **Integração com Excel**: Usa `pandas` para ler e processar dados a partir de planilhas Excel (.xlsx e .xlsm).
- **Geração de arquivos PDF**: Exporta as NFs criadas em formato PDF e salva em um diretório específico.

## Pré-requisitos

Antes de executar o script, certifique-se de que os seguintes requisitos estão atendidos:

1. **Navegador Edge**: O navegador Microsoft Edge deve estar instalado.
2. **WebDriver do Edge**: O WebDriver do Edge deve ser configurado corretamente para depuração remota.
3. **Python 3.8 ou superior**: O script foi desenvolvido usando Python 3.8.
4. **Pacotes Python necessários**:
   - `selenium`
   - `pandas`
   - `pyautogui`
   - `pyperclip`
   - `tkinter`

# Configuração

## Configuração do WebDriver
Certifique-se de que o WebDriver do Edge está configurado para depuração remota. Use o seguinte comando para iniciar o WebDriver:

```bash
& "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222 --user-data-dir="C:\edge_selenium"
```
# Estrutura da Planilha
O script espera que a planilha Excel tenha os dados nas seguintes colunas:

- **Coluna 2:** Tipo de Nota Fiscal
- **Coluna 3:** Número da Nota Fiscal
- **Coluna 4:** CNPJ
- **Coluna 6:** Inscrição Municipal
- **Coluna 7:** Data de Emissão
- **Coluna 9:** Data de Vencimento
- **Coluna 10:** Descrição da NF
- **Coluna 11:** Valor da NF
- **Coluna 12:** Código do Boleto
- **Coluna 15:** Tipo de Pagamento (Pix, Boleto, ou Dados Bancários)
- **Colunas 17-20:** Dados bancários (para tipo de pagamento "Dados Bancários")

# Como Usar

## Passo 1: Iniciar o WebDriver
Execute o WebDriver do Edge com o comando listado acima para iniciar a sessão de depuração remota.

## Passo 2: Executar o Script
Após configurar o WebDriver, execute o script para iniciar o processo de automação.

```bash
python script_nome.py
```
## Passo 3: Seleção de Arquivo Excel
Ao rodar o script, uma janela do sistema será aberta para que você selecione a planilha Excel contendo os dados das NFs.

## Passo 4: Acompanhamento do Processo
O script irá alternar automaticamente entre o navegador e as ações de automação para criar ou contabilizar as notas fiscais no sistema Thuban Web.

## Passo 5: Finalização
Após o término do processo, uma caixa de diálogo será exibida com o status do processo.

