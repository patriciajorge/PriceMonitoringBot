# PriceMonitoringBot
Projeto de monitoramento de preço que acessa a página de um produto, verifica o preço atual, extrai e armazena o valor numérico em uma planilha Excel, junto com a data da consulta, nome do produto e link direto para o checkout. O bot roda a cada 30 minutos, atualizando a planilha com novas informações de preço.

## Funcionalidades

- Acessa a página de um produto.
- Verifica o preço atual.
- Extrai e armazena o valor numérico em uma planilha Excel.
- Armazena a data da consulta, o nome do produto e um link direto para o checkout.
- Executa automaticamente a cada 30 minutos.

## Requisitos

- Python 3.x
- Google Chrome

## Instalação

1. Instale as dependências:
    ```bash
    pip install -r requirements.txt
    ```

2. Execute o script:
    ```bash
    python app.py
    ```

3. O bot acessará a página do produto, verificará o preço e atualizará a planilha Excel (`motorcycle_price.xlsx`) a cada 30 minutos.
