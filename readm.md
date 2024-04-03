# Projeto de Análise Financeira

Este projeto visa fornecer uma análise detalhada das transações financeiras, categorizando-as em entradas e saídas, e apresentando um resumo financeiro para o período analisado.

## Descrição

O script em Python lê dados de transações financeiras de um arquivo CSV, organiza-os em entradas e saídas, e gera uma planilha Excel com os dados organizados e resumidos. As entradas são transações com valores positivos, enquanto as saídas são transações com valores negativos. O resumo inclui o somatório das entradas, o somatório das saídas e o saldo final.

## Recursos

- **Leitura de Dados**: Lê transações financeiras de um arquivo CSV.
- **Organização de Dados**: Separa as transações em entradas e saídas.
- **Geração de Planilha Excel**: Cria uma planilha Excel com os dados organizados, aplicando formatações específicas para datas e valores monetários.
- **Resumo Financeiro**: Calcula o somatório das entradas, das saídas e o saldo final, apresentando-os na planilha Excel.

## Dependências

Para executar este projeto, você precisará instalar as seguintes bibliotecas Python:

- pandas
- openpyxl

Você pode instalar essas dependências executando:

```
pip install pandas openpyxl
```

Ou usando o arquivo `requirements.txt` incluído neste projeto:


```
pip install -r requirements.txt
```
## Uso

Para utilizar este script, siga os passos abaixo:

1. Certifique-se de que o arquivo CSV com os dados das transações esteja no formato correto e acessível.
2. Atualize o caminho do arquivo CSV no script para corresponder à localização do seu arquivo de dados.
3. Execute o script para gerar a planilha Excel com os dados organizados e o resumo financeiro.

## Contribuição

Contribuições para este projeto são bem-vindas. Sinta-se à vontade para forkar o repositório, fazer suas alterações e abrir um pull request.

## Licença

Este projeto está sob a licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.
