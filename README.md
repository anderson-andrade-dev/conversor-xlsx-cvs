
# Conversão de Arquivos Excel para CSV

Este projeto é responsável por converter arquivos Excel (.xlsx) para arquivos CSV, com a possibilidade de excluir colunas específicas durante a conversão. Ele foi criado com Java, utilizando a biblioteca Apache POI para ler arquivos Excel e realizar a conversão. 

## Funcionalidades

1. **Leitura de arquivos Excel**:
   - O projeto lê arquivos `.xlsx` e converte seu conteúdo para o formato CSV.
   
2. **Exclusão de colunas**:
   - O usuário pode especificar quais colunas deseja excluir durante a conversão para CSV.

3. **Saída em CSV**:
   - O conteúdo do Excel é convertido para CSV com e sem exclusão de colunas.

4. **Criação de diretórios**:
   - Caso os diretórios de saída para os arquivos CSV não existam, eles são automaticamente criados.

## Dependências

O projeto utiliza as seguintes dependências:

- Apache POI (para manipulação de arquivos Excel)
- Log4j2 (para logs)
  
## Como usar

1. **Configuração do projeto**:
   - O projeto foi desenvolvido em Java e pode ser executado diretamente em um ambiente de desenvolvimento que suporte Java 11 ou superior.

2. **Passos para executar**:

   - **Importar o projeto** para sua IDE preferida (Eclipse, IntelliJ, etc.).
   - **Executar o método `main`** para realizar a conversão do arquivo Excel.
   - O arquivo Excel deve estar localizado no diretório `target/arquivoExcel` (ou outro de sua escolha).
   - O resultado da conversão será salvo em arquivos CSV dentro do diretório `target/arquivoCSV`.

3. **Entrada de dados**:
   - O arquivo Excel de entrada deve estar no formato `.xlsx`.
   - O código lê a primeira planilha do arquivo Excel e converte para CSV.

4. **Exclusão de colunas**:
   - Colunas podem ser excluídas da conversão utilizando um conjunto de strings (nomes das colunas a serem excluídas).
   
5. **Saída**:
   - O arquivo CSV será salvo dentro do diretório `target/arquivoCSV`.
   - Dois arquivos CSV são gerados: um com exclusão de colunas e outro sem.

## Exemplo de execução

Ao rodar o método `main`, o código realizará os seguintes passos:

1. Ler o arquivo Excel `testeAline.xlsx` presente em `target/arquivoExcel`.
2. Gerar um arquivo CSV com exclusão de colunas especificadas (por exemplo, "Bicicleta" e "Teste").
3. Gerar outro arquivo CSV com todos os dados, sem exclusões.
4. Os arquivos CSV serão salvos em `target/arquivoCSV`.

## Exemplo de saída

```
Arquivo CSV com exclusão gerado em: target/arquivoCSV/saida_com_exclusao.csv
Arquivo CSV completo gerado em: target/arquivoCSV/saida_completo.csv
```

## Log

Se você receber o seguinte erro ao rodar o projeto:

```
ERROR StatusLogger Log4j2 could not find a logging implementation. Please add log4j-core to the classpath. Using SimpleLogger to log to the console...
```

Basta adicionar o `log4j-core` nas dependências do seu projeto, conforme descrito abaixo.

## Dependência do Log4j2 (pom.xml)

```xml
<dependency>
    <groupId>org.apache.logging.log4j</groupId>
    <artifactId>log4j-core</artifactId>
    <version>2.17.1</version>
</dependency>
```

## Licença

Este projeto está sob a Licença MIT. Veja o arquivo LICENSE para mais detalhes.
