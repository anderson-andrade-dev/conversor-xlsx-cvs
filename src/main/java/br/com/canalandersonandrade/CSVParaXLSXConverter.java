package br.com.canalandersonandrade;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

/**
 * Esta classe é responsável por converter um arquivo CSV para um arquivo Excel no formato XLSX.
 *
 * A classe é final para garantir que não seja estendida, e o método principal `converterCSVParaXLSX` é estático,
 * o que permite seu uso sem a necessidade de instanciar a classe.
 *
 * As boas práticas seguidas incluem:
 * - Garantir que recursos como arquivos sejam fechados adequadamente utilizando `try-with-resources`.
 * - Tratar erros de forma clara, com mensagens de exceção específicas.
 * - A classe é imutável, o que ajuda a evitar modificações indesejadas no comportamento.
 *
 * O método `converterCSVParaXLSX` é responsável por:
 * 1. Ler o conteúdo de um arquivo CSV utilizando a biblioteca Apache Commons CSV.
 * 2. Criar um arquivo Excel (XLSX) com base nos dados lidos do CSV.
 * 3. Salvar o arquivo XLSX no caminho especificado.
 *
 * @author Anderson Andrade Dev
 * @date 26/01/2025
 */
public final class CSVParaXLSXConverter {

    /**
     * Converte um arquivo CSV para um arquivo Excel no formato XLSX.
     *
     * Este método:
     * 1. Lê um arquivo CSV utilizando a biblioteca Apache Commons CSV.
     * 2. Cria um arquivo Excel utilizando o Apache POI.
     * 3. Preenche o arquivo Excel com os dados do CSV, incluindo o cabeçalho e os dados das linhas.
     * 4. Salva o arquivo XLSX no caminho especificado.
     *
     * A leitura do arquivo CSV é feita com a configuração para ignorar o cabeçalho e remover espaços extras.
     * Cada linha do arquivo CSV é convertida em uma linha do arquivo Excel.
     *
     * Exceções: Se ocorrer algum erro ao ler o arquivo CSV ou ao salvar o arquivo XLSX, uma exceção do tipo
     * `IOException` será lançada, com uma mensagem detalhada sobre a falha.
     *
     * @param caminhoCSV Caminho do arquivo CSV a ser lido.
     * @param caminhoXLSX Caminho do arquivo XLSX de saída.
     * @throws IOException Caso haja erro ao ler o CSV ou salvar o XLSX.
     */
    public static void converterCSVParaXLSX(String caminhoCSV, String caminhoXLSX) throws IOException {
        // Garantir que o arquivo CSV exista e seja acessível
        try (BufferedReader reader = Files.newBufferedReader(Paths.get(caminhoCSV));
             // Usando Apache Commons CSV para ler o arquivo CSV
             CSVParser csvParser = new CSVParser(reader, CSVFormat.DEFAULT.withHeader().withIgnoreHeaderCase().withTrim())) {

            // Criar o workbook para o Excel, no formato XLSX
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Sheet1");

            // Obter todos os registros (linhas) do arquivo CSV
            List<CSVRecord> records = csvParser.getRecords();
            int rowNum = 0;

            // Criar a primeira linha como cabeçalho (nome das colunas)
            Row headerRow = sheet.createRow(rowNum++);
            for (int i = 0; i < records.get(0).size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(records.get(0).get(i));  // Definir o valor da célula com o nome da coluna
            }

            // Preencher as linhas subsequentes com os dados do CSV
            for (CSVRecord record : records) {
                Row row = sheet.createRow(rowNum++);
                for (int i = 0; i < record.size(); i++) {
                    row.createCell(i).setCellValue(record.get(i));  // Definir o valor da célula com os dados
                }
            }

            // Salvar o arquivo Excel gerado
            try (FileOutputStream fileOut = new FileOutputStream(caminhoXLSX)) {
                workbook.write(fileOut);  // Gravar os dados no arquivo XLSX
            }

            // Log de sucesso
            System.out.println("Arquivo XLSX gerado com sucesso em: " + caminhoXLSX);

        } catch (IOException e) {
            // Captura e lança exceções específicas, com mensagens claras de erro
            throw new IOException("Erro ao processar o arquivo CSV ou salvar o arquivo Excel: " + e.getMessage(), e);
        }
    }
}
