package br.com.canalandersonandrade;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

/**
 * Classe para leitura de arquivos Excel e conversão para CSV.
 * <p>
 * Esta classe oferece funcionalidades para ler arquivos Excel e converter o conteúdo
 * em formato CSV. A conversão pode ser feita com ou sem exclusão de colunas específicas.
 * </p>
 *
 * @author Anderson Andrade Dev
 * @date 26/01/2025
 */
public final class ExcelReader {

    /**
     * Converte um arquivo Excel em CSV, ignorando as colunas especificadas.
     *
     * @param inputStream    O InputStream do arquivo Excel a ser lido.
     * @param colunasExcluir As colunas a serem excluídas do CSV. Pode ser {@code null} para não excluir colunas.
     * @return O conteúdo do arquivo Excel em formato CSV.
     * @throws IOException Se ocorrer um erro ao ler o arquivo Excel.
     */
    public static String converterExcelParaCSVComExclusao(InputStream inputStream, Set<String> colunasExcluir) throws IOException {
        if (inputStream == null) {
            throw new IllegalArgumentException("O InputStream não pode ser nulo.");
        }

        // Tratar colunasExcluir como um conjunto vazio se for nulo
        if (colunasExcluir == null) {
            colunasExcluir = Collections.emptySet();
        }

        StringBuilder csvBuilder = new StringBuilder();

        // Abrir o arquivo Excel
        try (Workbook workbook = new XSSFWorkbook(inputStream)) {

            // Obter a primeira planilha
            Sheet sheet = workbook.getSheetAt(0);

            // Iterar sobre as linhas da planilha
            Iterator<Row> rowIterator = sheet.iterator();

            // Extrair cabeçalhos (primeira linha da planilha) como chaves
            Map<Integer, String> cabeçalhos = new HashMap<>();
            if (rowIterator.hasNext()) {
                Row headerRow = rowIterator.next();
                for (Cell cell : headerRow) {
                    cabeçalhos.put(cell.getColumnIndex(), cell.getStringCellValue());
                }
            }

            // Criar a primeira linha de CSV (cabeçalhos)
            for (String chave : cabeçalhos.values()) {
                if (!colunasExcluir.contains(chave)) {
                    csvBuilder.append(chave).append(",");
                }
            }
            if (csvBuilder.length() > 0) {
                csvBuilder.deleteCharAt(csvBuilder.length() - 1); // Remover a última vírgula
                csvBuilder.append("\n");
            }

            // Iterar pelas outras linhas (dados) da planilha
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                for (Cell cell : row) {
                    String chave = cabeçalhos.get(cell.getColumnIndex());
                    if (!colunasExcluir.contains(chave)) {
                        csvBuilder.append(obterValorDaCelula(cell)).append(",");
                    }
                }
                if (csvBuilder.length() > 0 && csvBuilder.charAt(csvBuilder.length() - 1) == ',') {
                    csvBuilder.deleteCharAt(csvBuilder.length() - 1); // Remover a última vírgula
                }
                csvBuilder.append("\n");
            }
        } catch (IOException e) {
            throw new IOException("Erro ao processar o arquivo Excel: " + e.getMessage(), e);
        }

        return csvBuilder.toString();
    }

    /**
     * Converte um arquivo Excel em CSV sem excluir colunas.
     *
     * @param inputStream O InputStream do arquivo Excel a ser lido.
     * @return O conteúdo do arquivo Excel em formato CSV.
     * @throws IOException Se ocorrer um erro ao ler o arquivo Excel.
     */
    public static String converterExcelParaCSV(InputStream inputStream) throws IOException {
        return converterExcelParaCSVComExclusao(inputStream, null);
    }

    /**
     * Obtém o valor da célula em um formato apropriado.
     *
     * @param cell A célula a ser processada.
     * @return O valor da célula como uma String.
     */
    private static String obterValorDaCelula(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString(); // Retorna data como string
                } else {
                    return String.valueOf(cell.getNumericCellValue()); // Retorna número
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
}
