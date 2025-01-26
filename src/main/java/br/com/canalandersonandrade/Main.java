package br.com.canalandersonandrade;

import java.io.*;
import java.util.*;

public class Main {

    public static void main(String[] args) {
        // Caminho do arquivo Excel
        // Este é o caminho onde o arquivo Excel que será processado está armazenado.
        // Certifique-se de que o arquivo exista neste caminho.
        String caminhoArquivo = "target/arquivoExcel/testeAline.xlsx";

        // Caminhos para arquivos CSV
        // Esses são os locais onde os arquivos CSV gerados serão salvos:
        // 1. Com exclusão de colunas específicas.
        // 2. Completo (sem exclusões).
        String caminhoCSVComExclusao = "target/arquivoCSV/saida_com_exclusao.csv";
        String caminhoCSVCompleto = "target/arquivoCSV/saida_completo.csv";

        // Criar diretório necessário
        // Antes de salvar os arquivos CSV, verificamos se o diretório "target/arquivoCSV" existe.
        // Se não existir, o programa tenta criá-lo.
        if (!DirectoryUtil.criarDiretorioSeNaoExistir("target/arquivoCSV")) {
            // Caso o diretório não possa ser criado, exibimos um erro e encerramos o programa.
            System.err.println("Erro ao criar diretório para arquivos CSV.");
            return;
        }

        // Verificar se o arquivo Excel existe
        // Antes de processar o arquivo, verificamos se ele realmente existe no caminho especificado.
        File arquivoExcel = new File(caminhoArquivo);
        if (!arquivoExcel.exists()) {
            // Se o arquivo não existir, exibimos uma mensagem de erro e encerramos o programa.
            System.err.println("Arquivo Excel não encontrado: " + caminhoArquivo);
            return;
        }

        // Colunas para exclusão
        // Este conjunto (Set) contém os nomes das colunas que queremos excluir ao gerar o CSV.
        // Por exemplo, neste caso, as colunas "Bicicleta" e "Teste" serão ignoradas no arquivo gerado.
        Set<String> colunasExcluir = new HashSet<>(Arrays.asList("Bicicleta", "Teste"));

        // Método com exclusão
        try (InputStream inputStream = new FileInputStream(arquivoExcel)) {
            // Aqui, abrimos o arquivo Excel e passamos o InputStream para o método de conversão.
            // O método "converterExcelParaCSVComExclusao" retorna o conteúdo do CSV como uma String.
            String csvComExclusao = ExcelReader.converterExcelParaCSVComExclusao(inputStream, colunasExcluir);

            // Após gerar o CSV, salvamos o conteúdo no local especificado.
            salvarCSV(csvComExclusao, caminhoCSVComExclusao);

            // Exibimos uma mensagem no console indicando que o arquivo foi gerado com sucesso.
            System.out.println("Arquivo CSV com exclusão gerado em: " + caminhoCSVComExclusao);
        } catch (Exception e) {
            // Se ocorrer algum erro (ex.: arquivo corrompido ou falta de permissão), exibimos o erro no console.
            System.err.println("Erro ao gerar CSV com exclusão: " + e.getMessage());
        }

        // Método sem exclusão
        try (InputStream inputStream = new FileInputStream(arquivoExcel)) {
            // Aqui, utilizamos outro método, "converterExcelParaCSV", que gera um CSV completo
            // (ou seja, sem excluir nenhuma coluna do arquivo Excel).
            String csvCompleto = ExcelReader.converterExcelParaCSV(inputStream);

            // Salvamos o CSV completo no local especificado.
            salvarCSV(csvCompleto, caminhoCSVCompleto);

            // Exibimos uma mensagem no console indicando que o arquivo foi gerado com sucesso.
            System.out.println("Arquivo CSV completo gerado em: " + caminhoCSVCompleto);
        } catch (Exception e) {
            // Se ocorrer algum erro (ex.: falha na leitura do arquivo), exibimos o erro no console.
            System.err.println("Erro ao gerar CSV completo: " + e.getMessage());
        }
    }

    /**
     * Método auxiliar para salvar o conteúdo de um CSV em um arquivo.
     *
     * @param conteudoCSV   O conteúdo do arquivo CSV em formato de String.
     * @param caminhoArquivo O caminho onde o arquivo CSV será salvo.
     * @throws IOException Se ocorrer um erro ao salvar o arquivo.
     */
    private static void salvarCSV(String conteudoCSV, String caminhoArquivo) throws IOException {
        // Criamos um BufferedWriter para escrever o conteúdo do CSV no arquivo especificado.
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(caminhoArquivo))) {
            // Escrevemos o conteúdo no arquivo.
            writer.write(conteudoCSV);
        }
        // O recurso será fechado automaticamente ao sair do bloco try (graças ao try-with-resources).
    }
}
