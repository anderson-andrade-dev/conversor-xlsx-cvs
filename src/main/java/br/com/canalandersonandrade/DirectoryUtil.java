package br.com.canalandersonandrade;

import java.io.File;

/**
 * Classe utilitária para manipulação de diretórios.
 * Esta classe garante que um diretório exista, criando-o se necessário.
 *
 * @author Anderson Andrade Dev
 * @date 26/01/2025
 */
public class DirectoryUtil {

    /**
     * Verifica se o diretório existe e, caso não, o cria.
     *
     * @param caminho O caminho do diretório a ser verificado/criado.
     * @return true se o diretório foi criado ou já existia; false caso contrário.
     */
    public static boolean criarDiretorioSeNaoExistir(String caminho) {
        File diretório = new File(caminho);

        // Se o diretório não existir, tenta criá-lo
        if (!diretório.exists()) {
            return diretório.mkdirs(); // Retorna true se o diretório foi criado
        }

        // Retorna true se o diretório já existia
        return true;
    }
}