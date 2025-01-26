package br.com.canalandersonandrade;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.File;

/**
 * @author Anderson Andrade Dev
 * @Data de Criação 26/01/2025
 */
public class DirectoryUtilTest {
    @Test
    public void testCriarDiretorioSeNaoExistir_Sucesso() {
        // Caminho de teste
        String caminho = "target/testDir";

        // Remover o diretório de teste se já existir
        File dir = new File(caminho);
        if (dir.exists()) {
            dir.delete();
        }

        // Testa a criação do diretório
        assertTrue(DirectoryUtil.criarDiretorioSeNaoExistir(caminho));
    }

}
