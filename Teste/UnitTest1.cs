using AutomacaoMec.Models;

namespace Teste
{
    public class UnitTest1
    {
        public class WordServiceTests
        {
            [Fact]
            public void Deve_Extrair_Texto_Do_Documento_Word()
            {
                // Arrange
                string caminhoArquivo = @"C:\Doc\Documento.docx"; // Certifique-se de que o arquivo existe
                WordService wordService = new WordService();

                // Act
                List<string> textos = wordService.ExtrairTextoDoWord(caminhoArquivo);

                // Assert
                Assert.NotEmpty(textos); // Verifica se há texto extraído
                Assert.Contains("Texto esperado", textos); // Verifica se um texto específico foi encontrado
            }
        }
    }
}
