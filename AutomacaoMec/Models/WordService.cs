using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace AutomacaoMec.Models
{
    public class WordService
    {
        public string[] ExtrairTextoDoWord(string caminhoArquivo)
        {
            var textos = new List<string>();
            string textoAtual = string.Empty;
            bool comecou = false;  // Variável para controlar o início da extração após o primeiro título relevante

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(caminhoArquivo, false))
            {
                var body = wordDoc.MainDocumentPart.Document.Body;

                var isTextoVermelho = true;
                // Percorrendo todos os elementos de texto (run) no documento
                foreach (var par in body.Elements<Paragraph>())
                {
                    foreach (var run in par.Elements<Run>())
                    {
                        var texto = run.InnerText;
                        var textoAux = run.InnerText;

                        if ((texto.StartsWith("1.") || texto.StartsWith("2.") || texto.StartsWith("3.")))
                        {
                            isTextoVermelho = true;
                        }

                        if (isTextoVermelho)
                        {
                            texto = "";
                        }
                        // Ignorar se o texto está em vermelho (cor vermelha)
                        var cor = run.RunProperties?.Color?.Val;
                        if (cor != null && cor.Equals("FF0000"))  // Cor vermelha (Hex: #FF0000)
                        {
                            isTextoVermelho = false;
                            continue;  // Ignora o texto vermelho

                        }

                        // Iniciar a coleta de texto quando encontrar o título "1.", "2." ou "3."
                        if (!comecou && (textoAux.StartsWith("1.") || textoAux.StartsWith("2.") || textoAux.StartsWith("3.")))
                        {
                            comecou = true;  // Começar a coleta após o primeiro título relevante
                        }

                        // Acumula o texto da seção atual, se já tiver começado a coleta
                        if (comecou)
                        {
                            // Remover o título da linha (por exemplo, "1.1")
                            if (textoAux.StartsWith("1.") || textoAux.StartsWith("2.") || textoAux.StartsWith("3."))
                            {
                                texto = RemoverTitulo(texto);  // Chama a função para remover o título da linha
                            }

                            textoAtual += texto + " ";
                        }
                    }

                    // Se encontrar uma nova seção (com título válido), armazene o texto atual (se houver)
                    if ((par.InnerText.StartsWith("1.") || par.InnerText.StartsWith("2.") || par.InnerText.StartsWith("3."))
                        && !string.IsNullOrEmpty(textoAtual.Trim()))
                    {
                        textos.Add(textoAtual.Trim());
                        textoAtual = string.Empty;  // Resetar para a próxima seção
                    }


                    if (!string.IsNullOrWhiteSpace(textoAtual) && !textoAtual.EndsWith("\r\n"))
                    {
                        textoAtual = textoAtual.Trim() + @"\r\n";
                    }
                }


                // Adiciona o último texto se houver
                if (!string.IsNullOrEmpty(textoAtual.Trim()))
                {
                    textos.Add(textoAtual.Trim());
                }
            }

            return textos.ToArray();
        }

        // Função para remover o título da linha (ex: "1.1 Políticas institucionais no âmbito do curso")
        private string RemoverTitulo(string texto)
        {
            // Encontrar o primeiro espaço após o número do título (ex: "1.1")
            var posEspaco = texto.IndexOf(" ");
            if (posEspaco > 0)
            {
                // Retornar a parte do texto após o primeiro espaço (removendo o título)
                return texto.Substring(posEspaco).Trim();
            }

            // Se não encontrar o espaço (o que é improvável), retorna o texto original
            return texto;
        }
        // Método para gerar o script de preenchimento dos campos do formulário
        public string GerarScriptSelenium(string[] textos)
        {
            var script = new StringBuilder();

            // Adicionar o script para preencher os campos

            var IsPergunta = true;

            var valorQuestao = 0;

            for (int i = 0; i < textos.Length; i++)
            {
                if (i == 58)
                {
                    break;
                }

                if (IsPergunta)
                {
                    i--;
                    IsPergunta = false;
                    continue;
                }

                var valorTbody = "tbody1";

                if (i <= 23)
                {
                    valorTbody = "tbody1";
                }
                else if (i < 40)
                {
                    if (i == 24)
                    {
                        valorQuestao = 0;
                    }
                    valorTbody = "tbody2";
                }
                else
                {
                    if (i == 40)
                    {
                        valorQuestao = 0;
                    }
                    valorTbody = "tbody3";
                }

                var texto = textos[i]; // Remover quebras de linha excessivas
                script.AppendLine($"var  tbody = document.getElementById('{valorTbody}');");
                script.AppendLine($"var tr{i} = tbody.querySelector('#Questao{valorQuestao}');");
                script.AppendLine($"var textarea{i} = tr{i}.querySelector('textarea');");
                script.AppendLine($"textarea{i}.value = '{texto}'");
                script.AppendLine();  // Linha em branco para separar as questões no script

                IsPergunta = true;
                valorQuestao++;
            }

            return script.ToString();
        }
    }
}