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
            bool comecou = false; // Para controlar o início da extração após o primeiro título relevante
            bool isTextoVermelho = true; // Para ignorar textos em vermelho

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(caminhoArquivo, false))
            {
                var body = wordDoc.MainDocumentPart.Document.Body;

                foreach (var par in body.Elements<Paragraph>())
                {
                    string marcador = ObterMarcador(par); // Verifica se o parágrafo faz parte de uma lista
                    string textoParagrafo = string.Empty;

                    foreach (var run in par.Elements<Run>())
                    {
                        string texto = run.InnerText;
                        string textoAux = run.InnerText;

                        // Identifica início de uma nova seção
                        if (textoAux.StartsWith("1.") || textoAux.StartsWith("2.") || textoAux.StartsWith("3."))
                        {
                            isTextoVermelho = true;
                        }

                        // Ignora texto vermelho
                        var cor = run.RunProperties?.Color?.Val;
                        if (cor != null && cor.Equals("FF0000")) // Cor vermelha (Hex: #FF0000)
                        {
                            isTextoVermelho = false;
                            continue;
                        }

                        // Remove títulos e mantém a formatação correta
                        if (!comecou && (textoAux.StartsWith("1.") || textoAux.StartsWith("2.") || textoAux.StartsWith("3.")))
                        {
                            comecou = true;
                        }

                        if (comecou)
                        {
                            if (textoAux.StartsWith("1.") || textoAux.StartsWith("2.") || textoAux.StartsWith("3."))
                            {
                                texto = RemoverTitulo(texto);
                            }

                            textoParagrafo += texto + " ";
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(textoParagrafo))
                    {
                        // Adiciona o marcador de lista quando necessário
                        textoAtual += (!string.IsNullOrEmpty(marcador) ? marcador : "") + textoParagrafo.Trim() + @"\r\n";
                    }

                    // Se encontrou um novo título e há texto acumulado, adiciona à lista e reseta
                    if ((par.InnerText.StartsWith("1.") || par.InnerText.StartsWith("2.") || par.InnerText.StartsWith("3."))
                        && !string.IsNullOrEmpty(textoAtual.Trim()))
                    {
                        textos.Add(textoAtual.Trim());
                        textoAtual = string.Empty;
                    }
                }

                // Adiciona o último bloco de texto se houver
                if (!string.IsNullOrEmpty(textoAtual.Trim()))
                {
                    textos.Add(textoAtual.Trim());
                }
            }

            return textos.ToArray();
        }

        // Método para detectar marcadores de lista corretamente
        private string ObterMarcador(Paragraph par)
        {
            var numberingProperties = par.ParagraphProperties?.NumberingProperties;
            return numberingProperties != null ? "●    " : string.Empty;
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