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
            bool comecou = false;  // Variável para controlar o início da extração (após 1.1)

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(caminhoArquivo, false))
            {
                var body = wordDoc.MainDocumentPart.Document.Body;

                // Percorrendo todos os elementos de texto (run) no documento
                foreach (var par in body.Elements<Paragraph>())
                {
                    foreach (var run in par.Elements<Run>())
                    {
                        // Extraímos o texto sem formatação (ignorando a cor vermelha ou qualquer estilo)
                        var texto = run.InnerText;

                        // Ignorar se o texto está em vermelho (cor vermelha)
                        var cor = run.RunProperties?.Color?.Val;
                        if (cor != null && cor.Equals("FF0000"))  // Cor vermelha (Hex: #FF0000)
                        {
                            continue;  // Ignora o texto vermelho
                        }

                        // Se o texto for da seção 1.1 em diante, começa a armazenar
                        if (!comecou && texto.Contains("1.1"))
                        {
                            comecou = true;
                        }

                        if (comecou)
                        {
                            textoAtual += texto + " ";
                        }
                    }

                    // Se o texto da seção atual for diferente de vazio, adicionar à lista de textos
                    if (!string.IsNullOrEmpty(textoAtual.Trim()))
                    {
                        textos.Add(textoAtual.Trim());
                        textoAtual = string.Empty;  // Resetar para a próxima seção
                    }
                }
            }

            return textos.ToArray();
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

                var tbody1 = 23;
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


                var texto = textos[i].Replace("\r\n", " ").Replace("\n", " "); // Remover quebras de linha excessivas
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

