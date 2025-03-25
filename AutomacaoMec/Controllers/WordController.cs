using System.Diagnostics;
using AutomacaoMec.Models;
using Microsoft.AspNetCore.Mvc;

namespace AutomacaoMec.Controllers
{
    public class WordController : Controller
    {
        private readonly WordService _wordService;

        public WordController()
        {
            _wordService = new WordService();
        }

        // Tela inicial
        public IActionResult Index()
        {
            return View();
        }

        // Método para processar o arquivo enviado
        [HttpPost]
        public IActionResult GerarScript(IFormFile arquivo)
        {
            if (arquivo == null || arquivo.Length == 0)
            {
                ModelState.AddModelError("Arquivo", "Por favor, selecione um arquivo.");
                return View("Index");
            }

            // Caminho para salvar o arquivo temporariamente
            var uploadPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");

            // Verificar se o diretório existe e criar caso não exista
            if (!Directory.Exists(uploadPath))
            {
                Directory.CreateDirectory(uploadPath);
            }

            var caminhoArquivo = Path.Combine(uploadPath, arquivo.FileName);

            // Salvar o arquivo no servidor
            using (var stream = new FileStream(caminhoArquivo, FileMode.Create))
            {
                arquivo.CopyTo(stream);
            }

            // Extrair texto do arquivo e gerar script
            var textos = _wordService.ExtrairTextoDoWord(caminhoArquivo);
            var script = _wordService.GerarScriptSelenium(textos);

            // Exibir o script na view
            ViewBag.Script = script;

            // Apagar o arquivo depois de usar
            System.IO.File.Delete(caminhoArquivo);

            return View("Resultado");
        }
    }
}