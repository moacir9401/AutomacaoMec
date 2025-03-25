using System.Diagnostics;
using AutomacaoMec.Models;
using Microsoft.AspNetCore.Mvc;

namespace AutomacaoMec.Controllers
{
    public class HomeController : Controller
    {
        private readonly WordService _wordService;

        public HomeController()
        {
            _wordService = new WordService();
        }

        public IActionResult Index()
        {

            return View();
        }

        [HttpPost]
        public IActionResult GerarScript(IFormFile arquivo)
        {
            if (arquivo == null || arquivo.Length == 0)
            {
                ModelState.AddModelError("Arquivo", "Por favor, selecione um arquivo.");
                return View("Index");
            }

            // Caminho para salvar o arquivo temporariamente
            var caminhoArquivo = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads", arquivo.FileName);

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

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
