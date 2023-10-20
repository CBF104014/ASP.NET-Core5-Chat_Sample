using Core6MVC.Models;
using Core6MVC.ViewModel;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Tool.Doc;

namespace Core6MVC.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IWebHostEnvironment _webHostEnvironment;
        private WordTool _wordTool { get; set; }
        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment webHostEnvironment)
        {
            _logger = logger;
            _webHostEnvironment = webHostEnvironment;
            _wordTool = new WordTool(webHostEnvironment.WebRootPath);
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Document()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
        /// <summary>
        /// 附件下載
        /// </summary>
        [HttpPost]
        public IActionResult GetFileSample([FromBody] string fileCode)
        {
            var sampleFileFullName = "";
            switch (fileCode)
            {
                case "F1001":
                    sampleFileFullName = "Replace_hightlight測試.docx";
                    break;
                default:
                    throw new Exception("fileCode not found!");
            }
            var byteArrayData = System.IO.File.ReadAllBytes(Path.Combine(_webHostEnvironment.WebRootPath, "sampleFile", sampleFileFullName));
            return Json(new
            {
                sampleFileName = Path.GetFileNameWithoutExtension(sampleFileFullName),
                sampleFileType = Path.GetExtension(sampleFileFullName).Replace(".", ""),
                byteArrayData,
            });
        }
        /// <summary>
        /// Word-替換關鍵字
        /// </summary>
        [HttpPost]
        public IActionResult WordReplaceTag([FromBody] fileViewModel fileViewModelData)
        {
            for (int i = 0; i < fileViewModelData.byteArrayDatas.Count; i++)
            {
                fileViewModelData.byteArrayDatas[i] = _wordTool
                    .LoadFile(fileViewModelData.byteArrayDatas[i])
                    .FindAndReplace(fileViewModelData.keyValueDict);
            }
            return Json(new
            {
                fileViewModelData.byteArrayDatas,
            });
        }
        [HttpPost]
        public IActionResult WordToPdf([FromBody] List<byte[]> byteArrayDatas)
        {
            for (int i = 0; i < byteArrayDatas.Count; i++)
            {
                byteArrayDatas[i] = _wordTool
                    .LoadFile(byteArrayDatas[i])
                    .ToPDF();
            }
            return Json(new
            {
                byteArrayDatas,
            });
        }
    }
}