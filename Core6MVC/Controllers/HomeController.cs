using Core6MVC.Models;
using Core6MVC.ViewModel;
using DocTool;
using DocTool.Dto;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace Core6MVC.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IWebHostEnvironment _webHostEnvironment;
        private Tool docTool { get; set; }
        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment webHostEnvironment)
        {
            _logger = logger;
            _webHostEnvironment = webHostEnvironment;
            this.docTool = new Tool(@"E:\PortableApps\LibreOfficePortable\App\libreoffice\program\soffice.exe", Path.Combine(webHostEnvironment.WebRootPath, "tempData"));
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
            var replaceData = new Dictionary<string, ReplaceDto>()
            {
                ["PONo"] = new ReplaceDto() { textStr = "11300000A1" },
                ["BarCode1_Image"] = new ReplaceDto()
                {
                    replaceType = ReplaceType.Image,
                    fileName = "barcode",
                    fileType = "jpg",
                    fileByteArr = new byte[0],
                    imageDpi = 72,
                    imageHeight = 30,
                    imageWidth = 150,
                },
                ["PoVenTable"] = new ReplaceDto()
                {
                    replaceType = ReplaceType.Table,
                    tableData = this.docTool.Word.ToXMLTable(new List<string>() { "品名", "數量", "總價" }, new List<List<string>>() { new List<string>() { "A", "1", "1500" }, new List<string>() { "B", "2", "4500" } }),
                },
                ["AppP_Table"] = new ReplaceDto()
                {
                    replaceType = ReplaceType.HtmlString,
                    htmlStr = "<p style=\"background-color:Tomato;\">Lorem ipsum...</p>",
                },
                ["AppP2_TableRow"] = new ReplaceDto()
                {
                    replaceType = ReplaceType.TableRow,
                    tableRowDatas = new List<List<string>>() { new List<string>() { "A", "1", "1500" }, new List<string>() { "B", "2", "4500" } }
             .Select(x =>
             {
                 var row = this.docTool.Word.CreateRow();
                 x.ForEach(item => row.Append(this.docTool.Word.CreateCell(item)));
                 return row;
             }).ToList()
                }
            };
            for (int i = 0; i < fileViewModelData.byteArrayDatas.Count; i++)
            {
                fileViewModelData.byteArrayDatas[i] = this.docTool.Word
                    .ReplaceTag("aaa.docx", fileViewModelData.byteArrayDatas[i], replaceData)
                    .GetData()
                    .fileByteArr;
            }
            return Json(new
            {
                fileViewModelData.byteArrayDatas
            });
        }
        [HttpPost]
        public IActionResult WordToPdf([FromBody] List<byte[]> byteArrayDatas)
        {
            for (int i = 0; i < byteArrayDatas.Count; i++)
            {
                byteArrayDatas[i] = this.docTool.Word
                    .ToPDF("aaa.docx", byteArrayDatas[i])
                    .GetData()
                    .fileByteArr;
            }
            return Json(new
            {
                byteArrayDatas,
            });
        }
    }
}