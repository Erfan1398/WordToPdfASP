using WordToPdf.Models;
using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;
namespace WordToPdf.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        private Microsoft.AspNetCore.Hosting.IHostingEnvironment _environment;
        public HomeController(ILogger<HomeController> logger, Microsoft.AspNetCore.Hosting.IHostingEnvironment environment)
        {
            _logger = logger;
            _environment = environment;
        }

        public IActionResult Index()
        {
            //for test only
           // string output = Path.Combine(Path.Combine(_environment.WebRootPath, "files"), "3") + ".pdf";
           // string source = Path.Combine(Path.Combine(_environment.WebRootPath, "files"), "3.docx");
           // ConvertWordToSpecifiedFormat(source, output, Word.WdSaveFormat.wdFormatPDF);
            return View();
        }
        
        //[HttpPost] // for old .net
        //public FileResult Convert(HttpPostedFileBase postedFile)
        //{
        //    string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(postedFile.FileName);
        //    string filePath = Server.MapPath("~/Files/") + Path.GetFileName(postedFile.FileName);
        //    postedFile.SaveAs(filePath);
        //    string input = filePath;
        //    string output = Server.MapPath("~/Files/") + fileNameWithoutExtension + ".pdf";
        //    ConvertWordToSpecifiedFormat(input, output, Word.WdSaveFormat.wdFormatPDF);
        //    return File(output, "application/pdf", fileNameWithoutExtension + ".pdf");
        //}
        
        [HttpPost]
        public IActionResult uploader()
        {
            var filelist = HttpContext.Request.Form.Files;
            if (filelist.Count > 0)
            {
                foreach (var file in filelist)
                {

                    var uploads = Path.Combine(_environment.WebRootPath, "files");
                    string FileName = file.FileName;
                    using (var fileStream = new FileStream(Path.Combine(uploads, FileName), FileMode.Create))
                    {
                        file.CopyToAsync(fileStream);
                    }
                    string output = Path.Combine(Path.Combine(_environment.WebRootPath, "files"), file.FileName.Split(".").ToList().First().ToString()) + ".pdf";
                    string source = Path.Combine(Path.Combine(_environment.WebRootPath, "files"), file.FileName);
                    ConvertWordToSpecifiedFormat(source, output, Word.WdSaveFormat.wdFormatPDF);
                }
 
            }
            
            return View("Index");
        }
        private static void ConvertWordToSpecifiedFormat(object input, object output, object format)
        {
            Word._Application application = new Word.Application();
            application.Visible = false;
            object missing = Missing.Value;
            object isVisible = true;
            object readOnly = false;
            Word._Document document = application.Documents.Open(ref input, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing,
                                    ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);

            document.Activate();
            document.SaveAs(ref output, ref format, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            application.Quit(ref missing, ref missing, ref missing);
        }
    }
}