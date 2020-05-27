using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace selectdemo.Controllers.simplediv
{
    public class simpledivController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public simpledivController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }

        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "../PageOffice/POServer";
            //添加自定义按钮
            pageofficeCtrl.AddCustomToolButton("保存", "SaveFile()", 1);
            pageofficeCtrl.AddCustomToolButton("打印", "PrintFile()", 6);
            pageofficeCtrl.AddCustomToolButton("关闭", "CloseFile()", 21);
            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen("/simplediv/doc/test.doc", PageOfficeNetCore.OpenModeType.docRevisionOnly, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/simplediv/doc/" + fs.FileName);
            fs.Close();
            return Content("OK");
        }
    }
}