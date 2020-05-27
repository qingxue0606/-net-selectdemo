using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace selectdemo.Controllers.simpleselect
{
    public class simpleselectController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public simpleselectController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }
        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "../PageOffice/POServer";
            pageofficeCtrl.Titlebar = false;
            pageofficeCtrl.OfficeToolbars = false;
            pageofficeCtrl.Menubar = false;
            pageofficeCtrl.CustomToolbar = false;

            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen("/simpleselect/doc/test.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }
        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/simpleselect/doc/" + fs.FileName);
            fs.Close();
            return Content("OK");
        }
    }
}