using GroupDocs.Viewer.Config;
using GroupDocs.Viewer.Converter.Option;
using GroupDocs.Viewer.Domain;
using GroupDocs.Viewer.Domain.Requests;
using GroupDocs.Viewer.Handler;
using GroupDocs.Viewer.Helper;
using MvcSample.Models;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Mvc;
using System.Web.Script.Serialization;

namespace MvcSample.Controllers
{
    public class ViewerController : Controller
    {
        private readonly ViewerHtmlHandler _htmlHandler;
        private readonly ViewerImageHandler _imageHandler;

        public ViewerController()
        {
            var config = new ViewerConfig
            {
                StoragePath = @"C:\storage"
            };

            _htmlHandler = new ViewerHtmlHandler(config);
            _imageHandler = new ViewerImageHandler(config);

            _htmlHandler.SetLicense(@"D:\GroupDocs.Viewer.lic");
        }

        // GET: /Viewer/
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult ViewDocument(ViewDocumentParameters request)
        {
            var docInfo = _htmlHandler.GetDocumentInfo(new GetDocumentInfoRequest(request.Path));
            var docInfoAsJson = new FileDataJsonSerializer(docInfo.FileData, new FileDataOptions()).Serialize();

            var result = new ViewDocumentResponse
            {
                documentDescription = docInfoAsJson,
                pageCss = new string[] { },
                lic = true
            };

            if (request.UseHtmlBasedEngine)
            {
                var htmlPages = _htmlHandler.GetPages(new FileDescription { Guid = request.Path, Name = request.Path }, new HtmlOptions());
                result.pageHtml = htmlPages.Select(_ => _.HtmlContent).ToArray();
            }
            else
            {
                var imagePages = _imageHandler.GetPages(new FileDescription { Guid = request.Path, Name = request.Path }, new ImageOptions());

                // Save images some where and provide urls
                var urls = new List<string>();
                var tempFolderPath = Path.Combine(Server.MapPath("~"), "Content", "TempStorage");

                foreach (var pageImage in imagePages)
                {
                    var docFoldePath = Path.Combine(tempFolderPath, request.Path);

                    if (!Directory.Exists(docFoldePath))
                        Directory.CreateDirectory(docFoldePath);

                    var pageImageName = string.Format("{0}\\{1}.png", docFoldePath, pageImage.PageNumber);

                    using (var stream = pageImage.Stream)
                    using (FileStream fileStream = new FileStream(pageImageName, FileMode.Create))
                    {
                        stream.Seek(0, SeekOrigin.Begin);
                        stream.CopyTo(fileStream);
                    }

                    var baseUrl = Request.Url.Scheme + "://" + Request.Url.Authority + Request.ApplicationPath.TrimEnd('/') + "/";
                    urls.Add(string.Format("{0}Content/TempStorage/{1}/{2}.png", baseUrl, request.Path, pageImage.PageNumber));
                }
                result.imageUrls = urls.ToArray();
            }

            var serializedData = new JavaScriptSerializer().Serialize(result);
            return Content(serializedData, "application/json");
        }
    }
}