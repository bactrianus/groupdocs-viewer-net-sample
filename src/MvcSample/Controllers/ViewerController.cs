using GroupDocs.Viewer.Config;
using GroupDocs.Viewer.Converter.Options;
using GroupDocs.Viewer.Domain;
using GroupDocs.Viewer.Domain.Html;
using GroupDocs.Viewer.Domain.Options;
using GroupDocs.Viewer.Handler;
using MvcSample.Helpers;
using MvcSample.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mime;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using System.Web.Script.Serialization;
using GroupDocs.Viewer.Domain.Containers;
using WatermarkPosition = MvcSample.Models.WatermarkPosition;

namespace MvcSample.Controllers
{
    public class ViewerController : Controller
    {
        private readonly ViewerHtmlHandler _htmlHandler;
        private readonly ViewerImageHandler _imageHandler;

        // App_Data folder path
        private readonly string _storagePath = AppDomain.CurrentDomain.GetData("DataDirectory").ToString();
        private readonly string _tempPath = AppDomain.CurrentDomain.GetData("DataDirectory") + "\\Temp";

        // Image converter settings
        private const bool UsePdfInImageEngine = true;
        private readonly ConvertImageFileType _convertImageFileType = ConvertImageFileType.JPG;

        private readonly Dictionary<string, Stream> _streams = new Dictionary<string, Stream>();

        public ViewerController()
        {
            var htmlConfig = new ViewerConfig
            {
                StoragePath = _storagePath,
                TempPath = _tempPath,
                UseCache = true
            };

            _htmlHandler = new ViewerHtmlHandler(htmlConfig);

            var imageConfig = new ViewerConfig
            {
                StoragePath = _storagePath,
                TempPath = _tempPath,
                UseCache = true,
                UsePdf = UsePdfInImageEngine
            };

            _imageHandler = new ViewerImageHandler(imageConfig);

            _streams.Add("ProcessFileFromStreamExample_1.pdf", HttpWebRequest.Create("http://unfccc.int/resource/docs/convkp/kpeng.pdf").GetResponse().GetResponseStream());
            _streams.Add("ProcessFileFromStreamExample_2.doc", HttpWebRequest.Create("http://www.acm.org/sigs/publications/pubform.doc").GetResponse().GetResponseStream());
        }

        [HttpPost]
        public ActionResult ViewDocument(ViewDocumentParameters request)
        {
            if (Utils.IsValidUrl(request.Path))
                request.Path = DownloadToStorage(request.Path);
            else if (_streams.ContainsKey(request.Path))
                request.Path = SaveStreamToStorage(request.Path);

            var result = new ViewDocumentResponse
            {
                pageCss = new string[] { },
                lic = true,
                pdfDownloadUrl = GetPdfDownloadUrl(request),
                pdfPrintUrl = GetPdfPrintUrl(request),
                url = GetFileUrl(request),
                path = request.Path,
                name = Path.GetFileName(request.Path)
            };

            if (request.UseHtmlBasedEngine)
                ViewDocumentAsHtml(request, result);
            else
                ViewDocumentAsImage(request, result);

            return new LargeJsonResult { Data = result };
        }

        public ActionResult LoadFileBrowserTreeData(LoadFileBrowserTreeDataParameters parameters)
        {
            var path = _storagePath;
            if (!string.IsNullOrEmpty(parameters.Path))
                path = Path.Combine(path, parameters.Path);

            var request = new FileTreeOptions(path);
            var tree = _htmlHandler.LoadFileTree(request);

            var result = new FileBrowserTreeDataResponse
            {
                nodes = Utils.ToFileTreeNodes(parameters.Path, tree.FileTree).ToArray(),
                count = tree.FileTree.Count
            };

            return ToJsonResult(result);
        }

        public ActionResult GetImageUrls(GetImageUrlsParameters parameters)
        {
            var guid = parameters.Path;

            // Get document info
            var documentInfoContainer = _imageHandler.GetDocumentInfo(new DocumentInfoOptions(guid));
            
            var pageNumbers = documentInfoContainer.Pages.Select(_ => _.Number).ToArray();
            var applicationHost = GetApplicationHost();

            // Get image urls
            string[] imageUrls = ImageUrlHelper.GetImageUrls(applicationHost, pageNumbers, parameters);

            return ToJsonResult(new GetImageUrlsResult(imageUrls));
        }

        public ActionResult GetFile(GetFileParameters parameters)
        {
            var displayName = string.IsNullOrEmpty(parameters.DisplayName)
                ? Path.GetFileName(parameters.Path)
                : Uri.EscapeDataString(parameters.DisplayName);

            Stream fileStream;
            if (parameters.GetPdf)
            {
                displayName = Path.ChangeExtension(displayName, "pdf");

                var options = new PdfFileOptions
                {
                    Guid = parameters.Path,
                    Watermark = Utils.GetWatermark(parameters.WatermarkText, parameters.WatermarkColor, parameters.WatermarkPosition, parameters.WatermarkWidth),
                };

                if (parameters.IsPrintable)
                    options.AddPrintAction = true;

                if (parameters.SupportPageRotation)
                    options.Transformations |= Transformation.Rotate;

                options.Transformations |= Transformation.Reorder;

                var pdfFileResponse = _htmlHandler.GetPdfFile(options);
                fileStream = pdfFileResponse.Stream;
            }
            else
            {
                var fileResponse = _htmlHandler.GetFile(parameters.Path);
                fileStream = fileResponse.Stream;
            }

            //jquery.fileDownload uses this cookie to determine that a file download has completed successfully
            Response.SetCookie(new HttpCookie("jqueryFileDownloadJSForGD", "true") { Path = "/" });

            fileStream.Position = 0;
            using (var ms = new MemoryStream())
            {
                fileStream.CopyTo(ms);
                return File(ms.ToArray(), "application/octet-stream", displayName);
            }
        }

        public ActionResult GetPdfWithPrintDialog(GetFileParameters parameters)
        {
            var displayName = string.IsNullOrEmpty(parameters.DisplayName)
                ? Path.GetFileName(parameters.Path)
                : Uri.EscapeDataString(parameters.DisplayName);

            var options = new PdfFileOptions
            {
                Guid = parameters.Path,
                Watermark = Utils.GetWatermark(parameters.WatermarkText, parameters.WatermarkColor, parameters.WatermarkPosition, parameters.WatermarkWidth)
            };

            if (parameters.IsPrintable)
                options.AddPrintAction = true;

            if (parameters.SupportPageRotation)
                options.Transformations |= Transformation.Rotate;
            options.Transformations |= Transformation.Reorder;

            var response = _htmlHandler.GetPdfFile(options);

            var contentDispositionString = new ContentDisposition { FileName = displayName, Inline = true }.ToString();
            Response.AddHeader("Content-Disposition", contentDispositionString);

            return File(((MemoryStream)response.Stream).ToArray(), "application/pdf");
        }

        public string GetFileUrl(string path, bool getPdf, bool isPrintable, string fileDisplayName = null,
          string watermarkText = null, int? watermarkColor = null,
          WatermarkPosition? watermarkPosition = WatermarkPosition.Diagonal, float? watermarkWidth = 0,
          bool ignoreDocumentAbsence = false,
          bool useHtmlBasedEngine = false,
          bool supportPageRotation = false)
        {
            var queryString = HttpUtility.ParseQueryString(string.Empty);
            queryString["path"] = path;
            if (!isPrintable)
            {
                queryString["getPdf"] = getPdf.ToString().ToLower();

                if (fileDisplayName != null)
                    queryString["displayName"] = fileDisplayName;
            }

            if (watermarkText != null)
            {
                queryString["watermarkText"] = watermarkText;
                queryString["watermarkColor"] = watermarkColor.ToString();

                if (watermarkPosition.HasValue)
                    queryString["watermarkPosition"] = watermarkPosition.ToString();

                if (watermarkWidth.HasValue)
                    queryString["watermarkWidth"] = ((float)watermarkWidth).ToString(CultureInfo.InvariantCulture);
            }

            if (ignoreDocumentAbsence)
            {
                queryString["ignoreDocumentAbsence"] = ignoreDocumentAbsence.ToString().ToLower();
            }

            queryString["useHtmlBasedEngine"] = useHtmlBasedEngine.ToString().ToLower();
            queryString["supportPageRotation"] = supportPageRotation.ToString().ToLower();

            var handlerName = isPrintable ? "GetPdfWithPrintDialog" : "GetFile";

            var baseUrl = Request.Url.Scheme + "://" + Request.Url.Authority + Request.ApplicationPath.TrimEnd('/') + "/document-viewer/";

            var fileUrl = string.Format("{0}{1}?{2}", baseUrl, handlerName, queryString);
            return fileUrl;
        }

        public ActionResult GetDocumentPageHtml(GetDocumentPageHtmlParameters parameters)
        {
            if (Utils.IsValidUrl(parameters.Path))
                parameters.Path = Utils.GetFilenameFromUrl(parameters.Path);
            if (string.IsNullOrWhiteSpace(parameters.Path))
                throw new ArgumentException("A document path must be specified", "path");

            string guid = parameters.Path;
            int pageNumber = parameters.PageIndex + 1;

            HtmlOptions htmlOptions = new HtmlOptions
            {
                PageNumber = pageNumber,
                CountPagesToConvert = 1,
                IsResourcesEmbedded = false,
                HtmlResourcePrefix = GetHtmlResourcePrefix(guid),
            };

            HtmlPageContent pageContent = GetHtmlPageContents(guid, htmlOptions).Single();

            var result = new GetDocumentPageHtmlResult
            {
                pageHtml = pageContent.Html,
                pageCss = pageContent.Css
            };
            return ToJsonResult(result);
        }

        public ActionResult GetDocumentPageImage(GetDocumentPageImageParameters parameters)
        {
            var guid = parameters.Path;
            var pageIndex = parameters.PageIndex;
            var pageNumber = pageIndex + 1;

            var imageOptions = new ImageOptions
            {
                ConvertImageFileType = _convertImageFileType,
                Watermark = Utils.GetWatermark(parameters.WatermarkText, parameters.WatermarkColor, 
                parameters.WatermarkPosition, parameters.WatermarkWidth),
                Transformations = parameters.Rotate ? Transformation.Rotate : Transformation.None,
                CountPagesToConvert = 1,
                PageNumber = pageNumber,
                JpegQuality = parameters.Quality.GetValueOrDefault()
            };

            if (parameters.Rotate && parameters.Width.HasValue)
            {
                DocumentInfoContainer documentInfoContainer = _imageHandler.GetDocumentInfo(new DocumentInfoOptions(guid));

                int pageAngle = documentInfoContainer.Pages[pageIndex].Angle;
                var isHorizontalView = pageAngle == 90 || pageAngle == 270;

                int sideLength = parameters.Width.Value;
                if (isHorizontalView)
                    imageOptions.Height = sideLength;
                else
                    imageOptions.Width = sideLength;
            } else if (parameters.Width.HasValue)
            {
                imageOptions.Width = parameters.Width.Value;
            }

            var pageImage = _imageHandler.GetPages(guid, imageOptions).Single();

            return File(pageImage.Stream, Utils.GetMimeType(_convertImageFileType));
        }

        public ActionResult GetResourceForHtml(GetResourceForHtmlParameters parameters)
        {
            if (!string.IsNullOrEmpty(parameters.ResourceName) &&
                parameters.ResourceName.IndexOf("/", StringComparison.Ordinal) >= 0)
                parameters.ResourceName = parameters.ResourceName.Replace("/", "");

            var resource = new HtmlResource
            {   
                ResourceName = parameters.ResourceName,
                ResourceType = Utils.GetResourceType(parameters.ResourceName),
                DocumentPageNumber = parameters.PageNumber == 0 ? 1 : parameters.PageNumber
            };

            var stream = _htmlHandler.GetResource(parameters.DocumentPath, resource);

            if (stream == null || stream.Length == 0)
                return new HttpStatusCodeResult((int)HttpStatusCode.Gone);

            return File(stream, Utils.GetMimeType(parameters.ResourceName));
        }

        public ActionResult RotatePage(RotatePageParameters parameters)
        {
            string guid = parameters.Path;
            int pageIndex = parameters.PageNumber;

            DocumentInfoContainer documentInfoContainer = _imageHandler.GetDocumentInfo(new DocumentInfoOptions(guid));
            int pageNumber = documentInfoContainer.Pages[pageIndex].Number;

            RotatePageOptions rotatePageOptions = new RotatePageOptions(guid, pageNumber, parameters.RotationAmount);
            RotatePageContainer rotatePageContainer = _imageHandler.RotatePage(rotatePageOptions);

            RotatePageResponse response = new RotatePageResponse
            {
                resultAngle = rotatePageContainer.CurrentRotationAngle
            };

            return ToJsonResult(response);
        }

        public ActionResult ReorderPage(ReorderPageParameters parameters)
        {
            string guid = parameters.Path;

            DocumentInfoContainer documentInfoContainer = _imageHandler.GetDocumentInfo(new DocumentInfoOptions(guid));

            int pageNumber = documentInfoContainer.Pages[parameters.OldPosition].Number;
            int newPosition = parameters.NewPosition + 1;

            ReorderPageOptions reorderPageOptions = new ReorderPageOptions(guid, pageNumber, newPosition);
            _imageHandler.ReorderPage(reorderPageOptions);

            return ToJsonResult(new ReorderPageResponse());
        }

        private List<HtmlPageContent> GetHtmlPageContents(string guid, HtmlOptions htmlOptions)
        {
            var pageContents = new List<HtmlPageContent>();

            var documentInfo = _htmlHandler.GetDocumentInfo(new DocumentInfoOptions(guid));

            var htmlPages = _htmlHandler.GetPages(guid, htmlOptions);
            foreach (var page in htmlPages)
            {
                var html = page.HtmlContent;

                var indexOfBodyOpenTag = html.IndexOf("<body>", StringComparison.InvariantCultureIgnoreCase);
                if (indexOfBodyOpenTag > 0)
                    html = html.Substring(indexOfBodyOpenTag + "<body>".Length);

                var indexOfBodyCloseTag = html.IndexOf("</body>", StringComparison.InvariantCultureIgnoreCase);
                if (indexOfBodyCloseTag > 0)
                    html = html.Substring(0, indexOfBodyCloseTag);

                string css = string.Empty;
                foreach (var resource in page.HtmlResources.Where(_ => _.ResourceType == HtmlResourceType.Style))
                {
                    var resourceStream = _htmlHandler.GetResource(guid, resource);
                    var resourceContent = new StreamReader(resourceStream).ReadToEnd();

                    if (!string.IsNullOrEmpty(css))
                        css += " ";

                    css += resourceContent;
                }

                // wrap single image tags
                var match = Regex.Match(html, "^<img.+?src=[\"'](.+?)[\"'].*?>$", RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    var src = match.Groups[1].Value;
                    var pageData = documentInfo.Pages.Single(_ => _.Number == page.PageNumber);

                    css = ".grpdx .ie .doc-page {font-size:0;}";
                    html = string.Format("<div style='width:{0}px;height:{1}px;font-size:0'>" +
                        "<img style='width:{0}px;height:{1}px;font-size:0' src='{2}'/>" +
                        "</div>",
                        pageData.Width,
                        pageData.Height,
                        src);
                }

                //wrap svg tags
                if (html.StartsWith("<svg"))
                    html = "<div>" + html + "</div>";

                pageContents.Add(new HtmlPageContent(html, css));
            }

            return pageContents;
        }

        private ActionResult ToJsonResult(object result)
        {
            var serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var serializedData = serializer.Serialize(result);
            return Content(serializedData, "application/json");
        }

        private string DownloadToStorage(string url)
        {
            var fileNameFromUrl = Utils.GetFilenameFromUrl(url);
            var filePath = Path.Combine(_storagePath, fileNameFromUrl);

            using (new InterProcessLock(filePath))
                Utils.DownloadFile(url, filePath);

            return fileNameFromUrl;
        }

        private string SaveStreamToStorage(string key)
        {
            var stream = _streams[key];
            var savePath = Path.Combine(_storagePath, key);

            using (new InterProcessLock(savePath))
            {
                using (var fileStream = System.IO.File.Create(savePath))
                {
                    if (stream.CanSeek)
                        stream.Seek(0, SeekOrigin.Begin);
                    stream.CopyTo(fileStream);
                }

                return savePath;
            }
        }

        private void ViewDocumentAsImage(ViewDocumentParameters request, ViewDocumentResponse result)
        {
            var guid = request.Path;

            // Get document info
            var documentInfo = _imageHandler.GetDocumentInfo(new DocumentInfoOptions(guid));

            // Serialize document info
            SerializationOptions serializationOptions = new SerializationOptions
            {
                UsePdf = request.UsePdf,
                SupportListOfBookmarks = request.SupportListOfBookmarks,
                SupportListOfContentControls = request.SupportListOfContentControls
            };
            var documentInfoJson = new DocumentInfoJsonSerializer(documentInfo, serializationOptions).Serialize();

            // Build result
            result.documentDescription = documentInfoJson;
            result.docType = documentInfo.DocumentType;
            result.fileType = documentInfo.FileType;
            int[] pageNumbers = documentInfo.Pages.Select(_ => _.Number).ToArray();
            result.imageUrls = ImageUrlHelper.GetImageUrls(GetApplicationHost(), pageNumbers, request);
        }

        private void ViewDocumentAsHtml(ViewDocumentParameters request, ViewDocumentResponse result)
        {
            var guid = request.Path;
            var fileName = Path.GetFileName(request.Path);

            // Get document info
            var documentInfo = _htmlHandler.GetDocumentInfo(new DocumentInfoOptions(guid));

            // Serialize document info
            SerializationOptions serializationOptions = new SerializationOptions
            {
                UsePdf = false,
                SupportListOfBookmarks = request.SupportListOfBookmarks,
                SupportListOfContentControls = request.SupportListOfContentControls
            };
            var documentInfoJson = new DocumentInfoJsonSerializer(documentInfo, serializationOptions).Serialize();

            // Build html options
            var htmlOptions = new HtmlOptions
            {
                IsResourcesEmbedded = Utils.IsImage(fileName),
                HtmlResourcePrefix =  GetHtmlResourcePrefix(guid),
                PageNumber = 1,
                CountPagesToConvert = request.PreloadPagesCount.GetValueOrDefault(1)
            };
            var htmlPageContents = GetHtmlPageContents(guid, htmlOptions);

            // Build result
            result.pageHtml = htmlPageContents
                .Select(_ => _.Html)
                .ToArray();
            result.pageCss = htmlPageContents
                .Where(_ => !string.IsNullOrEmpty(_.Css))
                .Select(_ => _.Css)
                .ToArray();
            result.documentDescription = documentInfoJson;
            result.docType = documentInfo.DocumentType;
            result.fileType = GetFileTypeOrEmptyString(documentInfo.FileType);
        }

        private string GetFileTypeOrEmptyString(string fileType)
        {
            var textFileTypes = new[] {"txt", "htm", "html", "xml"};

            if (textFileTypes.Contains(fileType, StringComparer.InvariantCultureIgnoreCase))
                return string.Empty;

            return fileType;
        }

        private string GetFileUrl(ViewDocumentParameters request)
        {
            return GetFileUrl(request.Path, false, false, request.FileDisplayName);
        }

        private string GetPdfPrintUrl(ViewDocumentParameters request)
        {
            return GetFileUrl(request.Path, true, true, request.FileDisplayName,
                request.WatermarkText, request.WatermarkColor,
                request.WatermarkPosition, request.WatermarkWidth,
                request.IgnoreDocumentAbsence,
                request.UseHtmlBasedEngine, request.SupportPageRotation);
        }

        private string GetPdfDownloadUrl(ViewDocumentParameters request)
        {
            return GetFileUrl(request.Path, true, false, request.FileDisplayName,
                request.WatermarkText, request.WatermarkColor,
                request.WatermarkPosition, request.WatermarkWidth,
                request.IgnoreDocumentAbsence,
                request.UseHtmlBasedEngine, request.SupportPageRotation);
        }

        private string GetApplicationHost()
        {
            return Request.Url.Scheme + "://" + Request.Url.Authority + Request.ApplicationPath.TrimEnd('/');
        }

        private string GetHtmlResourcePrefix(string guid)
        {
            const string format =
                "/document-viewer/GetResourceForHtml?documentPath={0}&pageNumber={{page-number}}&resourceName=";

            return string.Format(format: format, arg0: HttpUtility.UrlEncode(guid));
        }
    }
}