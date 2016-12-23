using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;

namespace MvcSample
{
    public class MvcApplication : System.Web.HttpApplication
    {
        private readonly string _licensePath = "c:\\licenses\\GroupDocs.Viewer.lic";

        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);

            //Set GroupDocs.Viewer License
            GroupDocs.Viewer.License license = new GroupDocs.Viewer.License();
            license.SetLicense(_licensePath);
        }
    }
}
