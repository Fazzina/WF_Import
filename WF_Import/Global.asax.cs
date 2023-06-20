using System;
using System.Web;
using System.Web.Optimization;
using System.Web.Routing;
using System.Web.Http;

namespace WF_Import
{
    public class Global : HttpApplication
    {
        void Application_Start(object sender, EventArgs e)
        {
            // Code that runs on application startup
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);

            // Configura le rotte dell'API
            HttpConfiguration config = GlobalConfiguration.Configuration;
            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );
        }
    }
}