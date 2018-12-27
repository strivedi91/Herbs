using System.Web;
using System.Web.Optimization;

namespace E_InvoiceSolution
{
    public class BundleConfig
    {
        // For more information on bundling, visit https://go.microsoft.com/fwlink/?LinkId=301862
        public static void RegisterBundles(BundleCollection bundles)
        {
            bundles.Add(new ScriptBundle("~/bundles/header-js").Include(
                     "~/Scripts/jquery-1.8.3.min.js",
                     "~/Scripts/modernizr-2.6.2-respond-1.1.0.min.js",
                     "~/Scripts/common.js"));

            bundles.Add(new ScriptBundle("~/bundles/footer-js").Include(
                      "~/Scripts/bootstrap.min.js",
                      "~/Scripts/moment.js",
                      "~/Scripts/bootstrap-datetimepicker.js",
                      "~/Scripts/design-global.js",
                      "~/Scripts/easyResponsiveTabs.js"));

            bundles.Add(new StyleBundle("~/Content/css").Include(
                      "~/Content/bootstrap.css",
                      "~/Content/bootstrap-datetimepicker.min.css",
                      "~/Content/style.css?v=0.4",
                      "~/Content/btn-icon.css",
                      "~/Content/style-responsive.css",
                      "~/Content/sidebar.css",
                      "~/Content/common.css",
                      "~/Content/sidebar.css",
                      "~/Content/stylesheet.css",
                      "~/Content/table-data.css",
                      "~/Content/table_responsive.css",
                      "~/Content/responsivetabs.css",
                      "~/Content/font-awesome.min.css"));
        }
    }
}
