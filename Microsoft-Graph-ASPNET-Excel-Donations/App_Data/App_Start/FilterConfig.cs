using System.Web;
using System.Web.Mvc;

namespace Microsoft_Graph_ASPNET_Excel_Donations
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}
