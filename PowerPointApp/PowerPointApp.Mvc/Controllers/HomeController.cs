using System.Web.Mvc;

namespace PowerPointApp.Mvc.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return RedirectToAction("Index", "Presentation");
        }
    }
}
