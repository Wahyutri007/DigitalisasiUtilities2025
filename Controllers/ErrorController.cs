using Microsoft.AspNetCore.Mvc;

namespace UtilitiesHR.Controllers
{
    [Route("Error")]
    public class ErrorController : Controller
    {
        // Unhandled exception (500, dll)
        [Route("ServerError")]
        public IActionResult ServerError()
        {
            Response.StatusCode = 500;
            return View("ServerError");
        }

        // Semua status code: 404, 403, 401, dll
        [Route("StatusCode")]
        public IActionResult StatusCodeHandler(int code)
        {
            Response.StatusCode = code;
            ViewData["StatusCode"] = code;

            return code switch
            {
                404 => View("NotFound"),
                403 => View("Forbidden"),
                401 => View("Unauthorized"),
                _ => View("Generic")
            };
        }
    }
}
