using SecurityMaster_ST.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace SecurityMaster_ST.Controllers
{
    public class EquityController : ApiController
    {
        string connectionString = ConfigurationManager.ConnectionStrings["IVPDB"].ConnectionString;



    }
}
