using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebAPI.Models
{
    public class BInventory
    {
        public string ItemCode { get; set; }
        public string ItemDesc { get; set; }
        public string WhsCode { get; set; }

        public string Number { get; set; }
    }

}