using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebAPI.Models
{
    public  class AInventory
    {
        public string ItemCode { get; set; }
        public string ItemDesc { get; set; }
        public string WhsCode { get; set; }
        public string ABCNumber { get; set; }
        public string Entry { get; set; }
    }

}