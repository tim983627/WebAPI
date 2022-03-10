using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebAPI.Models
{
    public class Export
    {
        public string ItemCode { get; set; }
        public string WareHouse { get; set; }
        public int Quantity { get; set; }
        public int Price { get; set; }
    }
}