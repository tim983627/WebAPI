using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebAPI.Models
{
    public class BNumber
    {
        public string ItemCode { get; set; }
        public string BatchNumber { get; set; }
        public string Quantity { get; set; }
        public int Count { get; set; }

    }
}