using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Task_Excel.Model
{
    public class Product
    {
        public int ProductCode { get; set; }
        public string Name { get; set; }
        public string Unit { get; set; }
        public decimal Price { get; set; }

        public Product(int productCode, string name, string unit, decimal price)
        {
            ProductCode = productCode;
            Name = name;
            Unit = unit;
            Price = price;
        }
    }
}
