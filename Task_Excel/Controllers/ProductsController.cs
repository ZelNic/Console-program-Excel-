using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Task_Excel.Model;

namespace Task_Excel.Controllers
{
    public class ProductsController
    {
        public List<Product> ReadDataFromExcel(string filePath, int tableNumber)
        {
            List<Product> products = new List<Product>();

            using (var workbook = new XLWorkbook(filePath))
            {

                IXLWorksheet productsTable = null;
                try { productsTable = workbook.Worksheet(tableNumber); } catch { return null; }

                var rows = productsTable.RowsUsed();

                bool firstRowSkipped = false;

                foreach (var row in rows)
                {
                    if (!firstRowSkipped)
                    {
                        firstRowSkipped = true;
                        continue;
                    }

                    int productCode = Convert.ToInt32(row.Cell(1).Value.ToString());
                    string name = row.Cell(2).Value.ToString();
                    string unit = row.Cell(3).Value.ToString();
                    decimal price = Convert.ToDecimal(row.Cell(1).Value.ToString());

                    Product product = new Product(productCode, name, unit, price);
                    products.Add(product);
                }
            }

            return products;
        }
    }
}
