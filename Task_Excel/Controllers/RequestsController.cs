using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using Task_Excel.Model;

namespace Task_Excel.Controllers
{
    public class RequestsController
    {
        public List<Request> ReadDataFromExcel(string filePath, int tableNumber)
        {
            List<Request> requests = new List<Request>();

            using (var workbook = new XLWorkbook(filePath))
            {

                IXLWorksheet requestsTable = null;
                try { requestsTable = workbook.Worksheet(tableNumber); } catch { return null; }
                var rows = requestsTable.RowsUsed();

                bool firstRowSkipped = false;

                foreach (var row in rows)
                {
                    if (!firstRowSkipped)
                    {
                        firstRowSkipped = true;
                        continue;
                    }
                    int requestCode = Convert.ToInt32(row.Cell(1).Value.ToString());
                    int productCode = Convert.ToInt32(row.Cell(2).Value.ToString());
                    int clientCode = Convert.ToInt32(row.Cell(3).Value.ToString());
                    int requestNumber = Convert.ToInt32(row.Cell(4).Value.ToString());
                    int requiredQuantity = Convert.ToInt32(row.Cell(5).Value.ToString());
                    DateTime postingDate = Convert.ToDateTime(row.Cell(6).Value.ToString());

                    Request request = new Request(requestCode, productCode, clientCode, requestNumber, requiredQuantity, postingDate);

                    requests.Add(request);
                }
            }

            return requests;
        }
    }
}
