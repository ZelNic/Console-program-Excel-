using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.ExtendedProperties;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Task_Excel.Data;
using Task_Excel.Model;

namespace Task_Excel.Controllers
{
    public class ClientsController
    {

        public List<Сlient> ReadDataFromExcel(string filePath, int tableNumber)
        {
            List<Сlient> сlients = new List<Сlient>();

            using (var workbook = new XLWorkbook(filePath))
            {
                IXLWorksheet clientsTable = null;
                try { clientsTable = workbook.Worksheet(tableNumber); } catch { return null; }

                var rows = clientsTable.RowsUsed();

                bool firstRowSkipped = false;

                foreach (var row in rows)
                {
                    if (!firstRowSkipped)
                    {
                        firstRowSkipped = true;
                        continue;
                    }

                    int rowNumber = row.RowNumber();
                    int сlientСode = Convert.ToInt32(row.Cell(1).Value.ToString());
                    string nameCompany = row.Cell(2).Value.ToString();
                    string address = row.Cell(3).Value.ToString();
                    string contactPerson = row.Cell(4).Value.ToString();

                    Сlient сlient = new Сlient(сlientСode, nameCompany, address, contactPerson, rowNumber);

                    сlients.Add(сlient);
                }
            }

            return сlients;
        }
        public void GetClientsOrderedProduct(UnitOfWork unitOfWork, string productName)
        {
            try
            {
                var product = unitOfWork.Products.FirstOrDefault(p => p.Name.ToLower() == productName.ToLower());

                var result = unitOfWork.Requests.Where(r => r.ProductCode == product.ProductCode)
                                            .Join(unitOfWork.Сlients, request => request.ClientCode, client => client.ClientСode, (request, client) => new
                                            {
                                                client.ClientСode,
                                                client.ContactPerson,
                                                client.NameCompany,
                                                client.Address,
                                                request.RequiredQuantity,
                                                PostingDate = request.PostingDate.ToString("dd.MM.yyyy"),
                                            });

                if (result.Any())
                {
                    Console.WriteLine();
                    Console.WriteLine("Результат запроса:");

                    foreach (var r in result)
                    {
                        Console.WriteLine(new string('-', 57));
                        Console.WriteLine($"Код клиента: {r.ClientСode}");
                        Console.WriteLine($"Клиент: {r.ContactPerson}");
                        Console.WriteLine($"Название компание: {r.NameCompany}");
                        Console.WriteLine($"Адресс: {r.Address}");
                        Console.WriteLine($"Количество товара: {r.RequiredQuantity}");
                        Console.WriteLine($"Цена: {product.Price}");
                        Console.WriteLine($"Дата заказа: {r.PostingDate}");
                    }
                    Console.WriteLine(new string('-', 57));
                    Console.WriteLine();
                }
                else
                {
                    Console.WriteLine("Нет результата.");
                    Thread.Sleep(1000);
                    Console.Clear();
                }
            }
            catch
            {
                Console.WriteLine("Нет результата.");
                Thread.Sleep(1000);
            }
        }

        public bool ChangeContactPerson(ApplicationData data, UnitOfWork unitOfWork, string organizationName, string newContactPerson)
        {
            var oldСlientData = unitOfWork.Сlients.FirstOrDefault(c => c.NameCompany.ToLower() == organizationName.ToLower().Trim());

            using (var workbook = new XLWorkbook(data.filePath))
            {
                IXLWorksheet clientsTable = null;
                try { clientsTable = workbook.Worksheet(data.tableNumberClients); } catch { return false; }

                var row = clientsTable.Row(oldСlientData.RowNumber);
                if (row != null)
                {
                    var contactPersonCell = row.Cell(4);
                    contactPersonCell.Value = newContactPerson;
                }
                else
                {
                    return false;
                }

                workbook.Save();
            }

            return true;

        }

        public Сlient GetGoldenСlient(UnitOfWork unitOfWork, int year, int month)
        {
            Dictionary<int, int> customerOrderCounts = new Dictionary<int, int>();

            foreach (var request in unitOfWork.Requests)
            {
                if (request.PostingDate.Year == year && request.PostingDate.Month == month)
                {
                    if (customerOrderCounts.ContainsKey(request.ClientCode))
                    {
                        customerOrderCounts[request.ClientCode]++;
                    }
                    else
                    {
                        customerOrderCounts[request.ClientCode] = 1;
                    }
                }
            }

            int maxOrderCount = 0;
            int goldenCustomer = 0;

            foreach (var kvp in customerOrderCounts)
            {
                if (kvp.Value > maxOrderCount)
                {
                    maxOrderCount = kvp.Value;
                    goldenCustomer = kvp.Key;
                }
            }

            if (goldenCustomer != 0)
            {
                return unitOfWork.Сlients.FirstOrDefault(o => o.ClientСode == goldenCustomer);
            }

            return null;
        }
    }
}
