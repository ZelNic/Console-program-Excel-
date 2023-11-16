using System.Collections.Generic;

namespace Task_Excel.Model
{
    public class UnitOfWork
    {
        public List<Сlient> Сlients { get; set; }
        public List<Product> Products { get; set; }
        public List<Request> Requests { get; set; }

        public UnitOfWork(List<Сlient> clients, List<Product> products, List<Request> requests)
        {
            Сlients = clients;
            Products = products;
            Requests = requests;
        }
    }
}
