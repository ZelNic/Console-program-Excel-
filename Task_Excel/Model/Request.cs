using System;

namespace Task_Excel.Model
{
    public class Request
    {
        public int RequestCode { get; set; }
        public int ProductCode { get; set; }
        public int ClientCode { get; set; }
        public int RequestNumber { get; set; }
        public int RequiredQuantity { get; set; }
        public DateTime PostingDate { get; set; }

        public Request(int requestCode, int productCode, int clientCode, int requestNumber, int requiredQuantity, DateTime postingDate)
        {
            RequestCode = requestCode;
            ProductCode = productCode;
            ClientCode = clientCode;
            RequestNumber = requestNumber;
            RequiredQuantity = requiredQuantity;
            PostingDate = postingDate;
        }
    }
}
