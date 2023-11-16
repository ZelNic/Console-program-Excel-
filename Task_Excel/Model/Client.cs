public class Сlient
{
    public int ClientСode { get; set; }
    public string NameCompany { get; set; }
    public string Address { get; set; }
    public string ContactPerson { get; set; }
    public int RowNumber { get; set; }

    public Сlient(int сlientСode, string nameCompany, string address, string contactPerson, int rowNumber)
    {
        ClientСode = сlientСode;
        NameCompany = nameCompany;
        Address = address;
        ContactPerson = contactPerson;
        RowNumber = rowNumber;
    }
}