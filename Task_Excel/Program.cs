using DocumentFormat.OpenXml.ExtendedProperties;
using System;
using System.Collections.Generic;
using System.Threading;
using Task_Excel.Controllers;
using Task_Excel.Data;
using Task_Excel.Model;

public class Program
{
    public static void Main()
    {
        string numberOperation = SendStartMessage();
        ApplicationData data = new ApplicationData();
        GetFile(data, numberOperation);
        Console.Clear();

        ClientsController clientsController = new ClientsController();
        List<Сlient> cliens = clientsController.ReadDataFromExcel(data.filePath, data.tableNumberClients);

        ProductsController productsController = new ProductsController();
        List<Product> products = productsController.ReadDataFromExcel(data.filePath, data.tableNumberProducts);

        RequestsController requestsController = new RequestsController();
        List<Request> requests = requestsController.ReadDataFromExcel(data.filePath, data.tableNumberRequests);

        UnitOfWork unitOfWork = new UnitOfWork(cliens, products, requests);

        while (data.isOpen)
        {
            Menu(clientsController, unitOfWork, data);
        }
    }

    public static void Restart()
    {
        Main();
    }

    public static string SendStartMessage()
    {
        string numberOperation = "";
        Console.WriteLine("Здравствуйте, уважаемый пользователь");
        Console.WriteLine("Данная программа позволяет взаимодействовать с таблицами Excel определенного формата.");

        Console.WriteLine("\nИспользовать закрепленную таблицу или новую таблицу?");
        Console.WriteLine("\n1. Закрепленная таблица.");
        Console.WriteLine("2. Новая таблицы.\n");


        return numberOperation = Console.ReadLine();
    }
    public static void GetFile(ApplicationData data, string command)
    {
        switch (command)
        {
            case "1":
                Console.WriteLine("Будет использован закрепленный файл");
                break;
            case "2":
                Console.WriteLine("Введите путь до файла с данными (Excel):");

                data.filePath = Console.ReadLine();
                Console.WriteLine("Будет использована новая таблица.");
                Console.WriteLine("Введите номер таблицы Клиентов.");

                string[] numbers = new string[3];
                numbers[0] = Console.ReadLine();
                Console.WriteLine("Введите номер таблицы Продуктов.");
                numbers[1] = Console.ReadLine();
                Console.WriteLine("Введите номер таблицы Заявок.");
                numbers[2] = Console.ReadLine();

                try
                {
                    data.tableNumberClients = Convert.ToInt32(numbers[0]);
                    data.tableNumberProducts = Convert.ToInt32(numbers[1]);
                    data.tableNumberRequests = Convert.ToInt32(numbers[2]);
                }
                catch
                {
                    Console.WriteLine("Неверный формат введенных данных");
                    Restart();
                }

                break;
            default:
                Console.WriteLine("Такой операции нет. Перезагрузка программы...");
                Thread.Sleep(1000);
                Console.Clear();
                Restart();
                break;
        }
    }
    public static void СontinueAndClear()
    {
        Console.WriteLine("\nДля продолжение нажмите любую кнопку...");
        Console.ReadKey();
        Console.Clear();
    }
    public static void Menu(ClientsController clientsController, UnitOfWork unitOfWork, ApplicationData data)
    {
        Console.WriteLine("Выберите команду:\n1. Вывести информацию о клиентах, заказавших товар\n" +
            "2. Изменить контактное лицо клиента\n3. Определить золотого клиента за указанный год и месяц\n4. Выход\n");


        string command = Console.ReadLine();

        switch (command)
        {
            case "1":
                Console.Clear();
                Console.WriteLine("Введите наименование товара:");
                string productName = Console.ReadLine();
                clientsController.GetClientsOrderedProduct(unitOfWork, productName);
                break;
            case "2":
                Console.WriteLine();
                foreach (var client in unitOfWork.Сlients)
                {
                    Console.WriteLine($"Название компании: {client.NameCompany}");
                }
                Console.WriteLine("\nВведите название организации клиента:");
                string organizationName = Console.ReadLine();
                Console.WriteLine("\nВведите ФИО нового контактного лица:\n");
                string newContactPerson = Console.ReadLine();
                bool success = clientsController.ChangeContactPerson(data, unitOfWork, organizationName, newContactPerson);
                Console.WriteLine(success ? "Изменения сохранены." : "Ошибка.");
                СontinueAndClear();
                break;
            case "3":
                Console.WriteLine("\nВведите год:");
                int year = int.Parse(Console.ReadLine());
                Console.WriteLine("\nВведите месяц:");
                int month = int.Parse(Console.ReadLine());
                Сlient goldenСlient = clientsController.GetGoldenСlient(unitOfWork, year, month);
                if (goldenСlient != null)
                {
                    Console.WriteLine(new string('-', 57));
                    Console.WriteLine("Золотой клиент: " + goldenСlient.NameCompany);
                    Console.WriteLine(new string('-', 57));
                    СontinueAndClear();
                }
                else
                {
                    Console.WriteLine(new string('-', 57));
                    Console.WriteLine("Золотой клиент не найден.");
                    Console.WriteLine(new string('-', 57));
                    СontinueAndClear();
                }
                break;
            case "4":
                data.isOpen = false;
                break;
            default:
                Console.WriteLine("Неверная команда. Попробуйте еще раз.");
                СontinueAndClear();
                break;
        }
    }
}


