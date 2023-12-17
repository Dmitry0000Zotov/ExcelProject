using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using Domain.Entities;
using Domain.Models;

//D:\repos\ExcelProject\Практическое задание для кандидата.xlsx
Console.WriteLine("Введите путь к файлу Excel: ");
string filePath = Console.ReadLine();
while (string.IsNullOrEmpty(filePath.Replace(" ", "")))
{
    Console.WriteLine("Путь не может быть пустым!");
    Console.WriteLine("Введите путь к файлу Excel: ");
    filePath = Console.ReadLine();
}

var excelWorker = new ExcelWorker(filePath);
bool exit = false;

do
{
    Console.WriteLine("----------------------------МЕНЮ----------------------------");
    Console.WriteLine("Доступные действия");
    Console.WriteLine("1.Изменить путь к файлу...");
    Console.WriteLine("2.Вывести информацию о клиентах по наименованию товара...");
    Console.WriteLine("3.Изменение контактного лица клиента...");
    Console.WriteLine("4.Узнать \"золотого\" клиента...");
    Console.WriteLine("0.Выйти из программы");

    int choice = GetUserChoice(4);

    switch (choice)
    {
        case 0:
            exit = true;
            break;
        case 1:
            HandleMenuItem1(ref filePath);
            excelWorker = new ExcelWorker(filePath);
            break;
        case 2:
            var productName = HandleMenuItem2();
            excelWorker.OutputCustomerInfo(productName);
            break;
        case 3:
            var updatedCustomer = HandleMenuItem3();
            excelWorker.UpdateContactPerson(updatedCustomer);
            break;
        case 4:
            var startDate = HandleMenuItem4();
            excelWorker.GetGoldenCustomer(startDate);
            break;
        default:
            Console.WriteLine("Некорректный выбор. Попробуйте еще раз.");
            break;
    }
} while (!exit);

static int GetUserChoice(int maxChoice)
{
    int choice;

    do
    {
        Console.Write("Выберите действие (1-{0}): ", maxChoice);
    } while (!int.TryParse(Console.ReadLine(), out choice) || choice < 0 || choice > maxChoice);

    return choice;
}

static void HandleMenuItem1(ref string filePath)
{
    Console.WriteLine($"Текущий путь к файлу - {filePath}");
    Console.WriteLine("Новый путь к файлу: ");
    filePath = Console.ReadLine();
    while (string.IsNullOrEmpty(filePath.Replace(" ", "")))
    {
        Console.WriteLine("Путь не может быть пустым!");
        Console.WriteLine("Введите путь к файлу Excel: ");
        filePath = Console.ReadLine();
    }
}

static string HandleMenuItem2()
{
    Console.WriteLine("------Информация о клиенте по наименованию товара------");
    Console.WriteLine("Введите наименование товара:");
    var productName = Console.ReadLine();
    while (string.IsNullOrEmpty(productName.Replace(" ", "")))
    {
        Console.WriteLine("Наименование товара не может быть пустым. Введите значение еще раз: ");
        productName = Console.ReadLine();
    }

    return productName;
}

static Customer HandleMenuItem3()
{
    Console.WriteLine("------Изменение контактных данных клиента------");
    Console.WriteLine("Введите название организации для изменения контактного лица:");
    var customerName = Console.ReadLine();
    Console.WriteLine($"Введите новое контактное лицо (ФИО) для '{customerName}':");
    var newContactPerson = Console.ReadLine();

    var updatedCustomer = new Customer() 
    { 
        Name = customerName, 
        ContactPerson = newContactPerson 
    };

    return updatedCustomer;
}

static DateTime HandleMenuItem4()
{
    Console.WriteLine("------Золотой клиент------");
    Console.WriteLine("Введите год:");
    var year = Convert.ToInt32(Console.ReadLine());
    Console.WriteLine("Введите месяц:");
    var month = Convert.ToInt32(Console.ReadLine());

    DateTime startDate = new DateTime(year, month, 1);

    return startDate;
}
