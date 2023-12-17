using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Domain.Entities;

namespace Domain.Models
{
    public class ExcelWorker
    {
        private XLWorkbook _workbook;

        public ExcelWorker(string filePath)
        {
            LoadWotkbook(filePath);
        }

        private void LoadWotkbook(string filePath)
        {
            _workbook = new XLWorkbook(filePath);
        }

        public void OutputCustomerInfo(string productName)
        {
            var productSheet = _workbook.Worksheets.Worksheet("Товары");
            var customerSheet = _workbook.Worksheets.Worksheet("Клиенты");
            var ordersSheet = _workbook.Worksheets.Worksheet("Заявки");

            var productInfo = productSheet.RowsUsed().FirstOrDefault(row => row.Cell(2).GetString() == productName);

            if (productInfo == null)
            {
                Console.WriteLine($"Товар с наименованием '{productName}' не найден.");
                return;
            }

            var productCode = productInfo.FirstCell().GetString();
            var price = productInfo.Cell(4).GetDouble();

            var productOrders = ordersSheet.RowsUsed()
                .Where(row => row.Cell(2).GetString() == productCode)
                .Select(row => new
                {
                    CustomerCode = row.Cell(3).GetString(),
                    Quantity = row.Cell(5).GetDouble(),
                    OrderDate = row.Cell(7).GetString()
                });

            foreach (var order in productOrders)
            {
                var customerInfo = customerSheet.RowsUsed()
                    .Where(row => row.Cell(1).GetString() == order.CustomerCode)
                    .Select(row => new
                    {
                        OrganizationName = row.Cell(2).GetString(),
                        Address = row.Cell(3).GetString(),
                        ContactPerson = row.Cell(4).GetString()
                    })
                    .FirstOrDefault();

                Console.WriteLine($"Код клиента: {order.CustomerCode}");
                Console.WriteLine($"Название организации: {customerInfo?.OrganizationName}");
                Console.WriteLine($"Адрес: {customerInfo?.Address}");
                Console.WriteLine($"Контактное лицо: {customerInfo?.ContactPerson}");
                Console.WriteLine($"Количество: {order.Quantity}");
                Console.WriteLine($"Цена: {price}");
                Console.WriteLine($"Дата заказа: {order.OrderDate}");
                Console.WriteLine("Нажмите любую клавишу чтобы продолжить");
                Console.ReadKey();
            }
        }

        public void UpdateContactPerson(Customer updatedCustomer)
        {
            var customerSheet = _workbook.Worksheets.Worksheet("Клиенты");

            // Находим строку с соответствующим названием организации
            var customerRow = customerSheet.RowsUsed().FirstOrDefault(row => row.Cell(2).GetString() == updatedCustomer.Name);

            if (customerRow == null)
            {
                Console.WriteLine($"Организация '{updatedCustomer.Name}' не найдена.");
                return;
            }

            // Обновляем данные о контактном лице
            customerRow.Cell(4).Value = updatedCustomer.ContactPerson;

            // Сохраняем изменения в Excel-файле
            _workbook.Save();

            Console.WriteLine($"Контактное лицо для клиента '{updatedCustomer.Name}' успешно изменено на '{updatedCustomer.ContactPerson}'.");
            Console.WriteLine("Нажмите любую клавишу чтобы продолжить...");
            Console.ReadKey();
        }

        public void GetGoldenCustomer(DateTime startDate)
        {
            var customersSheet = _workbook.Worksheets.Worksheet("Клиенты");
            var ordersSheet = _workbook.Worksheets.Worksheet("Заявки");

            var goldenCustomer = ordersSheet.RowsUsed()
                .Where(row => {
                    var dateCell = row.Cell(6);
                    DateTime orderDate;
                    return DateTime.TryParse(dateCell.GetString(), out orderDate)
                        && orderDate >= startDate
                        && orderDate.Year == startDate.Year
                        && orderDate.Month == startDate.Month;
                })
                .GroupBy(row => row.Cell(3).GetString())
                .Select(group => new
                {
                    CustomerCode = group.Key,
                    OrderCount = group.Count()
                })
                .OrderByDescending(c => c.OrderCount)
                .FirstOrDefault();

            if (goldenCustomer != null)
            {
                var customerInfo = customersSheet.RowsUsed()
                    .Where(row => row.Cell(1).GetString() == goldenCustomer.CustomerCode)
                    .Select(row => new
                    {
                        Name = row.Cell(2).GetString()
                    })
                    .FirstOrDefault();

                Console.WriteLine($"Золотой клиент (с наибольшим количеством заказов) в {startDate:MMMM yyyy}: {customerInfo?.Name}");
                Console.WriteLine($"Количество заказов: {goldenCustomer.OrderCount}");
            }
            else
            {
                Console.WriteLine($"Нет золотого клиента для указанного месяца и года.");
            }

            Console.WriteLine("Нажмите любую клавишу чтобы продолжить...");
            Console.ReadKey();
        }
    }
}
