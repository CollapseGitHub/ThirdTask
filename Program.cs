

using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Runtime.CompilerServices;

namespace ThirdTask
{
    internal class Program
    {
        static XlsxManagement manager = new XlsxManagement();
        static void Main(string[] args)
        {
            Console.Title = "XLSX READER";
            Console.WriteLine("Добро пожаловать!");
            start:
            manager.PathOfFile = GetPathFromUser();
            manager.DB = manager.GetInfoFromDoc(manager.PathOfFile);
            if (manager.DB == null) //Обработка исключения Файл используется другим процессом.
            {
                Console.WriteLine("Возникло исключение: Файл занят другим процессом, освободите файл и повторите попытку");
                goto start;
            }
            MenuOptions();
            Console.ReadKey();
        }

        /// <summary>
        /// Метод получения пути до файла, проверка введеных пользователем данных
        /// </summary>
        /// <returns>Возвращает путь до существующего файла</returns>
        static string GetPathFromUser()
        {
        start:
            Console.WriteLine("\nУкажите путь до файла с данными");
            Console.Write("Путь: ");
            string userStringFromConsole = Console.ReadLine().Trim();
            #region Проверки
            if (string.IsNullOrEmpty(userStringFromConsole)) //Проверка на пустую строку или null значение p.s. перевести все проверки в один метод
            {
                Console.WriteLine("Пути к файлу не может быть пустым, проверьте правильность ввода\n");
                goto start;
            }

            if (System.IO.Path.Exists(userStringFromConsole) == false) //Проверка на существование файла
            {
                Console.WriteLine($"Файл по пути {userStringFromConsole} не существует.\nПроверьте расположение файла и повторите ввод\n");
                goto start;
            }

            if (userStringFromConsole.Contains(".xlsx")) //Проверка на наличие расширения файла в пути
            {
                return userStringFromConsole;
            }
            Console.WriteLine("В введенном вами пути не найдено расширение файла");
            goto start;
            #endregion
        }

        /// <summary>
        /// Метод отображения меню выбора функций
        /// </summary>
        static void MenuOptions()
        {
            Console.Clear();
        start:
            Console.WriteLine("\nДля выбора пункта меню напишите цифру необходимо пункта" +
                "\n1. Информация о заказах клиентов по наименованию товара" +
                "\n2. Просмотр и изменение списка клиентов" +
                "\n3. Определить \"золотого\" клиента");
            string choosedOption = Console.ReadLine().Trim();
            #region Проверки
            if (string.IsNullOrEmpty(choosedOption) || choosedOption.Length > 1) //Проверка на пустую строку или значения, которые больше одного символа
            {
                Console.WriteLine("\nПожалуйста напишите одну цифру необходимого пункта меню");
                goto start;
            }
            if (!int.TryParse(choosedOption, out int number)) //Проверка на введение буквы
            {
                Console.WriteLine("\nВы ввели букву, введите цифру пункта меню");
                goto start;
            }
            #endregion
            switch (choosedOption)
            {
                case "1": { MenuOfProducts(); break; }
                case "2": { MenuOfClients();  break; }
                case "3": { MenuOfGoldenClient(); break; }
                default:
                    {
                        Console.WriteLine("\nВ меню нет пункта " + choosedOption + "\nПроверьте правильность ввода"); //Обработка случая,
                                                                                                                      //когда введены несуществующие пункты меню
                        goto start;
                    }
            }

        }

        #region 1 Пункт меню
        /// <summary>
        /// Метод выбора продукта для запроса по наименованию
        /// </summary>
        static void MenuOfProducts()
        {
            Console.Clear();
            int index;
            Dictionary<int, string> products = new Dictionary<int, string>(); //Временный словарь, где keys - номер продукта, value - код товара из файла
        start:
            index = 1;
            Console.WriteLine("Введите пункт из меню товара для поиска");
            foreach (var item in manager.DB[0])
            {
                Console.WriteLine(index + ": " + item.Value.ToString());
                if (products.Keys.Contains(index) == false) //Если в products уже присутсвует индекс, значит словарь заполнен.
                {
                    products.Add(index, item.Key.ToString());
                    index++;
                }
                else index++;
            }
            string selectedProduct = Console.ReadLine().Trim();
            #region Проверки
            if (string.IsNullOrEmpty(selectedProduct) || selectedProduct.Length > 2) //Проверка на пустую строку или значения, которые больше одного символа
            {
                Console.WriteLine("\nПожалуйста напишите цифру необходимого пункта товара");
                goto start;
            }
            if (!int.TryParse(selectedProduct, out int number)) //Проверка на введение буквы
            {
                Console.WriteLine("\nВы ввели букву, введите цифру пункта товара");
                goto start;
            }
            if (products.Keys.Contains(int.Parse(selectedProduct)) == false)
            {
                Console.WriteLine("\nВ меню нет пункта " + selectedProduct + "\nПроверьте правильность ввода\n"); //Обработка случая,
                                                                                                                  //когда введены несуществующие пункты меню
                goto start;
            }
            #endregion
            GetInfoSelectedProduct(products[int.Parse(selectedProduct)]); // Передача управлению другому методу, в параметрах передаём код выбранного товара.
        }

        /// <summary>
        /// Метод для получения информации о заявках, клиенте, сумме заказа и т.д.
        /// </summary>
        /// <param name="codeOfProduct">Код выбравнного товара</param>
        static void GetInfoSelectedProduct(string codeOfProduct)
        {
            Console.Clear();
            int count = 0;
            foreach (var item in manager.DB[5])
            {
                if (item.Value.Contains(codeOfProduct) == true)
                {
                    string[] mainRequestInfo = item.Value.ToString().Split(',');
                    DateTime thisDate = DateTime.Parse(mainRequestInfo[4]); //Задаём формат даты p.s. изменить формат записи в бд, чтобы тут убрать
                    CultureInfo culture = new CultureInfo("ru-RU");
                    Console.WriteLine($"Код заявки - {item.Key} " + 
                        $"Клиент - {manager.DB[2][mainRequestInfo[1]]} " + 
                        $"\nКонтактное лицо - {manager.DB[4][mainRequestInfo[1]]} " + 
                        $"\nАдрес организации - {manager.DB[3][mainRequestInfo[1]]} " +
                        $"\nКоличество заказанного товара - {mainRequestInfo[3]} " + 
                        $"\nСумма заказа - {int.Parse(manager.DB[1][codeOfProduct]) * int.Parse(mainRequestInfo[3])} " +
                        $"\nДата размещения товара - {thisDate.ToString("d",culture)}\n");
                    count++;
                }
            }
            if (count == 0) Console.WriteLine("Данный товар пока что не заказывали =("+ 
                "\n\nВыбрать другой товар или вернуться в меню?\n1 - Выбрать другой товар\n2 - Вернуться в меню");
            else Console.WriteLine("\n\nВыбрать другой товар или вернуться в меню?\n1 - Выбрать другой товар\n2 - Вернуться в меню");
            #region Проверки
            userChoose:
            string userResult = Console.ReadLine().Trim();
            if (string.IsNullOrEmpty(userResult) || userResult.Length > 1) //Проверка на пустую строку или значения, которые больше одного символа
            {
                Console.WriteLine("\nПожалуйста напишите одну цифру необходимого пункта меню");
                goto userChoose;
            }
            if (!int.TryParse(userResult, out int number)) //Проверка на введение буквы
            {
                Console.WriteLine("\nВы ввели букву, введите цифру пункта меню");
                goto userChoose;
            }
            switch (userResult)
            {
                case "1": { MenuOfProducts(); break; }
                case "2": { MenuOptions(); break; }
                default:
                    {
                        Console.WriteLine("\nВ меню нет пункта " + userResult + "\nПроверьте правильность ввода"); //Обработка случая,
                                                                                                                   //когда введены несуществующие пункты меню
                        goto userChoose;
                    }
            }
            #endregion
        }
        #endregion

        #region 2 Пункт меню
        /// <summary>
        /// Метод выбора клиента для запроса по изменению ФИО контактного лица
        /// </summary>
        static void MenuOfClients()
        {
            Console.Clear();
            int index;
            Dictionary<int, string> clients = new Dictionary<int, string>(); //Временный словарь, где keys - номер клиента, value - код клиента
        start:
            index = 1;
            Console.WriteLine("Введите пункт из меню выбора клиента для изменения данных\n");
            foreach (var item in manager.DB[4])
            {
                Console.WriteLine(index + ". Организация: " + manager.DB[2][item.Key] + " Контактное лицо: " + item.Value.ToString());
                if (clients.Keys.Contains(index) == false)
                {
                    clients.Add(index, item.Key.ToString());
                    index++;
                }
                else index++;
            }
            string selectedClient = Console.ReadLine().Trim();
            #region Проверки
            if (string.IsNullOrEmpty(selectedClient) || selectedClient.Length > 2) //Проверка на пустую строку или значения, которые больше одного символа
            {
                Console.WriteLine("\nПожалуйста напишите цифру необходимого пункта товара");
                goto start;
            }
            if (!int.TryParse(selectedClient, out int number)) //Проверка на введение буквы
            {
                Console.WriteLine("\nВы ввели букву, введите цифру пункта товара");
                goto start;
            }
            if (clients.Keys.Contains(int.Parse(selectedClient)) == false)
            {
                Console.WriteLine("\nВ меню нет пункта " + selectedClient + "\nПроверьте правильность ввода\n"); //Обработка случая,
                                                                                                                  //когда введены несуществующие пункты меню
                goto start;
            }
            #endregion
            SetInfoSelectedClient(clients[int.Parse(selectedClient)]); // Передача управления другому методу, в параметрах передаём код выбранного клиента.
        }

        /// <summary>
        /// Метод для изменения ФИО контактного лица
        /// </summary>
        /// <param name="codeOfClient">Код выбравнной компании</param>
        static void SetInfoSelectedClient(string codeOfClient)
        {
            Console.Clear();
            Console.WriteLine($"Введите новое ФИО контактного лица компании {manager.DB[2][codeOfClient]}");
            Console.Write("ФИО: ");
            string newFIO = Console.ReadLine().Trim();
            Console.WriteLine($"\nКонтактные данные организации {manager.DB[2][codeOfClient]} успешно изменены с " + 
                manager.SetFIOClient(manager.PathOfFile, codeOfClient, newFIO));
            Console.WriteLine("\n\nВыбрать другого клиента или вернуться в меню?\n1 - Выбрать другого клиента\n2 - Вернуться в меню");
        #region Проверки
        userChoose:
            string userResult = Console.ReadLine().Trim();
            if (string.IsNullOrEmpty(userResult) || userResult.Length > 1) //Проверка на пустую строку или значения, которые больше одного символа
            {
                Console.WriteLine("\nПожалуйста напишите одну цифру необходимого пункта меню");
                goto userChoose;
            }
            if (!int.TryParse(userResult, out int number)) //Проверка на введение буквы
            {
                Console.WriteLine("\nВы ввели букву, введите цифру пункта меню");
                goto userChoose;
            }
            switch (userResult)
            {
                case "1": { MenuOfClients(); break; }
                case "2": { MenuOptions(); break; }
                default:
                    {
                        Console.WriteLine("\nВ меню нет пункта " + userResult + "\nПроверьте правильность ввода"); //Обработка случая,
                                                                                                                   //когда введены несуществующие пункты меню
                        goto userChoose;
                    }
            }
            #endregion
        }
        #endregion

        static void MenuOfGoldenClient() 
        {
            Console.Clear();
            Console.WriteLine("За какой период опредедляем \"золотого\" клиента?\n1 - Месяц\n2 - Год");
            #region Проверки
        userChoose:
            string userResult = Console.ReadLine().Trim();
            if (string.IsNullOrEmpty(userResult) || userResult.Length > 1) //Проверка на пустую строку или значения, которые больше одного символа
            {
                Console.WriteLine("\nПожалуйста напишите одну цифру необходимого пункта меню");
                goto userChoose;
            }
            if (!int.TryParse(userResult, out int number)) //Проверка на введение буквы
            {
                Console.WriteLine("\nВы ввели букву, введите цифру пункта меню");
                goto userChoose;
            }
            switch (userResult)
            {
                case "1": { MonthGoldenClient(); break; }
                case "2": { MenuOptions(); break; }
                default:
                    {
                        Console.WriteLine("\nВ меню нет пункта " + userResult + "\nПроверьте правильность ввода"); //Обработка случая,
                                                                                                                   //когда введены несуществующие пункты меню
                        goto userChoose;
                    }
            }
            #endregion
        }

        static void MonthGoldenClient()
        {
            Console.Clear();
            CultureInfo culture = new CultureInfo("ru-RU");
            DateTimeFormatInfo dtfi = culture.DateTimeFormat;
            Dictionary<int,string> months = new Dictionary<int,string>();
            int index = 1;
            Console.WriteLine("Для какого месяца определяем \"золотого\" клиента?");
            foreach (var item in dtfi.MonthNames)
            {
                if (index < 13)
                {
                    Console.WriteLine(index + ": " + item[0].ToString().ToUpperInvariant() + item.TrimStart(item[0]));
                    if (months.Keys.Contains(index) == false)
                    {
                        months.Add(index, item[0].ToString().ToUpperInvariant() + item.TrimStart(item[0]));
                        index++;
                    } else index++;
                }
                else break;
            }
        }

        static void YearsGoldenClient()
        { 
            
        }
    }
}
