

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
            if (manager.DB == null) //Обработка исключения Файл используется другим процессом. p.s. изменить, чтобы тебя за такое не побили
            {
                Console.WriteLine("Возникло исключение: Файл занят другим процессом, освободите файл и повторите попытку");
                goto start;
            }
            MenuOptions();
            Console.ReadKey();
        }

        #region Методы получения пути и отображения меню
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
        #endregion

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

        #region 3 Пункт меню
        /// <summary>
        /// Метод вывода меню выбора периода, за который мы ищем золотого клиента
        /// </summary>
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
                case "2": { YearsGoldenClient(); break; }
                default:
                    {
                        Console.WriteLine("\nВ меню нет пункта " + userResult + "\nПроверьте правильность ввода"); //Обработка случая,
                                                                                                                   //когда введены несуществующие пункты меню
                        goto userChoose;
                    }
            }
            #endregion
        }

        /// <summary>
        /// Метод выбора месяца и показ информации по этому месяцу
        /// </summary>
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
            #region Проверки
            userChoose:
                Console.Write("Месяц: ");
                string selectedMonth = Console.ReadLine();
                if (string.IsNullOrEmpty(selectedMonth) || selectedMonth.Length > 2) //Проверка на пустую строку или значения, которые больше двух символов
                {
                    Console.WriteLine("\nПожалуйста напишите одну цифру необходимого пункта меню");
                    goto userChoose;
                }
            if (!int.TryParse(selectedMonth, out int number)) //Проверка на введение буквы
            {
                Console.WriteLine("\nВы ввели букву, введите цифру пункта меню");
                goto userChoose;
            }
            else if (months.Keys.Contains(int.Parse(selectedMonth)) == false)
            {
                Console.WriteLine("\nВ меню нет пункта " + selectedMonth + "\nПроверьте правильность ввода"); //Обработка случая,
                                                                                                              //когда введены несуществующие пункты меню
                goto userChoose;
            }
            #endregion
            List<string> tempArray = new List<string>();
            foreach (var item in manager.DB[5])
            {
                string[] tempString = item.Value.ToString().Split(',');
                if (DateTime.Parse(tempString[4]).Month == int.Parse(selectedMonth))
                {
                    tempArray.Add(tempString[1]); // код клиента
                }
            }
            //Если в месяце был один заказ
            if (tempArray.Count == 1)
            {
                foreach (var item in tempArray)
                {
                    Console.WriteLine($"\"Золотой\" клиент на {months[int.Parse(selectedMonth)]} месяц - " +
                    $"{manager.DB[2][item]}\nКонтактное лицо организации - {manager.DB[4][item]}" +
                    $"\n\nВыбрать другой месяц или вернуться в меню?\n1 - Выбрать другой месяц\n2 - Вернуться в меню");
                #region Проверки
                userChooseSecondTime:
                    string userResult = Console.ReadLine().Trim();
                    if (string.IsNullOrEmpty(userResult) || userResult.Length > 1) //Проверка на пустую строку или значения, которые больше одного символа
                    {
                        Console.WriteLine("\nПожалуйста напишите одну цифру необходимого пункта меню");
                        goto userChooseSecondTime;
                    }
                    if (!int.TryParse(userResult, out int secnumber)) //Проверка на введение буквы
                    {
                        Console.WriteLine("\nВы ввели букву, введите цифру пункта меню");
                        goto userChooseSecondTime;
                    }
                    switch (userResult)
                    {
                        case "1": { MonthGoldenClient(); break; }
                        case "2": { MenuOptions(); break; }
                        default:
                            {
                                Console.WriteLine("\nВ меню нет пункта " + userResult + "\nПроверьте правильность ввода"); //Обработка случая,
                                                                                                                           //когда введены несуществующие пункты меню
                                goto userChooseSecondTime;
                            }
                    }
                    #endregion
                }
            }
            //Если заказов в выбраном месяце не было
            else if (tempArray.Count == 0)
            {
                Console.WriteLine($"За выбранный месяц - {months[int.Parse(selectedMonth)]} не было заказов" +
                    $"\n\nВыбрать другой месяц или вернуться в меню?\n1 - Выбрать другой месяц\n2 - Вернуться в меню");
            #region Проверки
            userChooseSecondTime:
                string userResult = Console.ReadLine().Trim();
                if (string.IsNullOrEmpty(userResult) || userResult.Length > 1) //Проверка на пустую строку или значения, которые больше одного символа
                {
                    Console.WriteLine("\nПожалуйста напишите одну цифру необходимого пункта меню");
                    goto userChooseSecondTime;
                }
                if (!int.TryParse(userResult, out int secnumber)) //Проверка на введение буквы
                {
                    Console.WriteLine("\nВы ввели букву, введите цифру пункта меню");
                    goto userChooseSecondTime;
                }
                switch (userResult)
                {
                    case "1": { MonthGoldenClient(); break; }
                    case "2": { MenuOptions(); break; }
                    default:
                        {
                            Console.WriteLine("\nВ меню нет пункта " + userResult + "\nПроверьте правильность ввода"); //Обработка случая,
                                                                                                                       //когда введены несуществующие пункты меню
                            goto userChooseSecondTime;
                        }
                }
                #endregion
            }
            //Если в месяце заказов больше 1
            else if (tempArray.Count > 1)
            {
                string tempString = "";
                index = 0;
                Dictionary<string, int> differents = new Dictionary<string, int>();
                foreach (var item in tempArray) //Перебираем выбранных клиентов
                {
                    if (index == 0) //Если это первая итерация заносим 1ого клиента во времнный словарь и указываем кол-во заказов - 1
                    {
                        differents.Add(item, 1);
                        tempString += item;
                        index++;
                    }
                    else
                    {
                        if (differents.Keys.Contains(item)) //Если в словаре уже есть ключ(код клиента) увеличиваем кол-во заказов на 1
                        {
                            differents[item] += 1;
                        }
                        else
                        {
                            differents.Add(item, 1); //Если нет, добавляем нового и его кол-во заказов увеличиваем до 1
                        }
                    }
                }
                var max = differents.MaxBy(kvp => kvp.Value).Key; //Определяем макмимальное значение value и передаём переменной max ключ(код клиента)
                Console.WriteLine($"\"Золотой\" клиент на {months[int.Parse(selectedMonth)]} месяц - " +
                    $"{manager.DB[2][max]}\nКонтактное лицо организации - {manager.DB[4][max]}"+
                    $"\n\nВыбрать другой месяц или вернуться в меню?\n1 - Выбрать другой месяц\n2 - Вернуться в меню");
            #region Проверки
            userChooseSecondTime:
                string userResult = Console.ReadLine().Trim();
                if (string.IsNullOrEmpty(userResult) || userResult.Length > 1) //Проверка на пустую строку или значения, которые больше одного символа
                {
                    Console.WriteLine("\nПожалуйста напишите одну цифру необходимого пункта меню");
                    goto userChooseSecondTime;
                }
                if (!int.TryParse(userResult, out int secnumber)) //Проверка на введение буквы
                {
                    Console.WriteLine("\nВы ввели букву, введите цифру пункта меню");
                    goto userChooseSecondTime;
                }
                switch (userResult)
                {
                    case "1": { MonthGoldenClient(); break; }
                    case "2": { MenuOptions(); break; }
                    default:
                        {
                            Console.WriteLine("\nВ меню нет пункта " + userResult + "\nПроверьте правильность ввода"); //Обработка случая,
                                                                                                                       //когда введены несуществующие пункты меню
                            goto userChooseSecondTime;
                        }
                }
                #endregion
            }
        }

        /// <summary>
        /// Метод выбора года и показ информации по этому году, если год один показывает информацию без выбора
        /// </summary>
        static void YearsGoldenClient()
        {
            Console.Clear();
            List<string> yearsArray = new List<string>();
            List<string> tempCodeClients = new List<string>();
            byte index = 0;
            foreach (var item in manager.DB[5]) //Перебираем выбранные года
            {
                string[] tempString = item.Value.ToString().Split(',');
                tempCodeClients.Add(tempString[1]);
                if (index == 0) //Если это первая итерация заносим год во времнный лист
                {
                    yearsArray.Add(DateTime.Parse(tempString[4]).Year.ToString()); // год
                    index++;
                }
                else if (yearsArray.Contains(DateTime.Parse(tempString[4]).Year.ToString())) //Если указанный год уже существует в листе, продолжаем перебор
                {
                    continue;
                }
                else yearsArray.Add(DateTime.Parse(tempString[4]).Year.ToString());
            }
            //Если в файле не было заказов за несколько лет, выводит золотого клиента за один год
            if (yearsArray.Count == 1) 
            {
                string tempString = "";
                index = 0;
                Dictionary<string, int> differents = new Dictionary<string, int>();
                foreach (var item in tempCodeClients) //Перебираем выбранных клиентов
                {
                    if (index == 0) //Если это первая итерация заносим 1ого клиента во временный словарь и указываем кол-во заказов - 1
                    {
                        differents.Add(item, 1);
                        tempString += item;
                        index++;
                    }
                    else
                    {
                        if (differents.Keys.Contains(item)) //Если в словаре уже есть ключ(код клиента) увеличиваем кол-во заказов на 1
                        {
                            differents[item] += 1;
                        }
                        else
                        {
                            differents.Add(item, 1); //Если нет, добавляем нового и его кол-во заказов увеличиваем до 1
                        }
                    }
                }
                var max = differents.MaxBy(kvp => kvp.Value).Key; //Определяем макмимальное значение value и передаём переменной max ключ(код клиента)
                Console.WriteLine($"\nВ выбранном вами файле был найден только один год\n\n\"Золотой\" клиент на {yearsArray.First()} год - " +
                    $"{manager.DB[2][max]}\nКонтактное лицо организации - {manager.DB[4][max]}" +
                    $"\nНажмите любую клавишу для возврата в меню");
                Console.ReadKey();
                MenuOptions();
            }
            //Если в файле присутсвуют данные за разные года, выводит выбор года
            else 
            {
                Dictionary<int,string> numRateYears = new Dictionary<int,string>();

                int secindex = 1;
                Console.WriteLine("Для какого года определяем \"золотого\" клиента?");
                foreach (var item in yearsArray) //Формируем словарь для выбора
                {
                    Console.WriteLine(secindex + ": " + item);
                    numRateYears.Add(secindex, item);
                    secindex++;
                }
            #region Проверки
            userChoose:
                Console.Write("Год: ");
                string selectedYear = Console.ReadLine();
                if (string.IsNullOrEmpty(selectedYear) || selectedYear.Length > 1) //Проверка на пустую строку или значения, которые больше двух символов
                {
                    Console.WriteLine("\nПожалуйста напишите одну цифру необходимого пункта меню");
                    goto userChoose;
                }
                if (!int.TryParse(selectedYear, out int number)) //Проверка на введение буквы
                {
                    Console.WriteLine("\nВы ввели букву, введите цифру пункта меню");
                    goto userChoose;
                }
                else if (numRateYears.Keys.Contains(int.Parse(selectedYear)) == false)
                {
                    Console.WriteLine("\nВ меню нет пункта " + selectedYear + "\nПроверьте правильность ввода"); //Обработка случая,
                                                                                                                  //когда введены несуществующие пункты меню
                    goto userChoose;
                }
                #endregion
                List<string> tempArray = new List<string>();
                foreach (var item in manager.DB[5])
                {
                    string[] tempString = item.Value.ToString().Split(',');
                    if (DateTime.Parse(tempString[4]).Year == int.Parse(numRateYears[int.Parse(selectedYear)]))
                    {
                        tempArray.Add(tempString[1]); // код клиента
                    }
                }
                string tempStringForDif = "";
                index = 0;
                Dictionary<string, int> differents = new Dictionary<string, int>();
                foreach (var item in tempArray) //Перебираем выбранных клиентов
                {
                    if (index == 0) //Если это первая итерация заносим 1ого клиента во временный словарь и указываем кол-во заказов - 1
                    {
                        differents.Add(item, 1);
                        tempStringForDif += item;
                        index++;
                    }
                    else
                    {
                        if (differents.Keys.Contains(item)) //Если в словаре уже есть ключ(код клиента) увеличиваем кол-во заказов на 1
                        {
                            differents[item] += 1;
                        }
                        else
                        {
                            differents.Add(item, 1); //Если нет, добавляем нового и его кол-во заказов увеличиваем до 1
                        }
                    }
                }
                var max = differents.MaxBy(kvp => kvp.Value).Key; //Определяем макмимальное значение value и передаём переменной max ключ(код клиента)
                Console.WriteLine($"\n\"Золотой\" клиент на {int.Parse(numRateYears[int.Parse(selectedYear)])} год - " +
                    $"{manager.DB[2][max]}\nКонтактное лицо организации - {manager.DB[4][max]}" +
                    $"\nНажмите любую клавишу для возврата в меню");
                Console.ReadKey();
                MenuOptions();
            }
        }
        #endregion
    }
}
