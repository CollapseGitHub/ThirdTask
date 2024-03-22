using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection.Metadata.Ecma335;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace ThirdTask
{
    internal class XlsxManagement : XLWorkbook
    {
        private string pathOfFile = ""; //Поле хранения пути до файла

        public string PathOfFile
        { get {return pathOfFile; }
          set {pathOfFile = value; }
        }

        private List<Dictionary<string, string>> db = new List<Dictionary<string, string>>(); //Поле хранения "БД"

        public List<Dictionary<string, string>> DB
        { get { return db; }
          set { db = value; }
        } 

        /// <summary>
        /// Метод создания "БД" проекта, 
        /// </summary>
        /// <param name="path">Путь до файла</param>
        /// <returns>Заполненная "БД"</returns>
        public List<Dictionary<string, string>>? GetInfoFromDoc(string path)
        {
            try
            {
                using (XLWorkbook workbook = new XLWorkbook(path))
                {
                    List<Dictionary<string, string>> db = new List<Dictionary<string, string>>(); //Общая БД экземпляра
                    #region БД продуктов
                    Dictionary<string, string> productsOnly = new Dictionary<string, string>(); //БД наименования продуктов,где keys - код товара, values - наименование товара
                    Dictionary<string, string> productPrices = new Dictionary<string, string>(); //БД цены,где keys - код товара, values - цена товара
                    #endregion
                    #region БД клиентов
                    Dictionary<string, string> clientsOrgName = new Dictionary<string, string>(); //БД наименования организации клиента, где keys - код клиента, values - наименование организаций
                    Dictionary<string, string> clientsAdress = new Dictionary<string, string>(); //БД адреса организации, где keys - код клиента, values - адрес организации
                    Dictionary<string, string> clientsFIO = new Dictionary<string, string>(); //БД ФИО рук-ля организации, где keys - код клиента, values - ФИО контактного лица
                    #endregion
                    #region БД заявок
                    Dictionary<string, string> requestsInfo = new Dictionary<string, string>(); //БД из листа заявок, keys - код заявки, value - string массив с остальными данными 
                    #endregion

                    foreach (var item in workbook.Worksheet(1).RangeUsed().Rows(2, 20))
                    {
                        productsOnly.Add(item.Cell(1).Value.ToString(), item.Cell(2).Value.ToString());
                        productPrices.Add(item.Cell(1).Value.ToString(), item.Cell(4).Value.ToString());
                    }
                    foreach (var item in workbook.Worksheet(2).RangeUsed().Rows(2, 5))
                    {
                        clientsOrgName.Add(item.Cell(1).Value.ToString(), item.Cell(2).Value.ToString());
                        clientsAdress.Add(item.Cell(1).Value.ToString(), item.Cell(3).Value.ToString());
                        clientsFIO.Add(item.Cell(1).Value.ToString(), item.Cell(4).Value.ToString());
                    }
                    foreach (var item in workbook.Worksheet(3).RangeUsed().Rows(2, 9))
                    {
                        requestsInfo.Add(item.Cell(1).Value.ToString(), item.Cell(2).Value.ToString() + ',' + item.Cell(3).Value.ToString() + ',' +
                            item.Cell(4).Value.ToString() + ',' + item.Cell(5).Value.ToString() + ',' + item.Cell(6).Value.ToString());
                    }
                    db.Add(productsOnly);
                    db.Add(productPrices);
                    db.Add(clientsOrgName);
                    db.Add(clientsAdress);
                    db.Add(clientsFIO);
                    db.Add(requestsInfo);
                    return db;
                }
            }
            catch (Exception)
            {
                return null;
                throw;
            }
            
        }

        /// <summary>
        /// Метод замены и сохранения новых данных по контактному лицу
        /// </summary>
        /// <param name="path">Путь до файла</param>
        /// <param name="selectedClient">Код клиента у которого необходимо заменить контактное лицо</param>
        /// <param name="FIO">ФИО на которое необходимо заменить предыдущее лицо</param>
        /// <returns></returns>
        public string SetFIOClient(string path, string selectedClient, string FIO)
        {
            string tempString = "";
                using (XLWorkbook workbook = new XLWorkbook(path))
                {
                    foreach (var item in workbook.Worksheet(2).RangeUsed().Rows(2,5))
                    {
                        if (item.Cell(1).Value.ToString() == selectedClient)
                        {
                            tempString = item.Cell(4).Value.ToString();
                            item.Cell(4).Value = FIO; break;
                        }
                    }
                    workbook.Save(); //Сохраняем измененные данные в файле xlsx
                    return $"{tempString} на {FIO}.";
                }
        }
    }
}
