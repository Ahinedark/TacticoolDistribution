using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TacticoolDistribution
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;

            //Считывание таблицы
            List<List<string>> maping = new List<List<string>>();
            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo("tacticoolDB.xlsx")))
            {
                var myWorksheet = xlPackage.Workbook.Worksheets.First();
                var totalRows = myWorksheet.Dimension.End.Row;
                var totalColumns = myWorksheet.Dimension.End.Column;

                var sb = new StringBuilder();
                for (int rowNum = 2; rowNum <= totalRows; rowNum++)
                {
                    var row = myWorksheet.Cells[rowNum, 2, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());
                    sb.AppendLine(string.Join(",", row));
                }

                char[] seps = new char[] { ',', '\n' };
                string[] tempAr = sb.ToString().Split(seps[1]);
                foreach (string a in tempAr)
                {
                    maping.Add(new List<string>(a.Split(seps[0])));
                }
                maping.RemoveAt(maping.Count - 1);
            }

            //Информация и начало работы
            Console.WriteLine("Распределение оперативников Tacticool. (c) Гневный, 2019 год.\n");
            Console.WriteLine("1 - РАЗБОРКА; 2 - HILDR; 3 - B.S.S.; 4 - ШТЫК;");
            Console.WriteLine("5 - МОЛОТ; 6 - НОЖ; 7 - МЕСТНЫЕ; 8 - ДОСТАВКА;");
            Console.WriteLine("9 - РАЗВЕДКА; 10 - ШТУРМ; 11 - ПРИКРЫТИЕ; 12 - ЗАЧИСТКА;");
            Console.WriteLine("13 - БАЗОВОЕ ЗАДАНИЕ; 14 - ТОЛЬКО ОБЫЧНЫЕ;");
            Console.WriteLine("15 - ТОЛЬКО НЕОБЫЧНЫЕ; 16 - ТОЛЬКО РЕДКИЕ.");
            Console.WriteLine("\nВведите номер задания, чтобы добавить его в список распределения, \nпосле чего нажмите Enter. Таким же образом введите номера всех остальных \nзаданий по порядку. Для окончания ввода, просто нажмите Enter: ");

            //Считывание номеров заданий. Если пришла пустая строка - закончить
            int mIndex = 1;
            string temp;
            List<int> mIndexList = new List<int>();
            while (true)
            {
                temp = Console.ReadLine();
                if (temp == "")
                    break;
                mIndex = Convert.ToInt32(temp);
                mIndexList.Add(mIndex);
            }

            //Отбор нужных строк по введённым номерам
            List<List<string>> sample = new List<List<string>>();
            foreach (int i in mIndexList)
            {
                sample.Add(maping[i - 1]);
            }

            //Распределение оперативников по заданиям справа налево
            int max = 0;
            List<int> result = new List<int>();
            for (int i = 0; i < sample[0].Count; i++)
            {
                result.Add(0);
                for (int j = sample.Count - 1; j >= 0; j--)
                {
                    if (Convert.ToInt32(sample[j][i]) > max)
                    {
                        max = Convert.ToInt32(sample[j][i]);
                        result[i] = j;
                    }
                }
                max = 0;
            }

            //Формирование результатов работы
            List<string> resultStrings = new List<string>();
            string header = "Распределение оперативников по заданиям";
            resultStrings.Add(header);
            for (int i = 0; i < mIndexList.Count; i++)
            {
                string resultText = (i + 1) + " задание - ";
                switch (mIndexList[i])
                {
                    case 1: resultText += "РАЗБОРКА: "; break;
                    case 2: resultText += "HILDR: "; break;
                    case 3: resultText += "B.S.S.: "; break;
                    case 4: resultText += "ШТЫК: "; break;
                    case 5: resultText += "МОЛОТ: "; break;
                    case 6: resultText += "НОЖ: "; break;
                    case 7: resultText += "МЕСТНЫЕ: "; break;
                    case 8: resultText += "ДОСТАВКА: "; break;
                    case 9: resultText += "РАЗВЕДКА: "; break;
                    case 10: resultText += "ШТУРМ: "; break;
                    case 11: resultText += "ПРИКРЫТИЕ: "; break;
                    case 12: resultText += "ЗАЧИСТКА: "; break;
                    case 13: resultText += "БАЗОВОЕ ЗАДАНИЕ: "; break;
                    case 14: resultText += "ТОЛЬКО ОБЫЧНЫЕ: "; break;
                    case 15: resultText += "ТОЛЬКО НЕОБЫЧНЫЕ: "; break;
                    case 16: resultText += "ТОЛЬКО РЕДКИЕ: "; break;
                }

                for (int j = 0; j < result.Count; j++)
                {
                    if (result[j] == i)
                    {
                        switch (j + 1)
                        {
                            case 1: resultText += "Новобранец"; break;
                            case 2: resultText += "Рик"; break;
                            case 3: resultText += "Борис"; break;
                            case 4: resultText += "Тор"; break;
                            case 5: resultText += "Мишка"; break;
                            case 6: resultText += "Ястреб"; break;
                            case 7: resultText += "Джейсон"; break;
                            case 8: resultText += "Трэвис"; break;
                            case 9: resultText += "Виктор"; break;
                            case 10: resultText += "Спенсер"; break;
                            case 11: resultText += "Батя"; break;
                            case 12: resultText += "Клаус"; break;
                            case 13: resultText += "Ши"; break;
                            case 14: resultText += "Валера"; break;
                            case 15: resultText += "Джо"; break;
                            case 16: resultText += "Варг"; break;
                            case 17: resultText += "Синдром"; break;
                            case 18: resultText += "Дерзкий"; break;
                            case 19: resultText += "Дэвид"; break;
                            case 20: resultText += "Злой"; break;
                            case 21: resultText += "Снэк"; break;
                            case 22: resultText += "Датч"; break;
                            case 23: resultText += "Диана"; break;
                        }
                        if (j != result.Count - 1)
                            resultText += ", ";
                        else
                            resultText += ".";
                    }
                }
                resultStrings.Add(resultText);
            }

            //Запись результатов в файл
            string writePath = @"resultMessage.txt";
            try
            {
                using (StreamWriter sw = new StreamWriter(writePath, false, System.Text.Encoding.Default))
                {
                    foreach (string s in resultStrings)
                    {
                        sw.WriteLine(s);
                    }
                }
                Console.WriteLine("Запись выполнена");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            Console.ReadKey();
        }
    }
}
