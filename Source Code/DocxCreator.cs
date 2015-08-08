using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace FinAnParser
{
    class DocxCreator
    {
        private WordprocessingDocument myDoc = null;

        public void Create(HtmlParser parser, Calculator calculator)
        {
            // Создание Wordprocessing документа 
            try
            {
                myDoc =
                   WordprocessingDocument.Create(MainWindow.Folder + Path.DirectorySeparatorChar + parser.ResultDirName + Path.DirectorySeparatorChar + parser.ResultFileNameWordTable,
                                 WordprocessingDocumentType.Document);
            }
            catch
            {
                // Укорачивание наименование файла в 2 раза, если превышает допустимое кол-во символов
                int halfLenghtWordFile = parser.ResultFileNameWordTable.Length / 2;
                string tmpWordStr = parser.ResultFileNameWordTable.Substring(0, halfLenghtWordFile);

                myDoc =
                   WordprocessingDocument.Create(MainWindow.Folder + Path.DirectorySeparatorChar + parser.ResultDirName + Path.DirectorySeparatorChar + tmpWordStr + ".docx",
                                 WordprocessingDocumentType.Document);
            }

            // Добавление главной части документа 
            MainDocumentPart mainPart = myDoc.AddMainDocumentPart();

            // Создание дерева DOM для простого документа
            mainPart.Document = new Document();
            Body body = new Body();
            Table table = new Table();
            TableProperties tblPr = new TableProperties();
            TableBorders tblBorders = new TableBorders();
            tblBorders.TopBorder = new TopBorder();
            tblBorders.TopBorder.Val = new EnumValue<BorderValues>(BorderValues.Single);
            tblBorders.BottomBorder = new BottomBorder();
            tblBorders.BottomBorder.Val = new EnumValue<BorderValues>(BorderValues.Single);
            tblBorders.LeftBorder = new LeftBorder();
            tblBorders.LeftBorder.Val = new EnumValue<BorderValues>(BorderValues.Single);
            tblBorders.RightBorder = new RightBorder();
            tblBorders.RightBorder.Val = new EnumValue<BorderValues>(BorderValues.Single);
            tblBorders.InsideHorizontalBorder = new InsideHorizontalBorder();
            tblBorders.InsideHorizontalBorder.Val = BorderValues.Single;
            tblBorders.InsideVerticalBorder = new InsideVerticalBorder();
            tblBorders.InsideVerticalBorder.Val = BorderValues.Single;
            tblPr.Append(tblBorders);
            table.Append(tblPr);
            // Первый ряд с наименованием файла
            TableRow tr;
            TableCell tc;
            tr = new TableRow();
            tc = new TableCell(new Paragraph(new Run(
                               new Text(parser.ResultFileNameWordTable.Replace(".docx", "")))));
            TableCellProperties tcp = new TableCellProperties();
            GridSpan gridSpan = new GridSpan();
            gridSpan.Val = 5;
            tcp.Append(gridSpan);
            tc.Append(tcp);
            tr.Append(tc);
            table.Append(tr);
            // Последующие ряды
            tr = new TableRow();
            tc = new TableCell();
            for (int i = 1; i <= 5; i++)
            {
                switch (i)
                {
                    case 1:
                        tr.Append(new TableCell(new Paragraph(new Run(new Text("№")))));
                        break;
                    case 2:
                        tr.Append(new TableCell(new Paragraph(new Run(new Text("Показатель")))));
                        break;
                    case 3:
                        tr.Append(new TableCell(new Paragraph(new Run(new Text("Взвешенная оценка")))));
                        break;
                    case 4:
                        tr.Append(new TableCell(new Paragraph(new Run(new Text("Оценка")))));
                        break;
                    case 5:
                        tr.Append(new TableCell(new Paragraph(new Run(new Text("Вес показателя")))));
                        break;
                }
            }
            table.Append(tr);
            int tmpInt = 1;
            for (int i = 1; i <= 9; i++)
            {
                tr = new TableRow();
                if (i <= 7)
                {
                    tc = new TableCell(new Paragraph(new Run(new Text(i.ToString()))));
                }
                else
                {
                    tc = new TableCell(new Paragraph(new Run(new Text(""))));
                }
                tr.Append(tc);
                for (int j = 1; j <= 4; j++)
                {
                    switch (tmpInt)
                    {
                        case 1:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text("Изменение выручки")))));
                            break;
                        case 2:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text(calculator.VzvesheniyItogIzmeneniyaViruchki.ToString())))));
                            break;
                        case 3:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text(calculator.ItogIzmeneniyaViruchki.ToString())))));
                            break;
                        case 4:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text("0,2")))));
                            break;
                        case 5:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text("Коэффициент текущей ликвидности")))));
                            break;
                        case 6:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text(calculator.VzvesheniyItogKoeficientaTekucsheyLikvidnosti.ToString())))));
                            break;
                        case 7:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text(calculator.ItogKoeficientaTekucsheyLikvidnosti.ToString())))));
                            break;
                        case 8:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text("0,2")))));
                            break;
                        case 9:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text("Коэффициент финансовой независимости")))));
                            break;
                        case 10:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text(calculator.VzvesheniyItogKoeficientaFinNezavisimosti.ToString())))));
                            break;
                        case 11:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text(calculator.ItogKoeficientaFinNezavisimosti.ToString())))));
                            break;
                        case 12:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text("0,2")))));
                            break;
                        case 13:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text("Чистые активы")))));
                            break;
                        case 14:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text(calculator.VzvesheniyItogChistihActivov.ToString())))));
                            break;
                        case 15:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text(calculator.ItogChistihActivov.ToString())))));
                            break;
                        case 16:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text("0,2")))));
                            break;
                        case 17:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text("Рентабельность продаж")))));
                            break;
                        case 18:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text(calculator.VzvesheniyItogRentabelnostiProdag.ToString())))));
                            break;
                        case 19:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text(calculator.ItogRentabelnostiProdag.ToString())))));
                            break;
                        case 20:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text("0,05")))));
                            break;
                        case 21:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text("Рентабельность деятельности")))));
                            break;
                        case 22:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text(calculator.VzvesheniyItogRentabelnostiDeyatelnosti.ToString())))));
                            break;
                        case 23:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text(calculator.ItogRentabelnostiDeyatelnosti.ToString())))));
                            break;
                        case 24:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text("0,15")))));
                            break;
                        case 25:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text("Коэффициент оборачиваемости (динамика)")))));
                            break;
                        case 26:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text()))));
                            break;
                        case 27:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text()))));
                            break;
                        case 28:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text("0,1")))));
                            break;
                        case 29:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text("Итого: ")))));
                            break;
                        case 30:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text()))));
                            break;
                        case 31:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text()))));
                            break;
                        case 32:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text()))));
                            break;
                        case 33:
                            tr.Append(new TableCell(new Paragraph(new Run(new Text("Оценка финансового положения заемщика")))));
                            break;
                        case 34:
                            tc = new TableCell(new Paragraph(new Run(
                               new Text(""))));
                            tcp = new TableCellProperties();
                            gridSpan = new GridSpan();
                            gridSpan.Val = 3;
                            tcp.Append(gridSpan);
                            tc.Append(tcp);
                            tr.Append(tc);
                            break;
                    }

                    tmpInt++;
                }
                table.Append(tr);
            }
            // Добавления таблицы к основе
            body.Append(table);
            // Добавление основы к документу
            mainPart.Document.Append(body);
            // Сохрание изменений главной части документа
            mainPart.Document.Save();
            myDoc.Dispose();
        }
    }
}
