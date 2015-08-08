using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace FinAnParser
{
    class HtmlParser
    {
        public string Zagolovok = "";
        public string Okved = "";
        public string OkvedNomer = "";
        public string OkvedVidDeyatelnosti = "";
        public double SobstveniyKapital = 0;
        public double ValutaBalansa = 0;
        public double ChistieAktiviNachalo = 0;
        public double ChistieAktiviKonec = 0;
        public double ChistieAktiviIsmeneniya = 0;
        public double KoeficTekucsheyObcsheyLikvidnosti = 0;
        public double Viruchka = 0.0;
        public double ViruchkaPlusMinusProcenti = 0;
        public double PribilUbitokProdag = 0;
        public double ChistayaPribilUbitok = 0;
        public double OborachivaemostDebitorskoyZadolgnosti = 0;
        public double OborachivaemostAktivov = 0;

        public double NomerPunkta = 0;

        public string ResultFileName = "";
        public string ResultFileNameWordTable = "";
        public string ResultFileNameWordTableExc = "";
        public string ResultDirName = "";

        private StringBuilder textBuilder = new StringBuilder();
        private Calculator calculator = new Calculator();

        public void Parse()
        {

            NomerPunkta = 1;

            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.Load(MainWindow.File);

            HtmlAgilityPack.HtmlNode root = doc.DocumentNode;

            HtmlAgilityPack.HtmlNode h1_node_zagolovok = root.Descendants("h1").First();

            Zagolovok = h1_node_zagolovok.InnerText.Trim();

            textBuilder.Append(Zagolovok + Environment.NewLine + Environment.NewLine);

            // Выбор ОКВЕД
            HtmlAgilityPack.HtmlNode p_node_okved = root.Descendants("p").First();
            string okvedRawString = p_node_okved.InnerText.Trim();

            Regex okvedRegExpr = new Regex(@"(?s)(?<=ОКВЭД).+?(?=\))");
            MatchCollection matches = okvedRegExpr.Matches(okvedRawString);
            foreach (Match mat in matches)
            {
                Okved = mat.Value;
            }
            Okved = Regex.Replace(Okved, "[^.0-9]", "");

            OkvedType();

            textBuilder.Append("ОКВЭД: " + OkvedNomer + Environment.NewLine);
            textBuilder.Append("Вид деятельности по ОКВЭД'у: " + OkvedVidDeyatelnosti + Environment.NewLine + Environment.NewLine);

            List<HtmlAgilityPack.HtmlNode> tr_nodes = root.Descendants("tr").ToList();

            foreach (HtmlAgilityPack.HtmlNode tr_node in tr_nodes)
            {
                List<HtmlAgilityPack.HtmlNode> td_nodes = tr_node.Descendants("td").ToList();

                foreach (HtmlAgilityPack.HtmlNode td_node in td_nodes)
                {
                    string rawString = td_node.InnerText.Trim();
                    StringReader stringReader = new StringReader(rawString);
                    string str = "";
                    string tempString = "";

                    str = stringReader.ReadLine();
                    if (!string.IsNullOrEmpty(str))
                    {
                        tempString = str.Trim().Replace("&nbsp;", "").Replace(" ", string.Empty);
                        tempString = Regex.Replace(tempString, @"\s+", "");
                    }

                    try
                    {
                        switch (tempString)
                        {
                            case "1.Собственныйкапитал":
                                {
                                    HtmlAgilityPack.HtmlNode td_tempNode = td_nodes[2];
                                    SobstveniyKapital = Convert.ToDouble(td_tempNode.InnerText.Trim().Replace("&nbsp;", ""), CultureInfo.InvariantCulture);
                                    textBuilder.Append(NomerPunkta + ") Собственный капитал 9 месяцев: " + SobstveniyKapital + Environment.NewLine);
                                    ++NomerPunkta;
                                    break;
                                }
                            case "Валютабаланса":
                                {
                                    HtmlAgilityPack.HtmlNode td_tempNode = td_nodes[2];
                                    ValutaBalansa = Convert.ToDouble(td_tempNode.InnerText.Trim().Replace("&nbsp;", ""), CultureInfo.InvariantCulture);
                                    textBuilder.Append(NomerPunkta + ") Валюта баланса 9 месяцев: " + ValutaBalansa.ToString() + Environment.NewLine);
                                    ++NomerPunkta;
                                    break;
                                }
                            case "1.Чистыеактивы":
                                {
                                    HtmlAgilityPack.HtmlNode td_tempNode1 = td_nodes[1];
                                    ChistieAktiviNachalo = Convert.ToDouble(td_tempNode1.InnerText.Trim().Replace("&nbsp;", ""), CultureInfo.InvariantCulture);
                                    textBuilder.Append(NomerPunkta + ".1) Чистые активы на начало периода: " + ChistieAktiviNachalo.ToString() + Environment.NewLine);

                                    HtmlAgilityPack.HtmlNode td_tempNode2 = td_nodes[2];
                                    ChistieAktiviKonec = Convert.ToDouble(td_tempNode2.InnerText.Trim().Replace("&nbsp;", ""), CultureInfo.InvariantCulture);
                                    textBuilder.Append(NomerPunkta + ".2) Чистые активы на конец периода: " + ChistieAktiviKonec.ToString() + Environment.NewLine);

                                    HtmlAgilityPack.HtmlNode td_tempNode3 = td_nodes[5];
                                    ChistieAktiviIsmeneniya = Convert.ToDouble(td_tempNode3.InnerText.Trim().Replace("&nbsp;", ""), CultureInfo.InvariantCulture);
                                    textBuilder.Append(NomerPunkta + ".3) Чистые активы изменение гр.3-гр.2: " + ChistieAktiviIsmeneniya.ToString() + Environment.NewLine);
                                    ++NomerPunkta;
                                    break;
                                }
                            case "1.Коэффициенттекущей(общей)ликвидности":
                                {
                                    HtmlAgilityPack.HtmlNode td_tempNode = td_nodes[2];
                                    KoeficTekucsheyObcsheyLikvidnosti = Convert.ToDouble(td_tempNode.InnerText.Trim().Replace("&nbsp;", ""));
                                    textBuilder.Append(NomerPunkta + ") Коэффициент текущей (общей) ликвидности 9 месяцев: " + KoeficTekucsheyObcsheyLikvidnosti.ToString() + Environment.NewLine);
                                    ++NomerPunkta;
                                    break;
                                }

                            case "1.Выручка":
                                {
                                    HtmlAgilityPack.HtmlNode td_tempNode = td_nodes[2];
                                    Viruchka = Convert.ToDouble(td_tempNode.InnerText.Trim().Replace("&nbsp;", ""), CultureInfo.InvariantCulture);
                                    textBuilder.Append(NomerPunkta + ") Выручка 9 месяцев: " + Viruchka.ToString() + Environment.NewLine);
                                    ++NomerPunkta;

                                    HtmlAgilityPack.HtmlNode td_tempNode2 = td_nodes[4];
                                    ViruchkaPlusMinusProcenti = Convert.ToDouble(td_tempNode2.InnerText.Trim().Replace("&nbsp;", ""), CultureInfo.InvariantCulture);
                                    textBuilder.Append(NomerPunkta + ") Выручка +-%((3-2):2): " + ViruchkaPlusMinusProcenti.ToString() + Environment.NewLine);
                                    ++NomerPunkta;
                                    break;
                                }

                            case "3.Прибыль(убыток)отпродаж(1-2)":
                                {
                                    HtmlAgilityPack.HtmlNode td_tempNode = td_nodes[2];
                                    PribilUbitokProdag = Convert.ToDouble(td_tempNode.InnerText.Trim().Replace("&nbsp;", ""), CultureInfo.InvariantCulture);
                                    textBuilder.Append(NomerPunkta + ") Прибыль (убыток) от продаж  (1-2) 9 месяцев: " + PribilUbitokProdag.ToString() + Environment.NewLine);
                                    ++NomerPunkta;
                                    break;
                                }

                            case "8.Чистаяприбыль(убыток)(5-6+7)":
                                {
                                    HtmlAgilityPack.HtmlNode td_tempNode = td_nodes[2];
                                    ChistayaPribilUbitok = Convert.ToDouble(td_tempNode.InnerText.Trim().Replace("&nbsp;", ""), CultureInfo.InvariantCulture);
                                    textBuilder.Append(NomerPunkta + ") Чистая прибыль (убыток) (5-6+7) 9 месяцев: " + ChistayaPribilUbitok.ToString() + Environment.NewLine);
                                    ++NomerPunkta;
                                    break;
                                }

                            case "Оборачиваемостьзапасов":
                                {
                                    HtmlAgilityPack.HtmlNode td_tempNode = td_nodes[1];
                                    string oborachivaemostZapasov = td_tempNode.InnerText.Trim().Replace("&nbsp;", "").Replace("&lt;", "<").Replace("&gt", ">");
                                    textBuilder.Append(NomerPunkta + ") Оборачиваемость запасов 9 месяцев: " + oborachivaemostZapasov + Environment.NewLine);
                                    ++NomerPunkta;
                                    break;
                                }

                            case "Оборачиваемостьдебиторскойзадолженности":
                                {
                                    HtmlAgilityPack.HtmlNode td_tempNode = td_nodes[1];
                                    OborachivaemostDebitorskoyZadolgnosti = Convert.ToDouble(td_tempNode.InnerText.Trim().Replace("&nbsp;", ""), CultureInfo.InvariantCulture);
                                    textBuilder.Append(NomerPunkta + ") Оборачиваемость дебиторской задолженности 9 месяцев: " + OborachivaemostDebitorskoyZadolgnosti.ToString() + Environment.NewLine);
                                    ++NomerPunkta;
                                    break;
                                }

                            case "Оборачиваемостьактивов":
                                {
                                    HtmlAgilityPack.HtmlNode td_tempNode = td_nodes[1];
                                    OborachivaemostAktivov = Convert.ToDouble(td_tempNode.InnerText.Trim().Replace("&nbsp;", ""), CultureInfo.InvariantCulture);
                                    textBuilder.Append(NomerPunkta + ") Оборачиваемость активов 9 месяцев: " + OborachivaemostAktivov.ToString() + Environment.NewLine);
                                    ++NomerPunkta;
                                    break;
                                }
                        }
                    }
                    catch (Exception) { }
                }
            }

            StringReader zagolovokReader = new StringReader(Zagolovok);
            string zagolovokString = "";

            string strWhile = "";
            int i = 0;
            do
            {

                strWhile = zagolovokReader.ReadLine();


                if ((strWhile != null && i == 1) || (strWhile != null && i == 2))
                {
                    zagolovokString += strWhile.Replace("\"", " ");
                }

                ++i;
            } while (strWhile != null);

            calculator.PodschetIzmeneniyaViruchki(ViruchkaPlusMinusProcenti);
            calculator.PodschetKoeficientaTekucsheyLikvidnosti(KoeficTekucsheyObcsheyLikvidnosti);
            calculator.PodschetKoeficientaFinNezavisimosti(SobstveniyKapital, ValutaBalansa, OkvedVidDeyatelnosti);
            calculator.PodschetChistihActivov(ChistieAktiviNachalo, ChistieAktiviKonec, ChistieAktiviIsmeneniya);
            calculator.PodschetRentabelnostiProdag(PribilUbitokProdag, Viruchka);
            calculator.PodschetRentabelnostiDeyatelnosti(ViruchkaPlusMinusProcenti, Viruchka);

            textBuilder.Append(Environment.NewLine + Environment.NewLine);
            textBuilder.Append("Итоговая оценка заёмщика" + Environment.NewLine + Environment.NewLine);

            textBuilder.Append(@"1. Изменение выручки: " + string.Format("{0:0.00}", ViruchkaPlusMinusProcenti) + "%" + Environment.NewLine + "   Оценка: "
                + calculator.ItogIzmeneniyaViruchki.ToString() + Environment.NewLine
                + "   Взвешенная оценка: " + calculator.VzvesheniyItogIzmeneniyaViruchki.ToString() + Environment.NewLine + Environment.NewLine);

            textBuilder.Append(@"2. Коэффициент текущей ликвидности: " + string.Format("{0:0.00}", KoeficTekucsheyObcsheyLikvidnosti)
                + Environment.NewLine + "   Оценка: " + calculator.ItogKoeficientaTekucsheyLikvidnosti.ToString()
                + Environment.NewLine + "   Взвешенная оценка: " + calculator.VzvesheniyItogKoeficientaTekucsheyLikvidnosti.ToString() + Environment.NewLine + Environment.NewLine);

            textBuilder.Append(@"3. Коэффициент финансовой независимости: " + string.Format("{0:0.00}", calculator.ResultatKoeficientaFinNezavisimosti)
                + Environment.NewLine + "   Оценка: " + calculator.ItogKoeficientaFinNezavisimosti.ToString() + Environment.NewLine + "   Взвешенная оценка: "
                + calculator.VzvesheniyItogKoeficientaFinNezavisimosti.ToString() + Environment.NewLine + Environment.NewLine);

            if (calculator.ItogChistihActivov == 1)
            {
                textBuilder.Append(@"4. Чистые активы на начало " + ChistieAktiviNachalo + " меньше чем на конец " + ChistieAktiviKonec +
                    " периода и итоговые чистые активы положительные " + ChistieAktiviIsmeneniya
                    + Environment.NewLine + "   Оценка: " + calculator.ItogChistihActivov.ToString() + Environment.NewLine
                    + "   Взвешенная оценка: " + calculator.VzvesheniyItogChistihActivov.ToString() + Environment.NewLine + Environment.NewLine);
            }

            if (calculator.ItogChistihActivov == 0.8)
            {
                textBuilder.Append(@"4. Чистые активы на начало " + ChistieAktiviNachalo + " больше чем на конец " + ChistieAktiviKonec +
                    " периода и итоговые чистые активы положительные " + ChistieAktiviIsmeneniya
                    + Environment.NewLine + "   Оценка: " + calculator.ItogChistihActivov.ToString() + Environment.NewLine
                    + "   Взвешенная оценка: " + calculator.VzvesheniyItogChistihActivov.ToString() + Environment.NewLine + Environment.NewLine);
            }

            if (calculator.ItogChistihActivov == 0.6)
            {
                textBuilder.Append(@"4. Чистые активы на начало " + ChistieAktiviNachalo + " меньше чем на конец " + ChistieAktiviKonec +
                    " периода и итоговые чистые активы отрицательные " + ChistieAktiviIsmeneniya
                    + Environment.NewLine + "   Оценка: " + calculator.ItogChistihActivov.ToString() + Environment.NewLine
                    + "   Взвешенная оценка: " + calculator.VzvesheniyItogChistihActivov.ToString() + Environment.NewLine + Environment.NewLine);
            }

            if (calculator.ItogChistihActivov == 0.4)
            {
                textBuilder.Append(@"4. Чистые активы на начало " + ChistieAktiviNachalo + " больше чем на конец " + ChistieAktiviKonec +
                    " периода и итоговые чистые активы отрицательные " + ChistieAktiviIsmeneniya
                    + Environment.NewLine + "   Оценка: " + calculator.ItogChistihActivov.ToString() + Environment.NewLine + "   Взвешенная оценка: " + calculator.VzvesheniyItogChistihActivov.ToString() + Environment.NewLine + Environment.NewLine);
            }

            textBuilder.Append(@"5. Рентабельность продаж: " + string.Format("{0:0.00}", calculator.ResultatRentabelnostiProdag) + "%" + Environment.NewLine + "   Оценка: " + calculator.ItogRentabelnostiProdag.ToString() + Environment.NewLine + "   Взвешенная оценка: " + calculator.VzvesheniyItogRentabelnostiProdag.ToString() + Environment.NewLine + Environment.NewLine);
            textBuilder.Append(@"6. Рентабельность деятельности: " + string.Format("{0:0.00}", calculator.ResultatRentabelnostiDeyatelnosti) + "%" + Environment.NewLine + "   Оценка: " + calculator.ItogRentabelnostiDeyatelnosti.ToString() + Environment.NewLine + "   Взвешенная оценка: " + calculator.VzvesheniyItogRentabelnostiDeyatelnosti.ToString() + Environment.NewLine + Environment.NewLine);

            ResultFileName = "Данные для оценки заёмщика из отчёта полного фин. анализа " + zagolovokString + ".txt";
            ResultFileNameWordTable = "Итоговая таблица оценки заёмщика " + zagolovokString + ".docx";
            ResultFileNameWordTableExc = "Итоговая таблица оценки заёмщика " + zagolovokString;
            ResultDirName = "Обработанные данные из отчёта полного фин. анализа " + zagolovokString;

            try
            {
                DirectoryInfo di = Directory.CreateDirectory(MainWindow.Folder + Path.DirectorySeparatorChar + ResultDirName);
            }
            catch (Exception)
            {
                int halfLenghtDir = ResultDirName.Length / 2;
                string tmpZagStr = ResultDirName.Substring(0, halfLenghtDir);
                DirectoryInfo di = Directory.CreateDirectory(MainWindow.Folder + Path.DirectorySeparatorChar + tmpZagStr);
            }

            try
            {
                File.WriteAllText(MainWindow.Folder + Path.DirectorySeparatorChar + ResultDirName + Path.DirectorySeparatorChar + ResultFileName, textBuilder.ToString(), Encoding.Default);
            }
            catch (Exception)
            {
                int halfLenghtFile = ResultFileName.Length / 2;
                string tmpFileStr = ResultFileName.Substring(0, halfLenghtFile);
                File.WriteAllText(MainWindow.Folder + Path.DirectorySeparatorChar + ResultDirName + Path.DirectorySeparatorChar + tmpFileStr, textBuilder.ToString(), Encoding.Default);
            }

            textBuilder = new StringBuilder();         
            new DocxCreator().Create(this, calculator); // Создание файла .docx с таблицей
            MainWindow.tbResult.Text = "OK! Папка с файлами \"" + ResultFileNameWordTableExc + "\"" + " cоздана!";
        }

        public void OkvedType()
        {
            switch (Okved)
            {
                case "12": OkvedNomer = "Сельское хозяйство, охота и лесное хозяйство (1,2)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "5": OkvedNomer = "Рыболовство, рыбоводство (5)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "1012": OkvedNomer = "Добыча топливно-энергетических полезных ископаемых (10-12)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "1314": OkvedNomer = "Добыча полезных ископаемых, кроме топливно-энергетических (13,14)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "1516": OkvedNomer = "Производство пищевых продуктов, включая напитки, и табака (15,16)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "1718": OkvedNomer = "Текстильное и швейное производство (17,18)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "19": OkvedNomer = "Производство кожи, изделий из кожи и производство обуви (19)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "20": OkvedNomer = "Обработка древесины и производство изделий из дерева (20)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "2122": OkvedNomer = "Целлюлозно-бумажное производство: издательская и полтграфическая деательность (21,22)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "23": OkvedNomer = "Производство кокса, нефтепродуктов и ядерных материалов (23)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "24": OkvedNomer = "Химическое производство (24)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "25": OkvedNomer = "Производство резиновых и пластмассовых изделий (25)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "26": OkvedNomer = "Производство прочих неметаллических минеральных продуктов (26)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "2728": OkvedNomer = "Металлургия, производство металлических изделий (27,28)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "29": OkvedNomer = "Производство машин и оборудования (29)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "3033": OkvedNomer = "Производство электро- и оптического оборудования (30-33)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "3435": OkvedNomer = "Производство транспортных средств и оборудования (34,35)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "3637": OkvedNomer = "Обрабатывающие производства; Прочие производства (36,37)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "4041": OkvedNomer = "Производство и распределение электроэнергии, газа и воды (40,41)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "45": OkvedNomer = "Сторительство (45)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "50": OkvedNomer = "Торговля авто и мотоциклами, их техобслуживание и ремонт (50)"; OkvedVidDeyatelnosti = "Торговля"; break;
                case "51": OkvedNomer = "Оптовая торговля, включая торговлю через агентов (51)"; OkvedVidDeyatelnosti = "Торговля"; break;
                case "52": OkvedNomer = "Розничная торговля; ремонт бытовых и личных изделий (52)"; OkvedVidDeyatelnosti = "Торговля"; break;
                case "55": OkvedNomer = "Гостиницы и ретсораны (55)"; OkvedVidDeyatelnosti = "Торговля"; break;
                case "6063": OkvedNomer = "Транспорт (60-63)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "64": OkvedNomer = "Связь (64)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "6567": OkvedNomer = "Финансовая деятельность (65-67)"; OkvedVidDeyatelnosti = "Торговля"; break;
                case "707174": OkvedNomer = "Операции с недвижимым имуществом; Аренда, бытовой прокат; Прочие услуги (70,71,74)"; OkvedVidDeyatelnosti = "Торговля"; break;
                case "72": OkvedNomer = "Деятельность в области информационных технологий (72)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "73": OkvedNomer = "Научные исследования и разработки (73)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "75": OkvedNomer = "Гос. управление и военная безопасность; соц. обеспечение (75)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "80": OkvedNomer = "Образование (80)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "85": OkvedNomer = "Здравоохранение и предоставление социальных услуг (85)"; OkvedVidDeyatelnosti = "Производство"; break;
                case "909395": OkvedNomer = "Социальные, коммунальные и персональные услуги (90-93,95)"; OkvedVidDeyatelnosti = "Производство"; break;
            }
        }
    }
}