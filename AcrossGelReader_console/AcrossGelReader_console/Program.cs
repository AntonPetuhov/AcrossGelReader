using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Configuration;
using System.Data.SqlClient;

namespace AcrossGelReader_console
{
    class Program
    {
        #region settings

        public static string AnalyzerCode = "908";     // код из аналайзер конфигурейшн, который связывает прибор в PSMV2
        public static string AnalyzerConfigurationCode = "GELRDR"; // код прибора из аналайзер конфигурейшн

        public static string user = "PSMExchangeUser"; // логин для базы обмена файлами и для базы CGM Analytix
        public static string password = "PSM_123456";  // пароль для базы обмена файлами и для базы CGM Analytix 

        public static string COMPortName = "COM5"; // порт из nport administration
        public static bool ServiceIsActive;        // флаг для запуска и остановки потока
        static bool _continue;
        static SerialPort _serialPort;
        static int WaitTimeOut = 50;

        public static string AnalyzerResultPath = AppDomain.CurrentDomain.BaseDirectory + "\\AnalyzerResults"; // папка для файлов с результатами

        public static List<Thread> ListOfThreads = new List<Thread>(); // список работающих потоков 

        static object ExchangeLogLocker = new object();    // локер для логов обмена
        static object FileResultLogLocker = new object();  // локер для логов функции
        static object ServiceLogLocker = new object();     // локер для логов драйвера

        #region flags

        static byte[] ENQ = { 0x05 }; // запрос 
        static byte[] ACK = { 0x06 }; // подтверждение

        static byte[] STX = { 0x02 }; // начало текста
        static byte[] ETX = { 0x03 }; // конец текста

        #endregion

        #endregion

        #region logs

        // Лог драйвера
        static void ServiceLog(string Message)
        {
            lock (ServiceLogLocker)
            {
                try
                {
                    string path = AppDomain.CurrentDomain.BaseDirectory + "\\Log\\Service";
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                    string filename = path + "\\ServiceThread_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
                    if (!System.IO.File.Exists(filename))
                    {
                        using (StreamWriter sw = System.IO.File.CreateText(filename))
                        {
                            sw.WriteLine(DateTime.Now + ": " + Message);
                        }
                    }
                    else
                    {
                        using (StreamWriter sw = System.IO.File.AppendText(filename))
                        {
                            sw.WriteLine(DateTime.Now + ": " + Message);
                        }
                    }
                }
                catch
                {

                }
            }
        }

        // Лог обмена с прибором
        static void ExchangeLog(string Message)
        {
            lock (ExchangeLogLocker)
            {
                string path = AppDomain.CurrentDomain.BaseDirectory + "\\Log\\Exchange";
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                string filename = path + "\\ExchangeThread_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
                if (!System.IO.File.Exists(filename))
                {
                    using (StreamWriter sw = System.IO.File.CreateText(filename))
                    {
                        sw.WriteLine(DateTime.Now + ": " + Message);
                    }
                }
                else
                {
                    using (StreamWriter sw = System.IO.File.AppendText(filename))
                    {
                        sw.WriteLine(DateTime.Now + ": " + Message);
                    }
                }

            }
        }

        // Лог обработки файлов результатов для CGM
        static void FileResultLog(string Message)
        {
            try
            {
                lock (FileResultLogLocker)
                {
                    //string path = AppDomain.CurrentDomain.BaseDirectory + "\\Log\\FileResult" + "\\" + DateTime.Now.Year + "\\" + DateTime.Now.Month;
                    string path = AppDomain.CurrentDomain.BaseDirectory + "\\Log\\FileResult";
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                    //string filename = path + $"\\{FileName}" + ".txt";
                    string filename = path + $"\\ResultLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";

                    if (!System.IO.File.Exists(filename))
                    {
                        using (StreamWriter sw = System.IO.File.CreateText(filename))
                        {
                            sw.WriteLine(DateTime.Now + ": " + Message);
                        }
                    }
                    else
                    {
                        using (StreamWriter sw = System.IO.File.AppendText(filename))
                        {
                            sw.WriteLine(DateTime.Now + ": " + Message);
                        }
                    }
                }
            }
            catch
            {

            }
        }

        #endregion

        #region functions

        // Преобразование байтов в строку
        static string TranslateBytes(byte BytePar)
        {
            switch (BytePar)
            {
                case 0x02:
                    return "<STX>";
                case 0x03:
                    return "<ETX>";
                case 0x04:
                    return "<EOT>";
                case 0x05:
                    return "<ENQ>";
                case 0x06:
                    return "<ACK>";
                case 0x15:
                    return "<NAK>";
                case 0x16:
                    return "<SYN>";
                case 0x17:
                    return "<ETB>";
                case 0x0A:
                    return "<LF>";
                case 0x0D:
                    return "<CR>";
                default:
                    return "<HZ>";
            }
        }

        // Для удобства чтения логов, делаем из байт строку и заменяем в ней управляющие байты на символы UTF8. Иначе в строке будут нечитаемые символы.
        public static string GetStringFromBytes(byte[] ReceivedDataPar)
        {
            byte[] BytesForCHecking = { 0x02, 0x03, 0x04, 0x05, 0x06, 0x15, 0x16, 0x17, 0x0D, 0x0A };
            int StepCount = 0; // позиция обнаруженного байта
            bool IsManageByte = false;
            Encoding utf8 = Encoding.UTF8;
            // кодировка с кириллицей
            Encoding win1251 = Encoding.GetEncoding("windows-1251");

            // проверяем, является ли байт в массиве управляющим байтом
            foreach (byte rec_byte in ReceivedDataPar)
            {
                foreach (byte check_byte in BytesForCHecking)
                {
                    if (rec_byte == check_byte)
                    {
                        IsManageByte = true;
                        break;
                    }
                }
                if (IsManageByte)
                {
                    break;
                };
                StepCount++;
            }

            // Если обнаружен управляющий байт 
            if (IsManageByte)
            {
                // объявляем новый массив, в который будет записаны все оставшиеся байты, начиная со следующей позиции после обнаруженного  
                byte[] SliceByteArray = new byte[ReceivedDataPar.Length - (StepCount + 1)];

                //(из какого массива, с какого индекса, в какой массив, с какого индекса, кол-во элементов)
                Array.Copy(ReceivedDataPar, StepCount + 1, SliceByteArray, 0, ReceivedDataPar.Length - (StepCount + 1));

                /*
                // возвращаем преобразованную строку
                return utf8.GetString(ReceivedDataPar, 0, StepCount)
                    + TranslateBytes(ReceivedDataPar[StepCount])
                    + GetStringFromBytes(SliceByteArray);
                */
                // возвращаем преобразованную строку
                return win1251.GetString(ReceivedDataPar, 0, StepCount)
                    + TranslateBytes(ReceivedDataPar[StepCount])
                    + GetStringFromBytes(SliceByteArray);


            }
            else
            {
                //return utf8.GetString(ReceivedDataPar, 0, ReceivedDataPar.Length);
                return win1251.GetString(ReceivedDataPar, 0, ReceivedDataPar.Length);
            }
        }

        //дописываем к номеру месяца ноль если нужно
        public static string CheckZero(int CheckPar)
        {
            string BackPar = "";
            if (CheckPar < 10)
            {
                BackPar = $"0{CheckPar}";
            }
            else
            {
                BackPar = $"{CheckPar}";
            }
            return BackPar;
        }

        // Создание файлов с результатами, которые будут разбираться
        static void MakeAnalyzerResultFile(string AllMessagePar)
        {
            if (!Directory.Exists(AnalyzerResultPath))
            {
                Directory.CreateDirectory(AnalyzerResultPath);
            }

            DateTime now = DateTime.Now;
            string filename = AnalyzerResultPath + "\\Results_" + now.Year + CheckZero(now.Month) + CheckZero(now.Day) + CheckZero(now.Hour) + CheckZero(now.Minute) + CheckZero(now.Second) + CheckZero(now.Millisecond) + ".res";

            using (FileStream fstream = new FileStream(filename, FileMode.OpenOrCreate))
            {
                foreach (string res in AllMessagePar.Split('\r'))
                {
                    Encoding utf8 = Encoding.UTF8;
                    byte[] ResByte = utf8.GetBytes(res + "\r\n");
                    fstream.Write(ResByte, 0, ResByte.Length);
                }
            }
        }

        // интерпретация результатов
        public static string ResultInterpretation(string AnalyzerResult)
        {
            switch (AnalyzerResult)
            {
                case "CCEE": return "C(+)c(-)E(+)e(-)";
                case "CCEe": return "C(+)c(-)E(+)e(+)";
                case "CCeE": return "C(+)c(-)E(+)e(+)";
                case "CCee": return "C(+)c(-)E(-)e(+)";

                case "CcEE": return "C(+)c(+)E(+)e(-)";
                case "CcEe": return "C(+)c(+)E(+)e(+)";
                case "CceE": return "C(+)c(+)E(+)e(+)";
                case "Ccee": return "C(+)c(+)E(-)e(+)";

                case "cCEE": return "C(+)c(+)E(+)e(-)";
                case "cCEe": return "C(+)c(+)E(+)e(+)";
                case "cCeE": return "C(+)c(+)E(+)e(+)";
                case "cCee": return "C(+)c(+)E(-)e(+)";

                case "ccEE": return "C(-)c(+)E(+)e(-)";
                case "ccEe": return "C(-)c(+)E(+)e(+)";
                case "cceE": return "C(-)c(+)E(+)e(+)";
                case "ccee": return "C(-)c(+)E(-)e(+)";
                case "A": return "Группа А (II)";
                case "AB": return "Группа АВ (IV)";
                case "B": return "Группа В (III)";
                case "0": return "Группа 0 (I)";
                case "A2": return "Группа А2(II)";
                case "A2B": return "Группа A2B(IV)";
                case "Поз.": return "положительный";
                case "Нег.": return "Отрицательный";
                case "Positive": return "положительный";
                case "Negative": return "Отрицательный";
                case "+": return "положительный";
                case "-": return "Отрицательный";
                default: return AnalyzerResult;
            }
        }

        // Функция преобразования кода теста прибора в код теста PSMV2 в CGM
        public static string TranslateToPSMCodes(string AnalyzerTestCodesPar)
        {
            string BackTestCode = "";
            try
            {
                string CGMConnectionString = ConfigurationManager.ConnectionStrings["CGMConnection"].ConnectionString;
                //string CGMConnectionString = @"Data Source=CGM-APP11\SQLCGMAPP11;Initial Catalog=KDLPROD; Integrated Security=True;";
                CGMConnectionString = String.Concat(CGMConnectionString, $"User Id = {user}; Password = {password}");
                using (SqlConnection Connection = new SqlConnection(CGMConnectionString))
                {
                    Connection.Open();
                    // Ищем только тесты, которые настроены для прибора и настроены для PSMV2
                    SqlCommand TestCodeCommand = new SqlCommand(
                       "SELECT k1.amt_analyskod  FROM konvana k " +
                       "LEFT JOIN konvana k1 ON k1.met_kod = k.met_kod AND k1.ins_maskin = 'PSMV2' " +
                       $"WHERE k.ins_maskin = '{AnalyzerConfigurationCode}' AND k.amt_analyskod = '{AnalyzerTestCodesPar}' ", Connection);
                    SqlDataReader Reader = TestCodeCommand.ExecuteReader();

                    if (Reader.HasRows) // если есть данные
                    {
                        while (Reader.Read())
                        {
                            if (!Reader.IsDBNull(0)) { BackTestCode = Reader.GetString(0); };
                        }
                    }
                    Reader.Close();
                    Connection.Close();
                }
            }
            catch (Exception error)
            {
                FileResultLog($"Error: {error}");
            }

            return BackTestCode;
        }

        // поиск результатов и тестов, в зависимости от типа гелевой карты
        public static string FindResultsInString(string gel_card, string res)
        {
            switch (gel_card)
            {
                case "Across Gel? Rh Phenotyping with Kell (K)":

                    FileResultLog($"Тип карты: {gel_card}");
                    FileResultLog($"Строка для поиска результата: '{res}'");

                    #region Определеяем результаты для тестов текущей гелевой карты, формируем и возвращаем строку с результатами
                    // определим тесты с прибора
                    // можно было бы прописать сразу коды PSM, но пусть лучше будет единообразно
                    string TestPheno = "Pheno";
                    string TestKell = "Kell";
                    string TestRh = "Rh";

                    // переменные для результатов
                    string PhenoRes = "";
                    string KellRes = "";
                    string RhRes = "";

                    // строки с результатами для каждого теста
                    string strPheno = "";
                    string strKell = "";
                    string strRh = "";
                    // итоговая строка с результатами
                    string resultStr = "";

                    // регулярные выражения для поиска результатов
                    string PhenoPattern = @"(?<Pheno>\w+)\s/\s.+";
                    //string KellPattern = @"\w+\s/\sKell\s(?<kell>[+-])\s";
                    string KellPattern = @"\w+\s/\sKell\s(?<kell>\w+)\s";
                    //string RhPattern = @"\w+\s/\sKell\s[+-]\sRH\s(?<rh>[+-])";
                    //string RhPattern = @"RH\s(?<rh>[+-])";
                    string RhPattern = @"Rh\s(?<rh>\w+)";

                    Regex PhenoRegex = new Regex(PhenoPattern, RegexOptions.None, TimeSpan.FromMilliseconds(150));
                    Regex KellRegex = new Regex(KellPattern, RegexOptions.None, TimeSpan.FromMilliseconds(150));
                    Regex RhRegex = new Regex(RhPattern, RegexOptions.None, TimeSpan.FromMilliseconds(150));

                    Match PhenoMatch = PhenoRegex.Match(res);
                    Match KellMatch = KellRegex.Match(res);
                    Match RhMatch = RhRegex.Match(res);

                    if (PhenoMatch.Success) 
                    { 
                        PhenoRes = PhenoMatch.Result("${Pheno}");
                        FileResultLog($"Тест {TestPheno}, результат: {PhenoRes}");
                        string PhenoPSMTestCode = TranslateToPSMCodes(TestPheno);
                        FileResultLog($"Тест {TestPheno} преобразован в код CGM (PSMV2): {PhenoPSMTestCode}");
                        FileResultLog($"Результат '{PhenoRes}' интерпретирован: {ResultInterpretation(PhenoRes)}");

                        // формируем строку с результатом
                        //string strPheno = $"R|1|^^^{PhenoPSMTestCode}^^^^{AnalyzerCode}|{PhenoRes}|||N||F||Chemwell^||20240101000001|{AnalyzerCode}" + "\r";
                        strPheno = $"R|1|^^^{PhenoPSMTestCode}^^^^{AnalyzerCode}|{ResultInterpretation(PhenoRes)}|||N||F||AcrossGelReader^||20240101000001|{AnalyzerCode}";

                        resultStr = resultStr + strPheno + "\r";
                    }

                    if (KellMatch.Success) 
                    { 
                        KellRes = KellMatch.Result("${kell}");
                        FileResultLog($"Тест {TestKell}, результат: {KellRes}");
                        string KellPSMTestCode = TranslateToPSMCodes(TestKell);
                        FileResultLog($"Тест {TestKell} преобразован в код CGM (PSMV2): {KellPSMTestCode}");
                        FileResultLog($"Результат '{KellRes}' интерпретирован: {ResultInterpretation(KellRes)}");

                        // формируем строку с результатом
                        strKell = $"R|1|^^^{KellPSMTestCode}^^^^{AnalyzerCode}|{ResultInterpretation(KellRes)}|||N||F||AcrossGelReader^||20240101000001|{AnalyzerCode}";

                        resultStr = resultStr + strKell + "\r";
                    }

                    if (RhMatch.Success) 
                    { 
                        RhRes = RhMatch.Result("${rh}");
                        FileResultLog($"Тест {TestRh}, результат: {RhRes}");
                        string RhPSMTestCode = TranslateToPSMCodes(TestRh);
                        FileResultLog($"Тест {TestRh} преобразован в код CGM (PSMV2): {RhPSMTestCode}");
                        FileResultLog($"Результат '{RhRes}' интерпретирован: {ResultInterpretation(RhRes)}");

                        // формируем строку с результатом
                        strRh = $"R|1|^^^{RhPSMTestCode}^^^^{AnalyzerCode}|{ResultInterpretation(RhRes)}|||N||F||AcrossGelReader^||20240101000001|{AnalyzerCode}";

                        resultStr = resultStr + strRh;
                    }

                   // resultStr = strPheno + "\r" + strKell + "\r" + strRh;
                   // FileResultLog("Строка с заданием: " + "\r" + resultStr);
                    if (resultStr == "")
                    {
                        FileResultLog($"Результаты не найдены.");
                    }

                    #endregion

                    return resultStr;


                case "Across Gel? Forward & Reverse ABO with Dv?-/Kell":

                    FileResultLog($"Тип карты: {gel_card}");
                    FileResultLog($"Строка для поиска результата: '{res}'");

                    #region Определеяем результаты для тестов текущей гелевой карты, формируем и возвращаем строку с результатами
                    // определим тесты с прибора
                    // можно было бы прописать сразу коды PSM, но пусть лучше будет единообразно
                    string TestABO = "ABO";
                    string TestKell_ = "Kell";
                    string TestRh_ = "Rh";

                    // переменные для результатов
                    string ABORes = "";
                    string Kell_Res = "";
                    string Rh_Res = "";

                    // строки с результатами для каждого теста
                    string strABO = "";
                    string strKell_ = "";
                    string strRh_ = "";
                    // итоговая строка с результатами
                    string resultStr_ = "";

                    // регулярные выражения для поиска результатов
                    string ABOPattern = @"(?<AB0>\S+)\sRH\s[+-]\sKell\s";
                    string Kell_Pattern = @"RH\s[+-]\sKell\s(?<kell>[+-])";
                    string Rh_Pattern = @"RH\s(?<rh>[+-])\sKell\s[+-]";

                    Regex ABORegex = new Regex(ABOPattern, RegexOptions.None, TimeSpan.FromMilliseconds(150));
                    Regex Kell_Regex = new Regex(Kell_Pattern, RegexOptions.None, TimeSpan.FromMilliseconds(150));
                    Regex Rh_Regex = new Regex(Rh_Pattern, RegexOptions.None, TimeSpan.FromMilliseconds(150));

                    Match ABOMatch = ABORegex.Match(res);
                    Match Kell_Match = Kell_Regex.Match(res);
                    Match Rh_Match = Rh_Regex.Match(res);

                    if (ABOMatch.Success)
                    {
                        ABORes = ABOMatch.Result("${AB0}");
                        FileResultLog($"Тест {TestABO}, результат: {ABORes}");
                        string ABOPSMTestCode = TranslateToPSMCodes(TestABO);
                        FileResultLog($"Тест {TestABO} преобразован в код CGM (PSMV2): {ABOPSMTestCode}");
                        FileResultLog($"Результат '{ABORes}' интерпретирован: {ResultInterpretation(ABORes)}");

                        // формируем строку с результатом
                        //string strPheno = $"R|1|^^^{PhenoPSMTestCode}^^^^{AnalyzerCode}|{PhenoRes}|||N||F||Chemwell^||20240101000001|{AnalyzerCode}" + "\r";
                        strABO = $"R|1|^^^{ABOPSMTestCode}^^^^{AnalyzerCode}|{ResultInterpretation(ABORes)}|||N||F||AcrossGelReader^||20240101000001|{AnalyzerCode}";

                        resultStr_ = resultStr_ + strABO + "\r";
                    }

                    if (Kell_Match.Success)
                    {
                        Kell_Res = Kell_Match.Result("${kell}");
                        FileResultLog($"Тест {TestKell_}, результат: {Kell_Res}");
                        string Kell_PSMTestCode = TranslateToPSMCodes(TestKell_);
                        FileResultLog($"Тест {TestKell_} преобразован в код CGM (PSMV2): {Kell_PSMTestCode}");
                        FileResultLog($"Результат '{Kell_Res}' интерпретирован: {ResultInterpretation(Kell_Res)}");

                        // формируем строку с результатом
                        strKell_ = $"R|1|^^^{Kell_PSMTestCode}^^^^{AnalyzerCode}|{ResultInterpretation(Kell_Res)}|||N||F||AcrossGelReader^||20240101000001|{AnalyzerCode}";

                        resultStr_ = resultStr_ + strKell_ + "\r";
                    }

                    if (Rh_Match.Success)
                    {
                        Rh_Res = Rh_Match.Result("${rh}");
                        FileResultLog($"Тест {TestRh_}, результат: {Rh_Res}");
                        string Rh_PSMTestCode = TranslateToPSMCodes(TestRh_);
                        FileResultLog($"Тест {TestRh_} преобразован в код CGM (PSMV2): {Rh_PSMTestCode}");
                        FileResultLog($"Результат '{Rh_Res}' интерпретирован: {ResultInterpretation(Rh_Res)}");

                        // формируем строку с результатом
                        strRh_ = $"R|1|^^^{Rh_PSMTestCode}^^^^{AnalyzerCode}|{ResultInterpretation(Rh_Res)}|||N||F||AcrossGelReader^||20240101000001|{AnalyzerCode}";

                        resultStr_ = resultStr_ + strRh_;
                    }

                    if (resultStr_ == "")
                    {
                        FileResultLog($"Результаты не найдены.");
                    }

                    #endregion

                    return resultStr_;

                case "Across Gel? Forward & Reverse ABO with Dv?-/Dv?+":

                    FileResultLog($"Тип карты: {gel_card}");
                    FileResultLog($"Строка для поиска результата: '{res}'");

                    #region Определеяем результаты для тестов текущей гелевой карты, формируем и возвращаем строку с результатами
                    // определим тесты с прибора
                    // можно было бы прописать сразу коды PSM, но пусть лучше будет единообразно
                    string TestAB0_AB0Dv = "ABO";
                    string TestRh_AB0Dv = "Rh";

                    // переменные для результатов
                    string AB0Res_AB0Dv = "";
                    string Rh_AB0Dv_Res = "";

                    // строки с результатами для каждого теста
                    string strABO_AB0Dv = "";
                    string strRh_AB0Dv = "";
                    // итоговая строка с результатами
                    string resultStr_AB0Dv = "";

                    // регулярные выражения для поиска результатов
                    //string ABO_AB0Dv_Pattern = @".+;(?<AB0>\S+)\sRH\s[+-]\s";
                    string ABO_AB0Dv_Pattern = @"(?<AB0>\S+)\sRH\s\w+\s";
                    //string Rh_AB0Dv_Pattern = @"RH\s(?<rh>[+-])\s";
                    string Rh_AB0Dv_Pattern = @"RH\s(?<rh>\w+)\s";

                    Regex ABO_AB0Dv_Regex = new Regex(ABO_AB0Dv_Pattern, RegexOptions.None, TimeSpan.FromMilliseconds(150));
                    Regex Rh_AB0Dv_Regex = new Regex(Rh_AB0Dv_Pattern, RegexOptions.None, TimeSpan.FromMilliseconds(150));

                    Match ABO_AB0Dv_Match = ABO_AB0Dv_Regex.Match(res);
                    Match Rh_AB0Dv_Match = Rh_AB0Dv_Regex.Match(res);

                    if (ABO_AB0Dv_Match.Success)
                    {
                        AB0Res_AB0Dv = ABO_AB0Dv_Match.Result("${AB0}");
                        FileResultLog($"Тест {TestAB0_AB0Dv}, результат: {AB0Res_AB0Dv}");
                        string ABO_AB0Dv_PSMTestCode = TranslateToPSMCodes(TestAB0_AB0Dv);
                        FileResultLog($"Тест {TestAB0_AB0Dv} преобразован в код CGM (PSMV2): {ABO_AB0Dv_PSMTestCode}");
                        FileResultLog($"Результат '{AB0Res_AB0Dv}' интерпретирован: {ResultInterpretation(AB0Res_AB0Dv)}");

                        // формируем строку с результатом
                        //string strPheno = $"R|1|^^^{PhenoPSMTestCode}^^^^{AnalyzerCode}|{PhenoRes}|||N||F||Chemwell^||20240101000001|{AnalyzerCode}" + "\r";
                        strABO_AB0Dv = $"R|1|^^^{ABO_AB0Dv_PSMTestCode}^^^^{AnalyzerCode}|{ResultInterpretation(AB0Res_AB0Dv)}|||N||F||AcrossGelReader^||20240101000001|{AnalyzerCode}";

                        resultStr_AB0Dv = resultStr_AB0Dv + strABO_AB0Dv + "\r";
                    }

                    if (Rh_AB0Dv_Match.Success)
                    {
                        Rh_AB0Dv_Res = Rh_AB0Dv_Match.Result("${rh}");
                        FileResultLog($"Тест {TestRh_AB0Dv}, результат: {Rh_AB0Dv_Res}");
                        string Rh_AB0Dv_PSMTestCode = TranslateToPSMCodes(TestRh_AB0Dv);
                        FileResultLog($"Тест {TestRh_AB0Dv} преобразован в код CGM (PSMV2): {Rh_AB0Dv_PSMTestCode}");
                        FileResultLog($"Результат '{Rh_AB0Dv_Res}' интерпретирован: {ResultInterpretation(Rh_AB0Dv_Res)}");

                        // формируем строку с результатом
                        strRh_AB0Dv = $"R|1|^^^{Rh_AB0Dv_PSMTestCode}^^^^{AnalyzerCode}|{ResultInterpretation(Rh_AB0Dv_Res)}|||N||F||AcrossGelReader^||20240101000001|{AnalyzerCode}";

                        resultStr_AB0Dv = resultStr_AB0Dv + strRh_AB0Dv;
                    }

                    if (resultStr_AB0Dv == "")
                    {
                        FileResultLog($"Результаты не найдены.");
                    }

                    #endregion

                    return resultStr_AB0Dv;

                case "Across Gel? AHG IgG-C3d (Indirect Coombs 4)":

                    FileResultLog($"Тип карты: {gel_card}");
                    FileResultLog($"Строка для поиска результата: '{res}'");

                    #region Определеяем результаты для тестов текущей гелевой карты, формируем и возвращаем строку с результатами
                    // определим тесты с прибора
                    // можно было бы прописать сразу коды PSM, но пусть лучше будет единообразно
                    string TestAHG = "AHG";
                    // переменные для результатов
                    string AHGRes = "";
                    // строки с результатами для каждого теста
                    string strAHG = "";

                    // регулярные выражения для поиска результатов
                    string AHGPattern = @"Indirect\sCoombs\s(?<AHG>\w+)";
                    Regex AHGRegex = new Regex(AHGPattern, RegexOptions.None, TimeSpan.FromMilliseconds(150));
                    Match AHGMatch = AHGRegex.Match(res);

                    if (AHGMatch.Success)
                    {
                        AHGRes = AHGMatch.Result("${AHG}");
                        FileResultLog($"Тест {TestAHG}, результат: {AHGRes}");
                        string AHGPSMTestCode = TranslateToPSMCodes(TestAHG);
                        FileResultLog($"Тест {TestAHG} преобразован в код CGM (PSMV2): {AHGPSMTestCode}");
                        FileResultLog($"Результат '{AHGRes}' интерпретирован: {ResultInterpretation(AHGRes)}");

                        // формируем строку с результатом
                        //string strPheno = $"R|1|^^^{PhenoPSMTestCode}^^^^{AnalyzerCode}|{PhenoRes}|||N||F||Chemwell^||20240101000001|{AnalyzerCode}" + "\r";
                        strAHG = $"R|1|^^^{AHGPSMTestCode}^^^^{AnalyzerCode}|{ResultInterpretation(AHGRes)}|||N||F||AcrossGelReader^||20240101000001|{AnalyzerCode}";
                    }

                    if (strAHG == "")
                    {
                        FileResultLog($"Результаты не найдены.");
                    }
                    #endregion

                    return strAHG;

                default:

                    FileResultLog($"Тип карты: '{gel_card}'. Поиск результата для карты не настроен. Не получилось найти результат в строке '{res}'");
                    return "";
            }
        }

            #endregion

        static void ResultsProcessing()
        {
            while (ServiceIsActive)
            {
                #region папки архива, результатов и ошибок

                string OutFolder = ConfigurationManager.AppSettings["FolderOut"];
                
                // для тестирования
                //string OutFolder = AnalyzerResultPath + @"\CGM";

                // архивная папка
                string ArchivePath = AnalyzerResultPath + @"\Archive";
                // папка для ошибок
                string ErrorPath = AnalyzerResultPath + @"\Error";
                // папка для файлов с результатами для CGM
                string CGMPath = AnalyzerResultPath + @"\CGM";

                if (!Directory.Exists(ArchivePath))
                {
                    Directory.CreateDirectory(ArchivePath);
                }

                /*
                if (!Directory.Exists(ErrorPath))
                {
                    Directory.CreateDirectory(ErrorPath);
                }

                if (!Directory.Exists(CGMPath))
                {
                    Directory.CreateDirectory(CGMPath);
                }
                */
                #endregion

                // получаем список всех файлов в текущей папке
                string[] Files = Directory.GetFiles(AnalyzerResultPath, "*.res");

                // шаблоны регулярных выражений для поиска данных
                string RIDPattern = @"[[]0[]];(?<RID>\d+);";
                string GelCardPattern = @"[[]1[]];\d+;;(?<GelCard>.+);";
                //string GelCardPattern = @"[[]1[]];\d+;;Across Gel[?] (?<GelCard>.+);";
                //string ResultPattern = @"[[]0[]];\d+;\S+;(?<Result>.+)";
                string ResultPattern = @"[[]0[]];\d+;.+;(?<Result>.+)";


                Regex RIDRegex = new Regex(RIDPattern, RegexOptions.None, TimeSpan.FromMilliseconds(150));
                Regex ResultRegex = new Regex(ResultPattern, RegexOptions.None, TimeSpan.FromMilliseconds(150));
                Regex GelCardRegex = new Regex(GelCardPattern, RegexOptions.None, TimeSpan.FromMilliseconds(150));

                // пробегаем по файлам с результатами 
                foreach (string file in Files)
                {
                    // строки для формирования файла (psm файла) с результатами для службы,
                    // которая разбирает файлы и записывает результаты в CGM
                    string MessageHead = "";
                    string MessageTest = "";
                    string AllMessage = "";

                    FileResultLog("Обработка файлов с результатами и формирование файлов для CGM");
                    FileResultLog(file);

                    string[] lines = System.IO.File.ReadAllLines(file);

                    string RID = "";
                    string Result = "";
                    string GelCard = "";

                    // обрезаем только имя текущего файла
                    string FileName = file.Substring(AnalyzerResultPath.Length + 1);
                    // название файла .ок, который должен создаваться вместе с результирующим для обработки службой FileGetterService
                    string OkFileName = "";

                    // проходим по строкам в файле
                    foreach (string line in lines)
                    {
                        Match RIDMatch = RIDRegex.Match(line);
                        Match ResultMatch = ResultRegex.Match(line);
                        Match GelCardMatch = GelCardRegex.Match(line);

                        // переменная для временного хранения строки с результатом
                        string line_ = "";

                        // если нашли ШК в строке
                        if (RIDMatch.Success)
                        {
                            RID = RIDMatch.Result("${RID}");
                            FileResultLog($"Заявка № {RID}");
                            MessageHead = $"O|1|{RID}||ALL|R|20240101000100|||||X||||ALL||||||||||F";

                            // запомним эту строчку для передачи в функцию поиска результата
                            line_ = line;
                        }

                        // найдем результат в строке
                        if (ResultMatch.Success)
                        {
                            Result = ResultMatch.Result("${Result}");
                        }

                        if (GelCardMatch.Success)
                        {
                            GelCard = GelCardMatch.Result("${GelCard}");

                            // ищем результат и формируем строку
                            //FindResultsInString(GelCard, Result);
                            MessageTest = FindResultsInString(GelCard, Result);

                            // обнулим переменные
                            RID = "";
                            Result = "";
                        }

                        // если сформированы строки с заявкой и заданием
                        if (MessageHead != "" && MessageTest != "")
                        {
                            // собираем полное сообщение с результатом
                            AllMessage = MessageHead + "\r" + MessageTest;
                            FileResultLog("Сообщение с заданием:" + "\r" + AllMessage);
                            FileResultLog("");
                            //FileResultLog(AllMessage);

                            #region создаем файл с результатом и помещаем его в папку для обработки другой службой
                            // создаем файл с результатом и записываем его в папку

                            DateTime now = DateTime.Now;
                            string resultFileName = "Results_" + now.Year + CheckZero(now.Month) + CheckZero(now.Day) + CheckZero(now.Hour) + CheckZero(now.Minute) + CheckZero(now.Second) + CheckZero(now.Millisecond) + ".res"; ;
                            // получаем название файла .ок на основании файла с результатом
                            if (FileName.IndexOf(".") != -1)
                            {
                                OkFileName = resultFileName.Split('.')[0] + ".ok";
                            }

                            FileStream fs = null;

                            try
                            {
                                // создаем файл для записи результата в папке для рез-тов
                                if (!File.Exists(OutFolder + @"\" + resultFileName))
                                {
                                    /*
                                    using (StreamWriter sw = File.CreateText(OutFolder + @"\" + resultFileName))
                                    {
                                        foreach (string msg in AllMessage.Split('\r'))
                                        {
                                            sw.WriteLine(msg);
                                        }
                                    }
                                    */

                                    fs = new FileStream(OutFolder + @"\" + resultFileName, FileMode.CreateNew);

                                    using (StreamWriter writer = new StreamWriter(fs, Encoding.GetEncoding("windows-1251")))
                                    {
                                        /*
                                        foreach (string msg in AllMessage.Split('\r'))
                                        {
                                            sw.WriteLine(msg);
                                        }
                                        */
                                        writer.Write(AllMessage);
                                    }

                                    //File.WriteAllText(OutFolder + resultFileName, AllMessage, Encoding.GetEncoding("windows-1251"));
                                }
                                else
                                {
                                    File.Delete(OutFolder + @"\" + resultFileName);
                                    /*
                                    using (StreamWriter sw = File.CreateText(OutFolder + @"\" + resultFileName))
                                    {
                                        foreach (string msg in AllMessage.Split('\r'))
                                        {
                                            sw.WriteLine(msg);
                                        }
                                    }
                                    */
                                    fs = new FileStream(OutFolder + @"\" + resultFileName, FileMode.CreateNew);
                                    using (StreamWriter writer = new StreamWriter(fs, Encoding.GetEncoding("windows-1251")))
                                    {
                                        /*
                                        foreach (string msg in AllMessage.Split('\r'))
                                        {
                                            sw.WriteLine(msg);
                                        }
                                        */
                                        writer.Write(AllMessage);
                                    }

                                    //File.WriteAllText(OutFolder + resultFileName, AllMessage, Encoding.GetEncoding("windows-1251"));
                                }

                                // создаем .ok файл в папке для рез-тов
                                if (OkFileName != "")
                                {
                                    if (!File.Exists(OutFolder + @"\" + OkFileName))
                                    {
                                        //using (StreamWriter sw = File.CreateText(CGMPath + @"\" + OkFileName))
                                        using (StreamWriter sw = File.CreateText(OutFolder + @"\" + OkFileName))
                                        {
                                            sw.WriteLine("ok");
                                        }
                                    }
                                    else
                                    {
                                        File.Delete(OutFolder + OkFileName);
                                        using (StreamWriter sw = File.CreateText(OutFolder + @"\" + OkFileName))
                                        {
                                            sw.WriteLine("ok");
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                FileResultLog($"{ex}");
                            }

                            #endregion

                            // после формирования файла обнулим переменные
                            MessageHead = "";
                            MessageTest = "";
                        }
                    }

                    // помещение файла в архивную папку
                    if (File.Exists(ArchivePath + @"\" + FileName))
                    {
                        File.Delete(ArchivePath + @"\" + FileName);
                    }
                    File.Move(file, ArchivePath + @"\" + FileName);

                    FileResultLog($"Файл {FileName} обработан и перемещен в папку Archive");
                    FileResultLog("");

                }

                Thread.Sleep(2000);
            }
        }

        static void ReadFromCOM()
        {
            int ServerCount = 0; // счетчик

            while (_continue)
            {
                //int ServerCount = 0; // счетчик

                // проверяем количество байтов, чтобы понимать, было ли отправлено что-то
                int CheckBytesCount = _serialPort.BytesToRead;

                // если есть байты для считывания
                if (CheckBytesCount > 0)
                {
                    Thread.Sleep(2000);

                    // определяем кол-во байт (еще раз, т.к. прибор довольно медленно передает, поэтому и таймаут)
                    int bytesCount = _serialPort.BytesToRead;
                    ServiceLog($"Bytes to read: {bytesCount}. Reading from serial port.");
                    // массив байтов
                    byte[] ByteArray = new byte[bytesCount];
                    _serialPort.Read(ByteArray, 0, bytesCount);
                    

                    // Для удобства дальнейшего чтения логов, формируем строку из считанного массива байт, заменяя управляющие байты, на символы UTF8
                    // иначе будут нечитаемые символы
                    string TMPString = "";
                    TMPString = GetStringFromBytes(ByteArray);

                    // пишем сообщение от прибора в лог обмена
                    ExchangeLog($"Analyzer: {TMPString}");
                    ExchangeLog($"");

                    // добавляем символ разделитель
                    TMPString = TMPString.Replace("<ETX>", "@");

                    // обработанное сообщение с результатом от прибора
                    string ResultMessage = "";

                    // разбиваем на массив строк
                    string[] mass = TMPString.Split('@');
                    foreach (string str_ in mass)
                    {
                        if (str_.IndexOf("<STX>")!= -1)
                        {
                            string res_str;
                            // отрезаем <STX> и CSUM перед ним
                            res_str = str_.Substring(str_.IndexOf("<STX>")+5);
                            // формируем строку без лишнего
                            ResultMessage = ResultMessage + res_str + '\r';
                        }
                    }

                    // записываем файл с результатом
                    MakeAnalyzerResultFile(ResultMessage);
                }
                // если прибор ничего не посылает
                else
                {              
                    ServerCount++;
                    
                    if (ServerCount == 30)
                    {
                        ServerCount = 0;
                        ServiceLog($"Bytes to read: {CheckBytesCount}. Listening serial port.");
                    }
                }

                Thread.Sleep(1000);
            }
        }
        static void COMPortSettings()
        {
            Thread readThread = new Thread(ReadFromCOM);
            readThread.Name = "ReadCOM";
            //ListOfThreads.Add(readThread);

            try
            {
                // Create a new SerialPort object
                _serialPort = new SerialPort();

                // настройки СОМ порта
                _serialPort.PortName = COMPortName;
                _serialPort.BaudRate = 9600;
                _serialPort.Parity = (Parity)Enum.Parse(typeof(Parity), "None", true);
                _serialPort.DataBits = 8;
                _serialPort.StopBits = (StopBits)Enum.Parse(typeof(StopBits), "One", true);
                _serialPort.Handshake = (Handshake)Enum.Parse(typeof(Handshake), "None", true);

                // Set the read/write timeouts
                _serialPort.ReadTimeout = 500;
                _serialPort.WriteTimeout = 500;

                _serialPort.Open();
                _continue = true;

                Console.WriteLine();

                // Запуск потока чтения порта
                readThread.Start();
                Console.WriteLine("Reading thread is started");
                ServiceLog("Reading thread is started");

            }
            catch(Exception ex)
            {
                Console.WriteLine("ERROR. Port cannot be opened:" + ex.ToString());
                ServiceLog("ERROR. Port cannot be opened:" + ex.ToString());
                return;
            }

        }

        static void Main(string[] args)
        {
            ServiceIsActive = true;
            Console.WriteLine("Service starts working");
            ServiceLog("Service starts working");

            // Настраиваем и запускаем поток чтения из ком порта
            COMPortSettings();

            // Поток обработки файлов с результатами
            Thread ResultProcessingThread = new Thread(ResultsProcessing);
            ResultProcessingThread.Name = "ResultsProcessing";
            ListOfThreads.Add(ResultProcessingThread);
            ResultProcessingThread.Start();

            Console.WriteLine("Service is working");
            ServiceLog("Service is working");


            Console.ReadLine();

        }
    }
}
