using System.Globalization;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;
using Humanizer;

class Program {
    const string salaryMainLineRegexp = @"основна винагорода, грн[.].*\n";
    const string salaryAdditionalLineRegexp = @"додаткова винагорода, грн[.].*\n";
    const string salaryVacationLineRegexp = @"оплата щорічної перерви, грн[.].*\n";
    const string salaryTotalLineRegexp = @"Загальна винагорода за .*\n";
    const string dateRegexp = @" \d\d[.]\d\d\d\d ";
    const string moneyRegexp = @" [\d ]+,[\d]+";

    static void Main() {
        string[] docxFiles = Directory.GetFiles(Environment.CurrentDirectory, "*.docx");
        string[] txtFiles = Directory.GetFiles(Environment.CurrentDirectory, "*.txt");
        if (!docxFiles.Any() || !txtFiles.Any()) return;
        ReplaceKeysWithValues(docxFiles[0], txtFiles[0]);
    }

    static void ReplaceKeysWithValues(string documentPath, string keyValueFilePath) {
        Application wordApp = new Application();
        Document doc = wordApp.Documents.Open(documentPath);

        Dictionary<string, string> keyValues = ReadKeyValues(keyValueFilePath);

        foreach (Range storyRange in doc.StoryRanges) {
            foreach (var keyValue in keyValues) {
                Find findObject = storyRange.Find;
                findObject.Text = keyValue.Key;
                findObject.Replacement.Text = keyValue.Value;
                object replaceAll = WdReplace.wdReplaceAll;
                findObject.Execute(Replace: replaceAll);
            }
        }

        doc.SaveAs2(documentPath);
        doc.Close();
        wordApp.Quit(); 
    }

    static Dictionary<string, string> ReadKeyValues(string filePath) {
        Dictionary<string, string> keyValues = new Dictionary<string, string>();

        string input = File.ReadAllText(filePath);

        string totalString = Regex.Match(input, salaryTotalLineRegexp, RegexOptions.IgnoreCase).Value;

        string dateString = Regex.Match(totalString, dateRegexp, RegexOptions.IgnoreCase).Value.Trim();
        DateTime.TryParseExact(dateString, "MM.yyyy", null, DateTimeStyles.None, out DateTime reportDate);

        decimal totalMoney = ParseMoney(totalString);

        string mainString = Regex.Match(input, salaryMainLineRegexp, RegexOptions.IgnoreCase).Value;
        decimal mainMoney = ParseMoney(mainString);

        string additionalString = Regex.Match(input, salaryAdditionalLineRegexp, RegexOptions.IgnoreCase).Value;
        decimal additionalMoney = ParseMoney(additionalString);

        string vacationString = Regex.Match(input, salaryVacationLineRegexp, RegexOptions.IgnoreCase).Value;
        decimal vacationMoney = ParseMoney(vacationString);

        keyValues["@$#%date_num@$#%"] = GetLastDayOfMonth(reportDate.Year, reportDate.Month);
        keyValues["@$#%month_ua@$#%"] = GetMonthNameUkrainian(reportDate.Month);
        keyValues["@$#%month_en@$#%"] = GetMonthNameEnglish(reportDate.Month);
        keyValues["@$#%year@$#%"] = reportDate.Year.ToString();

        MoneyToStr moneyToStr = new MoneyToStr("UAH", "UKR", "");

        keyValues["@$#%total@$#%"] = GetUahAmount(totalMoney);
        keyValues["@$#%total_text_en@$#%"] = Decimal.ToInt32(totalMoney).ToWords();
        keyValues["@$#%total_text_ua@$#%"] = moneyToStr.convertValue((double)totalMoney);
        keyValues["@$#%total_dec@$#%"] = Decimal.ToInt32(totalMoney % 1.0m * 100).ToString();

        keyValues["@$#%total_main@$#%"] = GetUahAmount(mainMoney);
        keyValues["@$#%total_main_text_en@$#%"] = Decimal.ToInt32(mainMoney).ToWords();
        keyValues["@$#%total_main_text_ua@$#%"] = moneyToStr.convertValue((double)mainMoney);
        keyValues["@$#%total_main_dec@$#%"] = Decimal.ToInt32(mainMoney % 1.0m * 100).ToString();

        keyValues["@$#%total_add@$#%"] = GetUahAmount(additionalMoney);
        keyValues["@$#%total_add_text_en@$#%"] = Decimal.ToInt32(additionalMoney).ToWords();
        keyValues["@$#%total_add_text_ua@$#%"] = moneyToStr.convertValue((double)additionalMoney);
        keyValues["@$#%total_add_dec@$#%"] = Decimal.ToInt32(additionalMoney % 1.0m * 100).ToString();

        keyValues["@$#%total_vac@$#%"] = GetUahAmount(vacationMoney);
        keyValues["@$#%total_vac_text_en@$#%"] = Decimal.ToInt32(vacationMoney).ToWords();
        keyValues["@$#%total_vac_text_ua@$#%"] = moneyToStr.convertValue((double)vacationMoney);
        keyValues["@$#%total_vac_dec@$#%"] = Decimal.ToInt32(vacationMoney % 1.0m * 100).ToString();

        return keyValues;
    }

    private static string GetUahAmount(decimal totalMoney) {
        return Decimal.ToInt32(totalMoney).ToString("N0", CultureInfo.CreateSpecificCulture("ru-RU"));
    }

    static decimal ParseMoney(string line) {
        string moneyString = Regex.Match(line, moneyRegexp, RegexOptions.IgnoreCase).Value
            .Trim()
            .Replace(" ","")
            .Replace(",",".");
        decimal parsedValue = decimal.Parse(moneyString, CultureInfo.InvariantCulture);
        return parsedValue;
    }

    static string GetMonthNameUkrainian(int monthNumber) =>
        monthNumber switch {
            1 => "січня",
            2 => "лютого",
            3 => "березня",
            4 => "квітня",
            5 => "травня",
            6 => "червня",
            7 => "липня",
            8 => "серпня",
            9 => "вересня",
            10 => "жовтня",
            11 => "листопада",
            12 => "грудня",
            _ => throw new Exception("Invalid month"),
        };

    static string GetMonthNameEnglish(int monthNumber) {
        DateTimeFormatInfo dateTimeFormat = new CultureInfo("en-US").DateTimeFormat;
        return dateTimeFormat.MonthNames[monthNumber - 1];
    }

    static string GetLastDayOfMonth(int year, int month) {
        int lastDay = DateTime.DaysInMonth(year, month);
        return lastDay.ToString();
    }
}