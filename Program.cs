using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Globalization;

int maxStudents = 10000;
int numStudents = 0;

Dictionary<string, int> months = new Dictionary<string, int>();
months["Januari"] = 1;
months["Februari"] = 2;
months["Mars"] = 3;
months["April"] = 4;
months["Maj"] = 5;
months["Juni"] = 6;
months["Juli"] = 7;
months["Augusti"] = 8;
months["September"] = 9;
months["Oktober"] = 10;
months["November"] = 11;
months["December"] = 12;


// Initialize and open the ChromeDriver
ChromeOptions options = new ChromeOptions();
options.AddArguments("start-maximized");
options.AddArguments("--log-level=3");
ChromeDriverService service = ChromeDriverService.CreateDefaultService();
service.SuppressInitialDiagnosticInformation = true;
service.HideCommandPromptWindow = true;
IWebDriver driver = new ChromeDriver(service, options);
Actions actions = new Actions(driver);

// Go to login page
driver.Navigate().GoToUrl("https://auth.vklass.se/organisation/189");

Console.WriteLine("Tryck på tangenten \"Enter\" när du öppnat sidan \"Min klass\"");
Console.ReadLine();
Console.WriteLine("Börjar scanna efter elever...");

WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

var resultatsidor = wait.Until(e => e.FindElements(By.LinkText("Info & resultat")));
var baseURL = driver.Url;

string saveFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\VKlass-frånvaro\\";

if (!Directory.Exists(saveFolder))
{
    Directory.CreateDirectory(saveFolder);
}

var resultLinkList = new List<string>();
var overviewList = new List<AttendanceData>();

foreach (var item in resultatsidor)
{
    string resultURL = item.GetAttribute("href");
    resultLinkList.Add(resultURL);
}

Console.WriteLine($"Hittade {resultLinkList.Count} st elever");
DateTime startTime = DateTime.Now;
Console.WriteLine("Start time: " + startTime.ToString("HH:mm:ss"));

string studentOverview = File.ReadAllText("student-overview.html");

var studentOverviewHTMLList = new List<string>();

foreach (var item in resultLinkList)
{
    // Go to the students "Info & närvaro" page
    driver.Navigate().GoToUrl(item);

    // Go to "Närvaro"
    IWebElement närvaroLink = wait.Until(e => e.FindElement(By.LinkText("Närvaro")));
    driver.Navigate().GoToUrl(närvaroLink.GetAttribute("href"));

    // Get attendance data for last 30 days
    var studentAttendance = new AttendanceData();
    var overviewElement = wait.Until(e => e.FindElement(By.Id("ctl00_ContentPlaceHolder2_attendanceMinutesLabel")));
    var overviewString = overviewElement.Text;
    var overviewRows = overviewString.Split("\r\n");
    studentAttendance.Attendance = double.Parse(overviewRows[0].Split()[0]);
    var secondRow = overviewRows[1].Split();
    studentAttendance.ValidAbsence = double.Parse(secondRow[3]);
    studentAttendance.InvalidAbsence = double.Parse(secondRow[7]);

    // Find "Månadsvy"-link and click on it
    var månadsvyLink = wait.Until(e => e.FindElement(By.XPath("//span[text()='Månadsvy']")));
    månadsvyLink.Click();

    // Find student name
    var nameLink = wait.Until(e => e.FindElement(By.Id("ctl00_ContentPlaceHolder2_studentLink")));
    var name = nameLink.Text;
    studentAttendance.Name = name;

    // Scroll down a little bit for a better screenshot

    actions.ScrollByAmount(0, 200).Perform();

    // Get name of previous month
    var previousLink = wait.Until(e => e.FindElement(By.Id("ctl00_ContentPlaceHolder2_AttandanceOverviewControl_PreviousMonth")));
    var previousMonthName = previousLink.Text.Split()[1];

    // Get name of current month
    var currentMonthSpan = wait.Until(e => e.FindElement(By.XPath("//a/following-sibling::span")));
    var currentMonthName = currentMonthSpan.Text.Split()[0];
    var currentYear = currentMonthSpan.Text.Split()[1];

    // Get all lessons in current month
    var lessonList = GetLessonsOnCurrentPage();

    // Get lessons from previous month
    previousLink.Click();
    Thread.Sleep(1000);
    var previousMonthsLessons = GetLessonsOnCurrentPage();
    var allLessons = lessonList.Concat(previousMonthsLessons).DistinctBy(x => x.StartTime).ToList();


    /*
    // Take a screenshot and save it in a folder on the desktop
    Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
    ss.SaveAsFile($"{saveFolder + name + "-" + currentMonthName}.png", ScreenshotImageFormat.Png);

    // Take a screenshot of previous month
    previousLink.Click();
    Thread.Sleep(3000);
    Screenshot ssPrev = ((ITakesScreenshot)driver).GetScreenshot();
    ssPrev.SaveAsFile($"{saveFolder + name + "-" + previousMonthName}.png", ScreenshotImageFormat.Png);
    */

    // Save overview data
    studentAttendance.Lessons = allLessons;
    overviewList.Add(studentAttendance);
    //Console.WriteLine($"{studentAttendance.Name}\t{studentAttendance.Attendance}\t{studentAttendance.ValidAbsence}\t{studentAttendance.InvalidAbsence}");






    var studentFolder = saveFolder + studentAttendance.Name + "\\";
    if (!Directory.Exists(studentFolder))
    {
        Directory.CreateDirectory(studentFolder);
    }

    var studentOverviewHTML = GetHTMLAttendanceOverview(studentAttendance);
    File.WriteAllText(studentFolder + "overview.html", studentOverview.Replace("%%STUDENT%%", studentAttendance.Name).Replace("%%BODY%%", studentOverviewHTML));
    File.Copy("student-overview.css", studentFolder + "student-overview.css", true);
    studentOverviewHTMLList.Add(studentOverviewHTML);


    numStudents++;
    if (numStudents >= maxStudents)
    {
        break;
    }
    else if (numStudents == 1)
    {
        var timeForOneStudent = DateTime.Now;
        var estimatedTime = timeForOneStudent - startTime;
        var finishTime = startTime + estimatedTime * resultLinkList.Count;
        Console.WriteLine("Estimated total running time: " + (estimatedTime * resultLinkList.Count).ToString());
        Console.WriteLine("Estimated finish time: " + finishTime.ToString("HH:mm"));
        Console.WriteLine();
    }
}


string overviewTableRows = "";
foreach (var item in overviewList)
{
    string rowClass = "";
    if (item.Attendance <= 90)
    {
        rowClass = "medium-absence";
    }
    if (item.Attendance <= 80)
    {
        rowClass = "high-absence";
    }
    overviewTableRows += $"<tr class=\"{rowClass}\"><td><a href=\"{item.Name + "/overview.html"}\">{item.Name}</a></td><td>{item.Attendance.ToString(CultureInfo.CreateSpecificCulture("en-US"))}</td><td>{Math.Round(100 - item.Attendance, 2).ToString(CultureInfo.CreateSpecificCulture("en-US"))}</td><td>{item.ValidAbsence.ToString(CultureInfo.CreateSpecificCulture("en-US"))}</td><td>{item.InvalidAbsence.ToString(CultureInfo.CreateSpecificCulture("en-US"))}</td></tr>\n";
}


var overviewTemplate = File.ReadAllText("overview.html");
var newOverview = overviewTemplate.Replace("%%OVERVIEW-DATA%%", overviewTableRows);
newOverview = newOverview.Replace("%RUBRIK%", "Information hämtad " + DateTime.Now.ToString("yy-MM-dd HH:mm"));
File.WriteAllText(saveFolder + "overview.html", newOverview);
File.Copy("style.css", saveFolder + "style.css", true);
File.Copy("sort-table.min.js", saveFolder + "sort-table.min.js", true);

if (!Directory.Exists(saveFolder + "HelaKlassen\\"))
{
    Directory.CreateDirectory(saveFolder + "HelaKlassen\\");
}
File.WriteAllText(saveFolder + "HelaKlassen\\klassen.html", studentOverview.Replace("%%BODY%%", string.Join("",studentOverviewHTMLList)));
File.Copy("student-overview.css", saveFolder + "student-overview.css", true);

Console.WriteLine("Programmet avslutas...");
driver.Quit();

List<Lesson> GetLessonsOnCurrentPage()
{
    // Get name of previous month
    var previousLink = wait.Until(e => e.FindElement(By.Id("ctl00_ContentPlaceHolder2_AttandanceOverviewControl_PreviousMonth")));
    var previousMonthName = previousLink.Text.Split()[1];

    // Get name of current month
    var currentMonthSpan = wait.Until(e => e.FindElement(By.XPath("//a/following-sibling::span")));
    var currentMonthName = currentMonthSpan.Text.Split()[0];
    var currentYear = currentMonthSpan.Text.Split()[1];

    var lessons = wait.Until(e => e.FindElements(By.ClassName("lessonObject")));
    var dayList = new List<string>();

    foreach (var lesson in lessons)
    {
        actions.MoveToElement(lesson).Perform();
        Thread.Sleep(300);
        var par = lesson.FindElement(By.XPath("./.."));
        var pDay = par.FindElement(By.ClassName("DayDate"));
        dayList.Add(pDay.Text);
    }

    var tooltipsAfter = driver.FindElements(By.ClassName("rtWrapperContent"));
    var lessonList = new List<Lesson>();

    for (int i = 0; i < tooltipsAfter.Count; i++)
    {
        var tooltip = tooltipsAfter[i];
        var splitTooltip = tooltip.GetAttribute("innerHTML").Split(new string[] { "<div>", "<br>", "</div>" }, StringSplitOptions.TrimEntries);
        //Console.WriteLine($"{currentYear} {currentMonthName} {dayList[i]}");
        string clockInfo = "";
        string course = "";
        string status = "";
        int missingMinutes = 0;
        foreach (var line in splitTooltip)
        {
            //Console.WriteLine(line);
            if (line.Contains("kl: "))
            {
                clockInfo = line[4..];
            }
            else if (line.Contains("Kurs :"))
            {
                course = line[7..];
            }
            else if (line.Contains("Status"))
            {

                if (line.Length >= 8)
                {
                    status = line[8..];
                }
                else
                {
                    status = "EjRapporterat";
                }
            }
            else if (line.Contains("<span"))
            {
                int sInd = line.IndexOf(":: ") + 3;
                int eInd = line.IndexOf(" min");
                missingMinutes = int.Parse(line[sInd..eInd]);
            }
        }

        var newLesson = new Lesson();
        newLesson.StartTime = new DateTime(int.Parse(currentYear), months[currentMonthName], int.Parse(dayList[i]), int.Parse(clockInfo[..2]), int.Parse(clockInfo[3..5]), 0);
        newLesson.StopTime = new DateTime(int.Parse(currentYear), months[currentMonthName], int.Parse(dayList[i]), int.Parse(clockInfo[8..10]), int.Parse(clockInfo[11..13]), 0);
        newLesson.Course = course;
        newLesson.Status = status == "Närvarande" ? LessonStatus.Närvarande : status.Contains("Giltigt") ? LessonStatus.GiltigFrånvaro : status.Contains("Ej") ? LessonStatus.EjRapporterat : LessonStatus.OgiltigFrånvaro;
        newLesson.MissingMinutes = missingMinutes;
        lessonList.Add(newLesson);
    }

    return lessonList;
}

string GetHTMLAttendanceOverview(AttendanceData attendance)
{


    string template = @"
    <header>
        <h1>%%STUDENT%%</h1>
        <a href=""../overview.html"" class=""returnlink"">Tillbaka till klassöversikten</a>
        <div class=""topinfo"">
            <div class=""colorbox"">
                <div>Närvaro&nbsp;</div>
                <div class=""topbox Närvarande""></div>
            </div>
            <div class=""colorbox"">
                <div>Anmäld frånvaro&nbsp;</div>
                <div class=""topbox GiltigFrånvaro""></div>
            </div>
            <div class=""colorbox"">
                <div>Ogiltig frånvaro&nbsp;</div>
                <div class=""topbox OgiltigFrånvaro""></div>
            </div>
        </div>
    </header>

    %%CONTENT%%
";

    attendance.Lessons.Sort((x,y) => x.StartTime.CompareTo(y.StartTime));
    int startWeek = ISOWeek.GetWeekOfYear(attendance.Lessons.First().StartTime);
    int endWeek = ISOWeek.GetWeekOfYear(attendance.Lessons.Last().StartTime);
    int startYear = attendance.Lessons.First().StartTime.Year;
    int endYear = attendance.Lessons.Last().StartTime.Year;

    if (endWeek < startWeek)
    {
        endWeek += 52;
    }

    var time = attendance.Lessons.First().StartTime;
    var stoptime = attendance.Lessons.Last().StopTime;

    string res = "";
    res += "<table class=\"overviewtable\">\n" +
        "\t<tr><th width=\"50\">Vecka</th><th>Måndag</th><th>Tisdag</th><th>Onsdag</th><th>Torsdag</th><th>Fredag</th></tr>\n";
    // Loop through all weeks
    for (DateTime dt = time; dt <= stoptime; dt += TimeSpan.FromDays(7))
    {
        int showWeek = ISOWeek.GetWeekOfYear(dt);
        res += $"\t<tr><td class=\"week\">{showWeek}</td>";
        // Loop through Monday - Friday
        for (int i = 1; i <= 5; i++)
        {
            res += "<td>";
            var low = ISOWeek.ToDateTime(dt.Year, showWeek, (DayOfWeek)i);
            var high = ISOWeek.ToDateTime(dt.Year, showWeek, (DayOfWeek)(i + 1));

            var matchingLessons = attendance.Lessons.Where(x => x.StartTime > low && x.StartTime < high);



            foreach (var lesson in matchingLessons)
            {
                var lessonLength = (lesson.StopTime - lesson.StartTime).TotalMinutes;
                int fractionLate = (int)Math.Round(((lesson.MissingMinutes / lessonLength) * 100));
                res += $"<div class=\"{lesson.Status} lesson\"><div class=\"background-overlay{(fractionLate > 0 ? " show-late" : "")}\" style=\"--late: {fractionLate}%\"><div class=\"coursename\">{lesson.Course}</div>";
                //res += $"<div class=\"lessontime\">{lesson.StartTime.ToString("HH:mm")} - {lesson.StopTime.ToString("HH:mm")}</div>";
                res += $"</div></div>\n";
            }

            res += "</td>";
        }
        res += "</tr>\n";
    }
    res += "</table>\n";
    

    return template.Replace("%%STUDENT%%", attendance.Name).Replace("%%CONTENT%%", res);
}
class AttendanceData
{
    public string Name { get; set; }
    public double Attendance { get; set; }
    public double ValidAbsence { get; set; }
    public double InvalidAbsence { get; set; }
    public List<Lesson> Lessons { get; set; } = new List<Lesson>();
}

enum LessonStatus
{
    Närvarande,
    GiltigFrånvaro,
    OgiltigFrånvaro,
    EjRapporterat
}

class Lesson
{
    public DateTime StartTime { get; set; }
    public DateTime StopTime { get; set; }
    public string Course { get; set; }
    public LessonStatus Status { get; set; }
    public int MissingMinutes { get; set; }

}

