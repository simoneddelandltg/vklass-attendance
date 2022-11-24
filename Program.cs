using FrånvaroöversiktVKlass;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Globalization;


// Change maxStudents for testing purposes to get an overview after only a few students have been processed
int maxStudents = 10000;
int numStudents = 0;

// Dictionary needed to convert months in swedish to month number
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

// Update ChromeDriver if needed
ChromeDriverInstaller installer = new ChromeDriverInstaller();
var installProcess = installer.Install();
await installProcess;

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

// Print user information
Console.WriteLine("Detta program hjälper dig att få en grafisk frånvaroöversikt från VKlass.");
Console.WriteLine("Programmet har öppnat ett nytt Chrome-fönster. Logga in på VKlass och gå till sidan \"Min Klass\".");
Console.WriteLine("När programmet har kört klart kommer frånvaroöversikten ligga i mappen \"VKlass-frånvaro\" på skrivbordet.");
Console.WriteLine();
Console.WriteLine("Kom tillbaka till detta programfönster och tryck på tangenten \"Enter\" när du öppnat sidan \"Min klass\"");
Console.ReadLine();
Console.WriteLine("Börjar scanna efter elever...");

WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

// Set up savefolder
string resourcesFolder = AppDomain.CurrentDomain.BaseDirectory + "\\Resources\\";
string saveFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\VKlass-frånvaro\\" + DateTime.Now.ToString("yyyy-MM-dd HH.mm") + "\\";

if (!Directory.Exists(saveFolder))
{
    Directory.CreateDirectory(saveFolder);
}

// Get links to all student pages
var resultatsidor = wait.Until(e => e.FindElements(By.LinkText("Info & resultat")));
var baseURL = driver.Url;
var resultLinkList = new List<string>();

foreach (var item in resultatsidor)
{
    string resultURL = item.GetAttribute("href");
    resultLinkList.Add(resultURL);
}


var overviewList = new List<AttendanceData>();
string studentOverview = File.ReadAllText(resourcesFolder + "student-overview.html");
var studentOverviewHTMLList = new List<string>();
var startTime = DateTime.Now;

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
    //Thread.Sleep(2000);
    wait.Until(e => e.FindElement(By.LinkText("Visa " + currentMonthName)));
    
    var previousMonthsLessons = GetLessonsOnCurrentPage();
    var allLessons = lessonList.Concat(previousMonthsLessons).DistinctBy(x => x.StartTime).ToList();

    // Save overview data
    studentAttendance.Lessons = allLessons;
    overviewList.Add(studentAttendance);

    // Create a folder for this students graphical overview
    var studentFolder = saveFolder + studentAttendance.Name + "\\";
    if (!Directory.Exists(studentFolder))
    {
        Directory.CreateDirectory(studentFolder);
    }

    // Get students graphical overview and save it in the students folder
    var studentOverviewHTML = GetHTMLAttendanceOverview(studentAttendance);
    File.WriteAllText(studentFolder + "overview.html", studentOverview.Replace("%%STUDENT%%", studentAttendance.Name).Replace("%%BODY%%", studentOverviewHTML));
    File.Copy(resourcesFolder + "student-overview.css", studentFolder + "student-overview.css", true);

    // Save student overview data for the "Hela Klassen" file
    studentOverviewHTMLList.Add(studentOverviewHTML);


    numStudents++;
    // Break if maxstudents have been reached, only used for testing purposes
    if (numStudents >= maxStudents)
    {
        break;
    }
    // Print estimated runtime after one student has been processed
    else if (numStudents == 1)
    {
        var timeForOneStudent = DateTime.Now;
        var estimatedTime = (timeForOneStudent - startTime) * resultLinkList.Count;
        var finishTime = startTime + estimatedTime;
        Console.WriteLine($"Uppskattad körtid: {(int)estimatedTime.TotalMinutes} minuter och {estimatedTime.Seconds} sekunder");
        Console.WriteLine("Uppskattad tid när programmet har kört klart: " + finishTime.ToString("HH:mm"));
    }
}

// Create rows for the table in the overview.html file
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

// Save overview.html + necessary CSS and JS files
var overviewTemplate = File.ReadAllText(resourcesFolder + "overview.html");
var newOverview = overviewTemplate.Replace("%%OVERVIEW-DATA%%", overviewTableRows);
newOverview = newOverview.Replace("%RUBRIK%", "Frånvaroinformation hämtad från VKlass " + DateTime.Now.ToString("HH:mm dd") +"/"+DateTime.Now.ToString("MM-yy"));
File.WriteAllText(saveFolder + "overview.html", newOverview);
File.Copy(resourcesFolder + "style.css", saveFolder + "style.css", true);
File.Copy(resourcesFolder + "sort-table.min.js", saveFolder + "sort-table.min.js", true);

// Save a file containing every students graphical overview
if (!Directory.Exists(saveFolder + "HelaKlassen\\"))
{
    Directory.CreateDirectory(saveFolder + "HelaKlassen\\");
}
File.WriteAllText(saveFolder + "HelaKlassen\\klassen.html", studentOverview.Replace("%%STUDENT%%", "Hela klassen").Replace("%%BODY%%", string.Join("",studentOverviewHTMLList)));
File.Copy(resourcesFolder + "student-overview.css", saveFolder + "HelaKlassen\\student-overview.css", true);

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

    // Get information from JavaScript variables on the page
    var jsExec = (IJavaScriptExecutor)driver;
    var radControls = jsExec.ExecuteScript(@"
let answerArray = Array(0)
window.TelerikCommonScripts.radControls.forEach(x => {if (x.hasOwnProperty(""_text"") && x._text.includes(""Kurs : "")){answerArray.push(x._text + '???' + x._targetControlID)}})
return answerArray
");

    var radControlsResult = ((System.Collections.ObjectModel.ReadOnlyCollection<object>)radControls).Select(x => x.ToString()).ToList();

    var lessonList = new List<Lesson>();

    // Create a lesson based on information found in the JavaScript variables as well as the date in the HTML view
    foreach (var radControlItem in radControlsResult)
    {
        var splitRadControlItem = radControlItem.Split("???");
        var tooltip = splitRadControlItem[0];
        var lessonHTMLID = splitRadControlItem[1];

        // Get day from ID
        var lesson = driver.FindElement(By.Id(lessonHTMLID));
        var par = lesson.FindElement(By.XPath("./.."));
        var pDay = par.FindElement(By.ClassName("DayDate"));
        var day = pDay.Text;

        // Get everything else from string found in JavaScript
        var splitTooltip = tooltip.Split("<br />",StringSplitOptions.TrimEntries);
        string clockInfo = "";
        string course = "";
        string status = "";
        int missingMinutes = 0;
        // Parse the data found in the string
        foreach (var line in splitTooltip)
        {
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

        // Create a lesson based on the information that was found
        var newLesson = new Lesson();
        newLesson.StartTime = new DateTime(int.Parse(currentYear), months[currentMonthName], int.Parse(day), int.Parse(clockInfo[..2]), int.Parse(clockInfo[3..5]), 0);
        newLesson.StopTime = new DateTime(int.Parse(currentYear), months[currentMonthName], int.Parse(day), int.Parse(clockInfo[8..10]), int.Parse(clockInfo[11..13]), 0);
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
    // Sort all lessons by starttime
    attendance.Lessons.Sort((x,y) => x.StartTime.CompareTo(y.StartTime));

    var time = attendance.Lessons.First().StartTime;
    var stoptime = attendance.Lessons.Last().StopTime;
    
    // Start with table headers
    string res = "";
    res += "<table class=\"overviewtable\">\n" +
        "\t<tr><th width=\"50\">Vecka</th><th>Måndag</th><th>Tisdag</th><th>Onsdag</th><th>Torsdag</th><th>Fredag</th></tr>\n";

    // Loop through all lessons based on starttime, beginning with the first lesson and then looping 7 days ahead each run
    for (DateTime dt = time; dt <= stoptime; dt += TimeSpan.FromDays(7))
    {
        int showWeek = ISOWeek.GetWeekOfYear(dt);
        res += $"\t<tr><td class=\"week\">{showWeek}</td>";

        // Loop through Monday - Friday for this week
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

