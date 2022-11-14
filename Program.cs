using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System.Globalization;


int maxStudents = 1000;
int numStudents = 0;


// Initialize and open the ChromeDriver
ChromeOptions options = new ChromeOptions();
options.AddArguments("start-maximized");
options.AddArguments("--log-level=3");
ChromeDriverService service = ChromeDriverService.CreateDefaultService();
service.SuppressInitialDiagnosticInformation = true;
service.HideCommandPromptWindow = true;
IWebDriver driver = new ChromeDriver(service, options);

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
    var actions = new Actions(driver);
    actions.ScrollByAmount(0, 200).Perform();

    // Get name of previous month
    var previousLink = wait.Until(e => e.FindElement(By.Id("ctl00_ContentPlaceHolder2_AttandanceOverviewControl_PreviousMonth")));
    var previousMonthName = previousLink.Text.Split()[1];

    // Get name of current month
    var currentMonthSpan = wait.Until(e => e.FindElement(By.XPath("//a/following-sibling::span")));
    var currentMonthName = currentMonthSpan.Text.Split()[0];

    
    // Take a screenshot and save it in a folder on the desktop
    Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
    ss.SaveAsFile($"{saveFolder + name + "-" + currentMonthName}.png", ScreenshotImageFormat.Png);

    // Take a screenshot of previous month
    previousLink.Click();
    Thread.Sleep(3000);
    Screenshot ssPrev = ((ITakesScreenshot)driver).GetScreenshot();
    ssPrev.SaveAsFile($"{saveFolder + name + "-" + previousMonthName}.png", ScreenshotImageFormat.Png);

    // Save overview data
    overviewList.Add(studentAttendance);
    //Console.WriteLine($"{studentAttendance.Name}\t{studentAttendance.Attendance}\t{studentAttendance.ValidAbsence}\t{studentAttendance.InvalidAbsence}");
    
    
    numStudents++;
    if (numStudents >= maxStudents)
    {
        break;
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
    overviewTableRows += $"<tr class=\"{rowClass}\"><td>{item.Name}</td><td>{item.Attendance.ToString(CultureInfo.CreateSpecificCulture("en-US"))}</td><td>{Math.Round(100 - item.Attendance, 2).ToString(CultureInfo.CreateSpecificCulture("en-US"))}</td><td>{item.ValidAbsence.ToString(CultureInfo.CreateSpecificCulture("en-US"))}</td><td>{item.InvalidAbsence.ToString(CultureInfo.CreateSpecificCulture("en-US"))}</td></tr>\n";
}


var overviewTemplate = File.ReadAllText("overview.html");
var newOverview = overviewTemplate.Replace("%%OVERVIEW-DATA%%", overviewTableRows);
newOverview = newOverview.Replace("%RUBRIK%", "Information hämtad " + DateTime.Now.ToString("yy-MM-dd HH:mm"));
File.WriteAllText(saveFolder + "overview.html", newOverview);
File.Copy("style.css", saveFolder + "style.css", true);
File.Copy("sort-table.min.js", saveFolder + "sort-table.min.js", true);

Console.WriteLine("Programmet avslutas...");
driver.Quit();

class AttendanceData
{
    public string Name { get; set; }
    public double Attendance { get; set; }
    public double ValidAbsence { get; set; }
    public double InvalidAbsence { get; set; }
}