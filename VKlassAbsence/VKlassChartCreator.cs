using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Support;
using System.Globalization;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic;

namespace VKlassAbsence
{
    public class VKlassChartCreator
    {
        Dictionary<string, int> months = new Dictionary<string, int>();

        IWebDriver driver;
        Actions actions;
        ChromeDriverService service;
        WebDriverWait wait;
        WebDriverWait shortWait;

        // Change maxStudents for testing purposes to get an overview after only a few students have been processed
        int maxStudents = 10000;
        int numStudents = 0;
        string resourcesFolder = "";
        string saveFolder;

        bool debugging;

        public bool Debugging { get => debugging; set => debugging = value; }

        public VKlassChartCreator()
        {
            InitializeValues();
        }

        private void InitializeValues()
        {
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
        }

        private void InitializeChromeDriver()
        {
            // Initialize and open the ChromeDriver
            ChromeOptions options = new ChromeOptions();
            options.AddArguments("start-maximized");
            options.AddArguments("--log-level=3");
            service = ChromeDriverService.CreateDefaultService();
            service.SuppressInitialDiagnosticInformation = true;
            service.HideCommandPromptWindow = true;
            driver = new ChromeDriver(service, options);
            actions = new Actions(driver);
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            shortWait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
        }

        public void StartChromeWindow()
        {
            InitializeChromeDriver();
            driver.Navigate().GoToUrl("https://auth.vklass.se/organisation/189");
        }

        private void SetupSaveFolder()
        {
            resourcesFolder = AppDomain.CurrentDomain.BaseDirectory + "\\Resources\\";
            saveFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\VKlass-frånvaro\\" + DateTime.Now.ToString("yyyy-MM-dd HH.mm") + "\\";

            if (!Directory.Exists(saveFolder))
            {
                Directory.CreateDirectory(saveFolder);
            }
        }

        public async Task GetAbsenceDataFromClass(IProgress<AbsenceProgress> progress, DateTime? startDate, DateTime? endDate)
        {
            driver.Navigate().GoToUrl("https://www.vklass.se/classlist.aspx");

            // Make sure save folder exists
            SetupSaveFolder();            

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

            progress.Report(new AbsenceProgress() { TotalStudents = resultLinkList.Count });

            // Calculate which months we need to collect data from
            bool allMonthsNeededCollected = false;

            // Is the last date available on the month before its "real" month?            
            var endDateCopy = endDate.Value;
            if (endDateCopy.DayOfWeek == DayOfWeek.Sunday)
            {
                endDateCopy = endDateCopy.AddDays(-1);
            }
            if (endDateCopy.DayOfWeek > (DayOfWeek)1)
            {
                while (endDateCopy.DayOfWeek != DayOfWeek.Monday)
                {
                    endDateCopy = endDateCopy.AddDays(-1);
                }
            }

            // Check the same for the startDate
            // TODO!!!!!!!
            var startDateCopy = startDate.Value;
            if (startDateCopy.DayOfWeek < (DayOfWeek)5 && startDateCopy.DayOfWeek > (DayOfWeek)0)
            {
                while (startDateCopy.DayOfWeek != DayOfWeek.Friday)
                {
                    startDateCopy = startDateCopy.AddDays(1);
                }
            }
            

            foreach (var item in resultLinkList)
            {
                // Go to the students "Info & närvaro" page
                driver.Navigate().GoToUrl(item);

                // Go to "Närvaro"
                IWebElement närvaroLink = wait.Until(e => e.FindElement(By.LinkText("Närvaro")));
                driver.Navigate().GoToUrl(närvaroLink.GetAttribute("href"));

                var swedishCulture = new CultureInfo("sv-SE");

                // Get attendance data for last 30 days
                var studentAttendance = new AttendanceData();
                var overviewElement = wait.Until(e => e.FindElement(By.Id("ctl00_ContentPlaceHolder2_attendanceMinutesLabel")));
                var overviewString = overviewElement.Text;
                var overviewRows = overviewString.Split("\r\n");
                studentAttendance.Attendance = double.Parse(overviewRows[0].Split()[0], swedishCulture);
                var secondRow = overviewRows[1].Split();
                studentAttendance.ValidAbsence = double.Parse(secondRow[3], swedishCulture);
                studentAttendance.InvalidAbsence = double.Parse(secondRow[7], swedishCulture);

                // Find "Månadsvy"-link and click on it
                var månadsvyLink = wait.Until(e => e.FindElement(By.XPath("//span[text()='Månadsvy']")));
                månadsvyLink.Click();

                // Find student name
                var nameLink = wait.Until(e => e.FindElement(By.Id("ctl00_ContentPlaceHolder2_studentLink")));
                var name = nameLink.Text;
                studentAttendance.Name = name;

                bool allDataGathered = false;
                List<Lesson> allLessons = new List<Lesson>();
                bool dataGatheringHasBegun = false;

                while (!allDataGathered)
                {
                    // Get name of previous month
                    var previousLink = wait.Until(e => e.FindElement(By.Id("ctl00_ContentPlaceHolder2_AttandanceOverviewControl_PreviousMonth")));
                    var previousMonthName = previousLink.Text.Split()[1];

                    // Get name of current month
                    var currentMonthSpan = wait.Until(e => e.FindElement(By.XPath("//a/following-sibling::span")));
                    var currentMonthName = currentMonthSpan.Text.Split()[0];
                    var currentYear = currentMonthSpan.Text.Split()[1];

                    if (int.Parse(currentYear) == endDateCopy.Year && months[currentMonthName] == endDateCopy.Month)
                    {
                        dataGatheringHasBegun = true;
                    }

                    // Get all lessons in current month
                    if (dataGatheringHasBegun)
                    {
                        var lessonList = GetLessonsOnCurrentPage();
                        allLessons = allLessons.Concat(lessonList).DistinctBy(x => x.StartTime).ToList();
                    }

                    if (int.Parse(currentYear) == startDateCopy.Year && months[currentMonthName] == startDateCopy.Month)
                    {
                        // All data gathered
                        allDataGathered = true;
                    }
                    else
                    {
                        // Click previous month link from previous month
                        previousLink.Click();

                        // Wait for previous month to load
                        bool changeDetected = false;
                        var prevLinkText = previousLink.Text;
                        int n = 0;

                        while (!changeDetected)
                        {
                            var currentPrevLink = driver.FindElement(By.Id("ctl00_ContentPlaceHolder2_AttandanceOverviewControl_PreviousMonth"));
                            if (prevLinkText != currentPrevLink.Text)
                            {
                                changeDetected = true;
                                break;
                            }
                            n++;

                            if (debugging)
                            {
                                Console.WriteLine("Run " + n + " failed after trying to load previous month");
                            }

                            if (n == 10)
                            {
                                if (debugging)
                                {
                                    Console.WriteLine("Problem med att ladda föregående månad, försöker igen.");
                                }

                                previousLink.Click();
                            }
                            else if (n == 40)
                            {
                                if (debugging)
                                {
                                    Console.WriteLine("Lyckas inte ladda föregående månad.");
                                }
                            }
                            Thread.Sleep(1000);
                        }
                    }
                }    


                // Only select lessons that are between the start and end dates.
                allLessons = allLessons.Where(x => x.StartTime > startDate && x.StartTime < (endDate + TimeSpan.FromHours(24))).ToList();

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

                progress.Report(new AbsenceProgress() {FinishedStudents = numStudents, TotalStudents = resultLinkList.Count });
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
            newOverview = newOverview.Replace("%RUBRIK%", "Frånvaroinformation hämtad från VKlass " + DateTime.Now.ToString("HH:mm dd") + "/" + DateTime.Now.ToString("MM-yy"));
            File.WriteAllText(saveFolder + "overview.html", newOverview);
            File.Copy(resourcesFolder + "style.css", saveFolder + "style.css", true);
            File.Copy(resourcesFolder + "sort-table.min.js", saveFolder + "sort-table.min.js", true);

            // Save a file containing every students graphical overview
            if (!Directory.Exists(saveFolder + "HelaKlassen\\"))
            {
                Directory.CreateDirectory(saveFolder + "HelaKlassen\\");
            }
            File.WriteAllText(saveFolder + "HelaKlassen\\klassen.html", studentOverview.Replace("%%STUDENT%%", "Hela klassen").Replace("%%BODY%%", string.Join("", studentOverviewHTMLList)));
            File.Copy(resourcesFolder + "student-overview.css", saveFolder + "HelaKlassen\\student-overview.css", true);

            Console.WriteLine("Programmet avslutas...");
            driver.Quit();
            // One more than full = Run is complete
            progress.Report(new AbsenceProgress() { FinishedStudents = resultLinkList.Count + 1, TotalStudents = resultLinkList.Count });
        }


        public async Task GetAbsenceDataFromClassByListOverview(IProgress<AbsenceProgress> progress, DateTime? startDate, DateTime? endDate)
        {
            driver.Navigate().GoToUrl("https://www.vklass.se/classlist.aspx");

            // Make sure save folder exists
            SetupSaveFolder();

            // Get links to all student pages
            var resultatsidor = wait.Until(e => e.FindElements(By.LinkText("Info & resultat")));
            var baseURL = driver.Url;
            var resultLinkList = new List<string>();

            foreach (var item in resultatsidor)
            {
                string resultURL = item.GetAttribute("href");
                resultLinkList.Add(resultURL.Replace("StudentResult", "UserAttendance"));
            }

            var overviewList = new List<AttendanceData>();
            string studentOverview = File.ReadAllText(resourcesFolder + "student-overview.html");
            var studentOverviewHTMLList = new List<string>();
            var startTime = DateTime.Now;

            progress.Report(new AbsenceProgress() { TotalStudents = resultLinkList.Count });

            // Is the last date available on the month before its "real" month?            
            var endDateCopy = endDate.Value;

            // Check the same for the startDate
            var startDateCopy = startDate.Value;


            foreach (var item in resultLinkList)
            {
                bool pageLoadedCorrectly = false;
                int numberOfTriesToLoadPage = 0;
                int maxTries = 5;

                var swedishCulture = new CultureInfo("sv-SE");

                var studentAttendance = new AttendanceData();


                while (!pageLoadedCorrectly && numberOfTriesToLoadPage < maxTries)
                {

                    // Go to the students "Info & närvaro" page
                    driver.Navigate().GoToUrl(item);
                    // Find "Hantera Närvaro"-link and click on it
                    var månadsvyLink = wait.Until(e => e.FindElement(By.XPath("//span[text()='Hantera närvaro']")));
                    månadsvyLink.Click();

                    // Find student name
                    var nameLink = wait.Until(e => e.FindElement(By.Id("ctl00_ContentPlaceHolder2_studentLink")));
                    var name = nameLink.Text;
                    studentAttendance.Name = name;

                    List<Lesson> allLessons = new List<Lesson>();

                    // Just fill dates in boxes instead?
                    var jsExec = (IJavaScriptExecutor)driver;
                    string testScript =
    @$"let dateBox = document.getElementById(""ctl00_ContentPlaceHolder2_StartDatePresence_dateInput"");
dateBox.value = """";
dateBox.focus();
document.execCommand(""insertText"", false, ""{(startDateCopy.Subtract(TimeSpan.FromDays(47))).ToString("yyyy-MM-dd")} 02:02"");
dateBox.value = """";
dateBox.focus();
document.execCommand(""insertText"", false, ""{startDateCopy.ToString("yyyy-MM-dd")} 01:01"");
dateBox.dispatchEvent(new Event('change'));
let dateBox2 = document.getElementById(""ctl00_ContentPlaceHolder2_EndDatePresence_dateInput"");
dateBox2.value = """";
dateBox2.focus();
document.execCommand(""insertText"", false, ""{(endDateCopy.Subtract(TimeSpan.FromDays(47))).ToString("yyyy-MM-dd")} 21:57"");
dateBox2.value = """";
dateBox2.focus();
document.execCommand(""insertText"", false, ""{endDateCopy.ToString("yyyy-MM-dd")} 22:58"");
dateBox.focus();
dateBox2.focus();
";
                    jsExec.ExecuteScript(testScript);

                    // Show lessons in the selected timespan
                    var showListButton = wait.Until(e => e.FindElement(By.Id("ctl00_ContentPlaceHolder2_ShowPresenceListButton")));
                    var radControls = jsExec.ExecuteScript(@"arguments[0].value = 'Visa ';", showListButton);
                    showListButton.Click();

                    // Wait until page is updated
                    try
                    {
                        shortWait.Until(e => e.FindElement(By.Id("ctl00_ContentPlaceHolder2_ShowPresenceListButton")).GetAttribute("value") == "Visa");

                    }
                    catch (Exception e)
                    {
                        numberOfTriesToLoadPage++;
                        if (numberOfTriesToLoadPage >= maxTries)
                        {
                            throw;
                        }
                        continue;
                    }
                    pageLoadedCorrectly = true;
                }


                /*
                // Get attendance data for last 30 days                
                var overviewElement = wait.Until(e => e.FindElement(By.Id("ctl00_ContentPlaceHolder2_attendanceMinutesLabel")));
                var overviewString = overviewElement.Text;
                var overviewRows = overviewString.Split("\r\n");
                studentAttendance.Attendance = double.Parse(overviewRows[0].Split()[0], swedishCulture);
                var secondRow = overviewRows[1].Split();
                studentAttendance.ValidAbsence = double.Parse(secondRow[3], swedishCulture);
                studentAttendance.InvalidAbsence = double.Parse(secondRow[7], swedishCulture);
                */


                var presenceList = wait.Until(e => e.FindElement(By.Id("presenceList")));
                var presenceListTBody = wait.Until(e => presenceList.FindElement(By.TagName("tbody")));

                string bodyText = presenceListTBody.GetAttribute("innerHTML").ReplaceLineEndings("");
                //bodyText += "";
                var matchingRows = Regex.Matches(bodyText, @"<tr>(?<row>.*?)<\/tr>");

                foreach (Match row in matchingRows)
                {
                    string lessonRowPattern =
                        @"<td>.*?<\/td>.*?" +
                        @"<td>.*?<a.*?>(?<timeInfo>.*?)</a>.*?<\/td>.*?" +
                        @"<td>(?<subject>.*?)<\/td>.*?" +
                        @"<td.*?<input.*?>(?<status>.*?)<\/td>.*?" +
                        @"<td.*?>(?<valid>.*?)<\/td>.*?" +
                        @"<td.*?>(?<invalid>.*?)<\/td>.*?";

                    var matchingCells = Regex.Match(row.Groups["row"].Value, lessonRowPattern);
                    foreach (Group matchGrp in matchingCells.Groups)
                    {
                        string test = matchGrp.Name + "" + matchGrp.Value;
                        test += "";
                    }

                    var newLesson = new Lesson();
                    newLesson.Course = matchingCells.Groups["subject"].Value;
                    newLesson.MissingMinutes = int.Parse(matchingCells.Groups["invalid"].Value.Split()[0]);
                    newLesson.MissingValidMinutes = int.Parse(matchingCells.Groups["valid"].Value.Split()[0]);
                    var status = matchingCells.Groups["status"].Value;
                    newLesson.Status = status == "Närvarande" ? LessonStatus.Närvarande : status.Contains("Giltigt") ? LessonStatus.GiltigFrånvaro : status.Contains("Ej") ? LessonStatus.EjRapporterat : LessonStatus.OgiltigFrånvaro;

                    // Get start and stoptimes
                    var splitDateInfo = matchingCells.Groups["timeInfo"].Value.Split();
                    var startClock = splitDateInfo[4].Split(":");
                    var endClock = splitDateInfo[6].Split(":");
                    var yearAndMonth = splitDateInfo[2].Split("-");
                    var monthForLesson = months[months.Where(kv => kv.Key[..3].ToLower() == yearAndMonth[0]).First().Key];

                    newLesson.StartTime = new DateTime(int.Parse("20" + yearAndMonth[1]), monthForLesson, int.Parse(splitDateInfo[1]), int.Parse(startClock[0]), int.Parse(startClock[1]), 0);
                    newLesson.StopTime = new DateTime(int.Parse("20" + yearAndMonth[1]), monthForLesson, int.Parse(splitDateInfo[1]), int.Parse(endClock[0]), int.Parse(endClock[1]), 0);

                    
                    studentAttendance.Lessons.Add(newLesson);
                    progress.Report(new AbsenceProgress() { FinishedStudents = numStudents, TotalStudents = resultLinkList.Count });
                }

                // Calculate students attendance, valid absence and invalid absence for the time interval chosen
                long totalLessonTime = (long)studentAttendance.Lessons.Where(x => x.Status != LessonStatus.EjRapporterat).Sum(x => (x.StopTime - x.StartTime).TotalMinutes);
               
                long validAbsenceTimeFromFullLessons = (long)studentAttendance.Lessons.Where(x => x.Status == LessonStatus.GiltigFrånvaro).Sum(x => (x.StopTime - x.StartTime).TotalMinutes);
                long invalidAbsenceTimeFromFullLessons = (long)studentAttendance.Lessons.Where(x => x.Status == LessonStatus.OgiltigFrånvaro).Sum(x => (x.StopTime - x.StartTime).TotalMinutes);
                
                long validMissingMinutes = studentAttendance.Lessons.Where(x => x.Status == LessonStatus.Närvarande).Sum(x => x.MissingValidMinutes);
                long invalidMissingMinutes = studentAttendance.Lessons.Where(x => x.Status == LessonStatus.Närvarande).Sum(x => x.MissingMinutes);

                long totalInvalidMinutes = invalidAbsenceTimeFromFullLessons + invalidMissingMinutes;
                long totalValidMinutes = validAbsenceTimeFromFullLessons + validMissingMinutes;
                long totalAbsence = totalInvalidMinutes + totalValidMinutes;

                studentAttendance.Attendance = Math.Round((totalLessonTime - totalAbsence) * 100 / (double)totalLessonTime, 1);
                studentAttendance.ValidAbsence = Math.Round((totalValidMinutes * 100 / (double)totalLessonTime), 1);
                studentAttendance.InvalidAbsence = Math.Round(totalInvalidMinutes * 100/ (double)totalLessonTime, 1);
                

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
                progress.Report(new AbsenceProgress() { FinishedStudents = numStudents, TotalStudents = resultLinkList.Count });

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
            newOverview = newOverview.Replace("de senaste 30 dagarna", $"perioden {startDateCopy.ToString("yyyy-MM-dd")} till {endDateCopy.ToString("yyyy-MM-dd")}");
            newOverview = newOverview.Replace("%RUBRIK%", "Frånvaroinformation hämtad från VKlass " + DateTime.Now.ToString("HH:mm dd") + "/" + DateTime.Now.ToString("MM-yy"));
            File.WriteAllText(saveFolder + "overview.html", newOverview);
            File.Copy(resourcesFolder + "style.css", saveFolder + "style.css", true);
            File.Copy(resourcesFolder + "sort-table.min.js", saveFolder + "sort-table.min.js", true);

            // Save a file containing every students graphical overview
            if (!Directory.Exists(saveFolder + "HelaKlassen\\"))
            {
                Directory.CreateDirectory(saveFolder + "HelaKlassen\\");
            }
            File.WriteAllText(saveFolder + "HelaKlassen\\klassen.html", studentOverview.Replace("%%STUDENT%%", "Hela klassen").Replace("%%BODY%%", string.Join("", studentOverviewHTMLList)));
            File.Copy(resourcesFolder + "student-overview.css", saveFolder + "HelaKlassen\\student-overview.css", true);

            Console.WriteLine("Programmet avslutas...");
            driver.Quit();
            // One more than full = Run is complete
            progress.Report(new AbsenceProgress() { FinishedStudents = resultLinkList.Count + 1, TotalStudents = resultLinkList.Count, PathToOverview = saveFolder + "overview.html" });
        }


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
                var splitTooltip = tooltip.Split("<br />", StringSplitOptions.TrimEntries);
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


                // IS THE ROW IN THE SECOND ROW IN TBODY? THEN IT !!!!MIGHT!!!! BELONG TO PREVIOUSMONTH
                // IS IT IN THE LAST ROW IN TBODY? THEN IT !!!MIGHT!!! BELONG TO THE NEXT MONTH.
                // CHECK DAY NUMBER (less than or larger than 15)
                var month = months[currentMonthName];
                var firstWeek = par.FindElement(By.XPath("./../../tr[position()=2]/td[1]"));
                var currentWeekRow = par.FindElement(By.XPath("./../td[1]"));
                var lastWeekRow = par.FindElement(By.XPath("./../../tr[last()]/td[1]"));
                int year = int.Parse(currentYear);

                // Does the lesson belong to the previous month?
                if (currentWeekRow.Text == firstWeek.Text && int.Parse(day) > 15)
                {
                    month--;
                    if (month <= 0)
                    {
                        year--;
                        month = 12;
                    }

                }
                // Does the lesson belong to the next month?
                if (lastWeekRow.Text == currentWeekRow.Text && int.Parse(day) < 15)
                {
                    month++;
                    if (month > 12)
                    {
                        year++;
                        month = 1;
                    }
                }

                // Create a lesson based on the information that was found
                var newLesson = new Lesson();

                newLesson.StartTime = new DateTime(year, month, int.Parse(day), int.Parse(clockInfo[..2]), int.Parse(clockInfo[3..5]), 0);
                newLesson.StopTime = new DateTime(year, month, int.Parse(day), int.Parse(clockInfo[8..10]), int.Parse(clockInfo[11..13]), 0);
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
            attendance.Lessons.Sort((x, y) => x.StartTime.CompareTo(y.StartTime));



            // Start with table headers
            string res = "";
            res += "<table class=\"overviewtable\">\n" +
                "\t<tr><th width=\"50\">Vecka</th><th>Måndag</th><th>Tisdag</th><th>Onsdag</th><th>Torsdag</th><th>Fredag</th></tr>\n";

            if (attendance.Lessons.Count > 0)
            {
                var time = attendance.Lessons.First().StartTime;
                var stoptime = attendance.Lessons.Last().StopTime;

                // Loop through all lessons based on starttime, beginning with the first lesson and then looping 7 days ahead each run

                // BUT!!!!! WHAT IF THE FIRST LESSON IS NOT ON A MONDAY? TODO: FIX IT!
                // dt needs to start at early morning on the monday in the same week as the first lesson, that is the variable "time".
                time = ISOWeek.ToDateTime(time.Year, ISOWeek.GetWeekOfYear(time), DayOfWeek.Monday) + TimeSpan.FromHours(3);

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
                            int fractionLate = (int)Math.Round((((lesson.MissingMinutes + lesson.MissingValidMinutes) / lessonLength) * 100));
                            int fractionValidLate = (int)Math.Round((((lesson.MissingValidMinutes) / lessonLength) * 100));
                            res += $"<div class=\"{lesson.Status} lesson\"><div class=\"background-overlay{(fractionLate > 0 ? " show-late" : "")}\" style=\"--late: {fractionLate}%; --valid-late: {fractionValidLate}%\"><div class=\"coursename\">{lesson.Course}</div>";
                            res += $"</div></div>\n";
                        }
                        res += "</td>";
                    }
                    res += "</tr>\n";
                }
                res += "</table>\n";
            }
         

            return template.Replace("%%STUDENT%%", attendance.Name).Replace("%%CONTENT%%", res);
            
        }

    }
}