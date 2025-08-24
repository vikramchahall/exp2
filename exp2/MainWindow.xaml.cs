using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using System.IO.IsolatedStorage;
using System.Timers;
using Microsoft.Win32; // For file dialog
using System.Data;
using System.Windows.Controls;
using System.Windows.Input;
using Forms = System.Windows.Forms;
using Newtonsoft.Json.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TrayNotify;
using System.Windows.Media.Media3D;
using System.Windows.Media;
using System.Xml.Linq;
using System.IO.Packaging;
using System.Diagnostics;
using System.Globalization;
using System.Security.Cryptography;
using static OfficeOpenXml.ExcelErrorValue;
using System.Windows.Controls.Primitives;

namespace MessManagementSystem
{
    public partial class MainWindow : Window
    {
        private static readonly string[] Scopes = { DriveService.Scope.DriveReadonly, SheetsService.Scope.SpreadsheetsReadonly, SheetsService.Scope.Spreadsheets };
        private static readonly string ApplicationName = "GoogleDriveApp";
        private static readonly string ServiceAccountKeyFilePath = @"C:\Users\vikra\source\repos\exp2\exp2\bin\Debug\service-account-file.json";
        private string uploadedSpreadsheetPath;
        private string uploadedImageFolderPath;
        private string spreadsheetId;
        private DateTime lastSyncTimeExcel;
        private DateTime lastSyncTimeSheets;
        private Timer syncTimer;
        private FileSystemWatcher excelWatcher;
        private string excelFilePath;
        private List<string> imageFilePaths;
        private string ratesExcelFilePath;
        private List<StudentData> studentDataList;
        private Dictionary<string, Dictionary<string, double>> ratesData;
        private Dictionary<string, Dictionary<string, double>> extraMealOptionsData; // Sheet 3
        private DateTime lastRollNumberEntryTime;
        private Timer extraMealSelectionTimer;
        private string lastSelectedOption = null;
        private Dictionary<string, Dictionary<string, double>> extraMealsRatesData;
        private bool isLocked = false;
        private string pin = "1234";

        protected override void OnActivated(EventArgs e)
        {
            base.OnActivated(e);
            ForceFocusOnInput();
        }





        [Obsolete]
        public MainWindow()
        {
            InitializeComponent();

            imageFilePaths = new List<string>();

            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            RollNumberTextBox.TextChanged += (s, e) => {
            };
            LoadFilePaths();
            LoadPaths();
            InitializeSyncTimer();
            extraMealSelectionTimer = new Timer(3000); // 3 seconds
            extraMealSelectionTimer.AutoReset = false;
            LoadPastEntries();
            RollNumberTextBoxExtraMeals.TextChanged += RollNumberTextBoxExtraMeals_TextChanged;
            MainTabControl.SelectionChanged += MainTabControl_SelectionChanged;
            ForceFocusOnInput();


        }
        public class InputDialog : Window
        {
            private TextBox inputTextBox;

            public string Input { get; private set; }

            public InputDialog(string prompt, string title)
            {
                Title = title;
                Width = 300;
                Height = 150;
                WindowStartupLocation = WindowStartupLocation.CenterScreen;

                var stackPanel = new StackPanel { Margin = new Thickness(10) };
                Content = stackPanel;

                stackPanel.Children.Add(new TextBlock { Text = prompt, Margin = new Thickness(0, 0, 0, 5) });

                inputTextBox = new TextBox { Margin = new Thickness(0, 0, 0, 5) };
                stackPanel.Children.Add(inputTextBox);

                var buttonPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
                stackPanel.Children.Add(buttonPanel);

                var okButton = new Button { Content = "OK", Width = 60, Margin = new Thickness(0, 0, 5, 0) };
                okButton.Click += (s, e) => { Input = inputTextBox.Text; DialogResult = true; };
                buttonPanel.Children.Add(okButton);

                var cancelButton = new Button { Content = "Cancel", Width = 60 };
                cancelButton.Click += (s, e) => DialogResult = false;
                buttonPanel.Children.Add(cancelButton);
            }
        }
        private void ForceFocusOnInput()
        {
            TextBox textBoxToFocus = null;

            if (MainTabControl.SelectedItem is System.Windows.Controls.TabItem selectedTab)
            {
                if (selectedTab.Name == "Dashboard")
                {
                    textBoxToFocus = RollNumberTextBox;
                }
                else if (selectedTab.Name == "ExtraMeals")
                {
                    textBoxToFocus = RollNumberTextBoxExtraMeals;
                }
            }

            if (textBoxToFocus != null)
            {
                textBoxToFocus.Focus();
                textBoxToFocus.CaretIndex = textBoxToFocus.Text.Length;

                // Use a DispatcherTimer to continuously check and refocus
                DispatcherTimer focusTimer = new DispatcherTimer();
                focusTimer.Interval = TimeSpan.FromMilliseconds(50);
                focusTimer.Tick += (sender, e) =>
                {
                    if (!textBoxToFocus.IsFocused)
                    {
                        textBoxToFocus.Focus();
                        textBoxToFocus.CaretIndex = textBoxToFocus.Text.Length;
                    }
                };
                focusTimer.Start();
            }
        }
        private void LockButton_Click(object sender, RoutedEventArgs e)
        {
            if (isLocked)
            {
                // Prompt for PIN
                var inputDialog = new InputDialog("Enter PIN to unlock:", "Unlock");
                if (inputDialog.ShowDialog() == true)
                {
                    string enteredPin = inputDialog.Input;
                    if (enteredPin == pin)
                    {
                        isLocked = false;
                        LockButton.Content = "🔓";
                        EnableAllTabs();
                    }
                    else
                    {
                        MessageBox.Show("Incorrect PIN");
                    }
                }
            }
            else
            {
                isLocked = true;
                LockButton.Content = "🔒";
                DisableAllTabsExceptSnacks();
            }
        }

        private void DisableAllTabsExceptSnacks()
        {
            foreach (System.Windows.Controls.TabItem tab in MainTabControl.Items)
            {
                if (tab.Name != "Snacks")
                {
                    tab.IsEnabled = false;
                }
            }
        }

        private void EnableAllTabs()
        {
            foreach (System.Windows.Controls.TabItem tab in MainTabControl.Items)
            {
                tab.IsEnabled = true;
            }
        }

        private async void FetchAndSave_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string folderUrl = FolderUrl.Text;

                if (string.IsNullOrEmpty(folderUrl))
                {
                    StatusLabel.Text = "Please provide both the spreadsheet and folder URLs.";
                    return;
                }


                var credential = GoogleCredential.FromFile(ServiceAccountKeyFilePath).CreateScoped(Scopes);

                var sheetsService = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                var driveService = new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });


                await DownloadFilesFromFolderAsync(driveService, folderUrl);

                StatusLabel.Text = "images downloaded successfully adn uploaded to application!";
            }
            catch (Exception ex)
            {
                StatusLabel.Text = $"Error: {ex.Message}";
            }
        }

        private async Task DownloadFilesFromFolderAsync(DriveService driveService, string folderUrl)
        {
            try
            {
                string folderId = ExtractFolderId(folderUrl);

                var request = driveService.Files.List();
                request.Q = $"'{folderId}' in parents";
                var files = (await request.ExecuteAsync()).Files;

                if (files != null && files.Count > 0)
                {
                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string imagesFolder = Path.Combine(desktopPath, "DownloadedImages");
                    Directory.CreateDirectory(imagesFolder);

                    var tasks = files.Select(async file =>
                    {
                        using (var stream = new MemoryStream())
                        {
                            var requestGet = driveService.Files.Get(file.Id);
                            await requestGet.DownloadAsync(stream);

                            // Save the file with .jpg extension regardless of its original type
                            string filePath = Path.Combine(imagesFolder, file.Name.EndsWith(".jpg") ? file.Name : file.Name + ".jpg");
                            using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                            {
                                stream.Position = 0;
                                await stream.CopyToAsync(fileStream);
                            }
                        }
                    });

                    await Task.WhenAll(tasks);

                    // Save the folder path for later use in DataUpload tab
                    uploadedImageFolderPath = imagesFolder;
                    SavePaths();
                }

                StatusLabel.Text = "Files downloaded successfully!";
            }
            catch (Exception ex)
            {
                StatusLabel.Text = $"Error downloading files: {ex.Message}";
            }
        }

        private static string ExtractFolderId(string url)
        {
            var match = Regex.Match(url, @"\/folders\/([a-zA-Z0-9-_]+)");
            if (match.Success)
            {
                return match.Groups[1].Value;
            }
            throw new ArgumentException("Invalid folder URL.");
        }




        private async Task SaveSpreadsheetToExcelAsync(SheetsService sheetsService, string spreadsheetId)
        {
            var request = sheetsService.Spreadsheets.Get(spreadsheetId);
            Spreadsheet spreadsheet = request.Execute();
            var sheet = spreadsheet.Sheets.First(); // Use First() instead of FirstOrDefault()

            if (sheet != null)
            {
                var range = $"{sheet.Properties.Title}!A1:Z1000";
                var valueRequest = sheetsService.Spreadsheets.Values.Get(spreadsheetId, range);
                var response = await valueRequest.ExecuteAsync();
                var values = response.Values;

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add(sheet.Properties.Title);

                    for (int i = 0; i < values.Count; i++)
                    {
                        for (int j = 0; j < values[i].Count; j++)
                        {
                            worksheet.Cells[i + 1, j + 1].Value = values[i][j];
                        }
                    }

                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string spreadsheetFilePath = Path.Combine(desktopPath, "Spreadsheet.xlsx");

                    package.SaveAs(new FileInfo(spreadsheetFilePath));

                    // Save the path for later use in DataUpload tab
                    uploadedSpreadsheetPath = spreadsheetFilePath;
                    UploadedSpreadsheetPath.Text = spreadsheetFilePath;
                    SavePaths();

                    // Initialize FileSystemWatcher for the saved Excel file
                    InitializeExcelWatcher(spreadsheetFilePath);
                }
            }
        }


        private void BrowseSpreadsheet_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                uploadedSpreadsheetPath = openFileDialog.FileName;
                UploadedSpreadsheetPath.Text = uploadedSpreadsheetPath;
                SavePaths();

                // Initialize FileSystemWatcher for the chosen Excel file
                InitializeExcelWatcher(uploadedSpreadsheetPath);
            }
        }


        private void BrowseImageFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var folderDialog = new Forms.FolderBrowserDialog())
            {
                if (folderDialog.ShowDialog() == Forms.DialogResult.OK)
                {
                    uploadedImageFolderPath = folderDialog.SelectedPath;
                    SavePaths();
                }
            }
        }


        private async void UploadFiles_Click(object sender, RoutedEventArgs e)
        {
            await UploadFilesAsync();
        }


        private async Task UploadFilesAsync()
        {
            if (string.IsNullOrEmpty(uploadedSpreadsheetPath) || string.IsNullOrEmpty(uploadedImageFolderPath))
            {
                UploadStatusLabel.Text = "Please upload both the spreadsheet and image folder.";
                return;
            }

            try
            {
                var credential = GoogleCredential.FromFile(ServiceAccountKeyFilePath).CreateScoped(Scopes);
                var sheetsService = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                await UpdateGoogleSheetsFromExcel(sheetsService);

                UploadStatusLabel.Text = "Files uploaded successfully.";
            }
            catch (Exception ex)
            {
                UploadStatusLabel.Text = $"Error: {ex.Message}";
            }
        }


        private void LoadPaths()
        {
            // Load previously saved paths and URLs using Isolated Storage
            var storage = IsolatedStorageFile.GetUserStoreForDomain();
            var pathsFile = "paths.txt";

            if (storage.FileExists(pathsFile))
            {
                using (var stream = new IsolatedStorageFileStream(pathsFile, FileMode.Open, storage))
                using (var reader = new StreamReader(stream))
                {
                    uploadedSpreadsheetPath = reader.ReadLine();
                    uploadedImageFolderPath = reader.ReadLine();
                    spreadsheetId = reader.ReadLine();
                    FolderUrl.Text = reader.ReadLine();

                    UploadedSpreadsheetPath.Text = uploadedSpreadsheetPath;
                }
            }
        }


        private void SavePaths()
        {
            // Save paths and URLs using Isolated Storage
            var storage = IsolatedStorageFile.GetUserStoreForDomain();
            var pathsFile = "paths.txt";

            using (var stream = new IsolatedStorageFileStream(pathsFile, FileMode.Create, storage))
            using (var writer = new StreamWriter(stream))
            {
                writer.WriteLine(uploadedSpreadsheetPath);
                writer.WriteLine(uploadedImageFolderPath);
                writer.WriteLine(spreadsheetId);
                writer.WriteLine(FolderUrl.Text);
            }
        }


        [Obsolete]
        private void InitializeSyncTimer()
        {
            syncTimer = new Timer(60000); // Set to 1 minute
            syncTimer.Elapsed += (sender, e) =>
            {
                Task task = CheckForGoogleSheetsUpdates();
            };
            syncTimer.AutoReset = true;
            syncTimer.Start();
        }


        [Obsolete]
        private async Task CheckForGoogleSheetsUpdates()
        {
            try
            {
                var credential = GoogleCredential.FromFile(ServiceAccountKeyFilePath).CreateScoped(Scopes);
                var sheetsService = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                var driveService = new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                var fileRequest = driveService.Files.Get(spreadsheetId);
                fileRequest.Fields = "modifiedTime";
                var file = await fileRequest.ExecuteAsync();

                if (file != null && DateTime.TryParse(file.ModifiedTime.ToString(), out DateTime modifiedTime))
                {
                    if (modifiedTime > lastSyncTimeSheets)
                    {
                        await Dispatcher.InvokeAsync(() => SaveSpreadsheetToExcelAsync(sheetsService, spreadsheetId));
                        lastSyncTimeSheets = modifiedTime;
                    }
                }
            }
            catch (Exception ex)
            {
                await Dispatcher.InvokeAsync(() => UploadStatusLabel.Text = $"Error checking for updates: {ex.Message}");
            }
        }


        [Obsolete]
        private async void ManualSyncButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                UploadStatusLabel.Text = "Manual sync in progress...";
                await Task.Run(() => CheckForGoogleSheetsUpdates());
                UploadStatusLabel.Text = "Manual sync completed.";
            }
            catch (Exception ex)
            {
                UploadStatusLabel.Text = $"Error during manual sync: {ex.Message}";
            }
        }


        private void InitializeExcelWatcher(string excelFilePath)
        {
            excelWatcher = new FileSystemWatcher
            {
                Path = Path.GetDirectoryName(excelFilePath),
                Filter = Path.GetFileName(excelFilePath),
                NotifyFilter = NotifyFilters.LastWrite,
                EnableRaisingEvents = true
            };

            excelWatcher.Changed += async (sender, e) =>
            {
                try
                {
                    var lastWriteTime = File.GetLastWriteTime(excelFilePath);
                    if (lastWriteTime > lastSyncTimeExcel)
                    {
                        lastSyncTimeExcel = lastWriteTime;

                        var credential = GoogleCredential.FromFile(ServiceAccountKeyFilePath).CreateScoped(Scopes);
                        var sheetsService = new SheetsService(new BaseClientService.Initializer()
                        {
                            HttpClientInitializer = credential,
                            ApplicationName = ApplicationName,
                        });

                        await UpdateGoogleSheetsFromExcel(sheetsService);
                    }
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() => UploadStatusLabel.Text = $"Error uploading Excel updates: {ex.Message}");
                }
            };
        }//idk maybe this updates the excel on updating the sheets


        private async Task UpdateGoogleSheetsFromExcel(SheetsService sheetsService)
        {
            using (var package = new ExcelPackage(new FileInfo(uploadedSpreadsheetPath)))
            {
                var worksheet = package.Workbook.Worksheets.First();
                var rows = worksheet.Dimension.Rows;
                var cols = worksheet.Dimension.Columns;

                var values = new List<IList<object>>();
                for (int row = 1; row <= rows; row++)
                {
                    var valueRow = new List<object>();
                    for (int col = 18; col <= cols; col++) // Start reading from column 18 (R)
                    {
                        valueRow.Add(worksheet.Cells[row, col].Text);
                    }
                    values.Add(valueRow);
                }

                var valueRange = new ValueRange
                {
                    Values = values
                };

                var updateRange = $"{worksheet.Name}!R1"; // Start updating from column 18 (R)
                var updateRequest = sheetsService.Spreadsheets.Values.Update(valueRange, spreadsheetId, updateRange);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                await updateRequest.ExecuteAsync();
            }
        }



        private void LoadFilePaths()
        {
            if (File.Exists("dataFilePath.txt"))
            {
                excelFilePath = File.ReadAllText("dataFilePath.txt");
                ExcelStatusTextBlock.Text = "Data Excel file loaded successfully.";
                ProcessExcelData(); // Process data after loading the file path
                DisplayProcessedData();
            }

            if (File.Exists("ratesFilePath.txt"))
            {
                ratesExcelFilePath = File.ReadAllText("ratesFilePath.txt");
                RatesExcelStatusTextBlock.Text = "Rates Excel file loaded successfully.";
                ProcessRatesData();
                DisplayRatesData();
                PopulateExtraMealOptions();


            }

            if (File.Exists("imageFilePaths.txt"))
            {
                imageFilePaths = File.ReadAllLines("imageFilePaths.txt").ToList();
                ImagesStatusTextBlock.Text = "Images loaded successfully.";
            }
        }


        private void SaveFilePath(string filePath, string fileName)
        {
            File.WriteAllText(fileName, filePath);
        }
        private void SaveFilePaths(List<string> filePaths, string fileName)
        {
            File.WriteAllLines(fileName, filePaths);
        }


        private void UploadExcelButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Title = "Select an Excel File"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                excelFilePath = openFileDialog.FileName;
                ExcelStatusTextBlock.Text = "Excel file uploaded successfully.";
                SaveFilePath(excelFilePath, "dataFilePath.txt");
                ProcessExcelData();
                DisplayProcessedData();
            }
        }


        private void UploadImagesButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp",
                Title = "Select Images",
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() == true)
            {
                imageFilePaths = openFileDialog.FileNames.ToList();
                ImagesStatusTextBlock.Text = "Images uploaded successfully.";
                SaveFilePaths(imageFilePaths, "imageFilePaths.txt");
            }
        }


        private void ProcessDataButton_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(excelFilePath))
            {
                ProcessExcelData();
                ProcessImages();
                DisplayProcessedData();
                MessageBox.Show("Data processing complete.");
            }
            else
            {
                MessageBox.Show("Please upload an Excel file first.");
                return;
            }
        }


        private void ProcessExcelData()
        {
            studentDataList = new List<StudentData>();

            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    var rollNumber = worksheet.Cells[row, 3].Text;
                    var name = worksheet.Cells[row, 2].Text;
                    var fathersname = worksheet.Cells[row, 4].Text;
                    studentDataList.Add(new StudentData { RollNumber = rollNumber, Name = name, Amount = fathersname });
                }


                studentDataList = studentDataList.OrderBy(item => item.RollNumber).ToList();
            }

        }

        private void ProcessImages()
        {
            imageFilePaths = imageFilePaths.OrderBy(path => Path.GetFileNameWithoutExtension(path)).ToList();
        }


        private void DisplayProcessedData()
        {
            DataListBox.Items.Clear();
            int frameSize = 10;
            for (int i = 0; i < studentDataList.Count; i += frameSize)
            {
                var framePanel = new StackPanel { Orientation = Orientation.Vertical, Margin = new Thickness(5) };
                for (int j = i; j < i + frameSize && j < studentDataList.Count; j++)
                {
                    var student = studentDataList[j];
                    var imageFilePath = imageFilePaths.FirstOrDefault(path => Path.GetFileNameWithoutExtension(path) == student.RollNumber);
                    var displayText = $"{student.RollNumber}: {student.Name} - {student.Amount}";

                    if (imageFilePath != null)
                    {
                        var bitmap = new BitmapImage(new Uri(imageFilePath));
                        var image = new Image { Source = bitmap, Width = 50, Height = 50, Margin = new Thickness(5) };
                        framePanel.Children.Add(new StackPanel { Orientation = Orientation.Horizontal, Children = { new TextBlock { Text = displayText }, image } });
                    }
                    else
                    {
                        framePanel.Children.Add(new TextBlock { Text = displayText });
                    }
                }
                DataListBox.Items.Add(framePanel);
            }
        }


        private void UploadRatesExcelButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Title = "Select an Excel File"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                ratesExcelFilePath = openFileDialog.FileName;
                RatesExcelStatusTextBlock.Text = "Rates Excel file uploaded successfully.";
                SaveFilePath(ratesExcelFilePath, "ratesFilePath.txt");
                ProcessRatesData();
                DisplayRatesData();
            }
        }


        private void ProcessRatesData()
        {
            ratesData = new Dictionary<string, Dictionary<string, double>>();
            extraMealsRatesData = new Dictionary<string, Dictionary<string, double>>();
            extraMealOptionsData = new Dictionary<string, Dictionary<string, double>>();

            using (var package = new ExcelPackage(new FileInfo(ratesExcelFilePath)))
            {
                // Process regular meals (Sheet 1)
                var worksheet = package.Workbook.Worksheets[0];
                ProcessRatesDataForSheet(worksheet, ratesData);

                if (package.Workbook.Worksheets.Count > 1)
                {
                    var extraMealsWorksheet = package.Workbook.Worksheets[1];
                    ProcessRatesDataForSheet(extraMealsWorksheet, extraMealsRatesData);
                }
                // Process extra meal options(Sheet 3)
                if (package.Workbook.Worksheets.Count > 2)
                {
                    var extraMealOptionsWorksheet = package.Workbook.Worksheets[2];
                    ProcessRatesDataForSheet(extraMealOptionsWorksheet, extraMealOptionsData);
                }
            }
            DisplayRatesData();
            DisplayExtraMealsDefaultRates();
        }
        private void DisplayExtraMealsDefaultRates()
        {
            ExtraMealsDefaultRatesListBox.Items.Clear();
            foreach (var mealType in extraMealsRatesData.Keys)
            {
                ExtraMealsDefaultRatesListBox.Items.Add(new TextBlock { Text = mealType, FontWeight = FontWeights.Bold });
                foreach (var dayRate in extraMealsRatesData[mealType])
                {
                    ExtraMealsDefaultRatesListBox.Items.Add(new TextBlock { Text = $"{dayRate.Key}: {dayRate.Value:C}" });
                }
            }
        }
        private void ExtraMealsRollNumberTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (ExtraMealsRollNumberTextBox.Text.Length == 8)
            {
                string rollNumber = ExtraMealsRollNumberTextBox.Text;
                DisplayExtraMealStudentDetails(rollNumber);
                ExtraMealsRollNumberTextBox.Clear();
                ExtraMealsRollNumberTextBox.Focus(); // Keep the cursor active for the next entry
            }
        }
        private void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ForceFocusOnInput();
        }
        private void SetFocusToActiveTextBox()
        {
            Dispatcher.InvokeAsync(() =>
            {
                if (MainTabControl.SelectedItem is System.Windows.Controls.TabItem selectedTab)
                {
                    if (selectedTab.Name == "Dashboard")
                    {
                        RollNumberTextBox.Focus();
                        RollNumberTextBox.SelectAll();
                    }
                    else if (selectedTab.Name == "ExtraMeals")
                    {
                        RollNumberTextBoxExtraMeals.Focus();
                        RollNumberTextBoxExtraMeals.SelectAll();
                    }
                }
            }, System.Windows.Threading.DispatcherPriority.Input);
        }


        private void LoadPastEntries()
        {
            DashboardPastEntriesListView.Items.Clear();
            ExtraMealsPastEntriesListView.Items.Clear();
        }
        private double GetExtraMealApplicableRate()
        {
            var dayOfWeek = DateTime.Now.DayOfWeek.ToString();
            var currentTime = DateTime.Now.TimeOfDay;
            var mealType = GetMealType(currentTime);

            if (extraMealsRatesData.ContainsKey(mealType) && extraMealsRatesData[mealType].ContainsKey(dayOfWeek))
            {
                return extraMealsRatesData[mealType][dayOfWeek];
            }
            else
            {
                // If the specific meal type is not found, try to find a default rate
                foreach (var meal in extraMealsRatesData.Keys)
                {
                    if (extraMealsRatesData[meal].ContainsKey(dayOfWeek))
                    {
                        return extraMealsRatesData[meal][dayOfWeek];
                    }
                }
            }

            // If no rate is found, return a default value or throw an exception
           StatusLabel.Text = $"No rate found for {mealType} on {dayOfWeek}. Using default rate of 0.";
            return 0.0;
        }
        private void SaveExtraMealEntry(string rollNumber)
        {
            var student = studentDataList.FirstOrDefault(s => s.RollNumber == rollNumber);
            if (student != null)
            {
                var applicableRate = GetExtraMealApplicableRate();
                var mealType = GetMealType(DateTime.Now.TimeOfDay);
                var date = DateTime.Now.ToString("dd-MM-yyyy");
                var time = DateTime.Now.ToString("HH:mm:ss");

                SaveExtraMealDetails(rollNumber, student.Name, student.Amount, applicableRate, mealType, date, time);
                UpdateExtraMealsPastEntriesListView(rollNumber, student.Name, mealType, applicableRate, date, time);

                MessageBox.Show($"Extra meal saved for {student.Name} ({rollNumber})\nMeal: {mealType}\nRate: {applicableRate:C}");
            }
            else
            {
                MessageBox.Show("Student not found.");
            }
        }

        private void SaveExtraMealDetails(string rollNumber, string name, string amount, double applicableRate, string mealType, string date, string time)
        {
            var logEntry = $"Roll Number: {rollNumber} Name: {name} Amount: {amount} Rate: {applicableRate:C} Meal: (Extra) {mealType}  Date: {date} Time: {time}\n";
            File.AppendAllText("extra_meal_log.txt", logEntry);

            // Update the Excel file
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    if (worksheet.Cells[row, 3].Text == rollNumber)
                    {
                        int col = 18;
                        while (!string.IsNullOrEmpty(worksheet.Cells[row, col].Text))
                        {
                            col++;
                        }
                        worksheet.Cells[row, col].Value = $"{applicableRate:C} - (Extra) {mealType}- {date}";
                        break;
                    }
                }

                package.Save();
            }
        }
        private void UpdateExtraMealsPastEntriesListView(string rollNumber, string name, string mealType, double applicableRate, string date, string time)
        {
            var imageFilePath = imageFilePaths.FirstOrDefault(path => Path.GetFileNameWithoutExtension(path) == rollNumber);
            var imageSource = imageFilePath != null ? new BitmapImage(new Uri(imageFilePath)) : null;

            ExtraMealsPastEntriesListView.Items.Insert(0, new ExtraMealPastEntry
            {
                SerialNumber = ExtraMealsPastEntriesListView.Items.Count + 1,
                RollNumber = rollNumber,
                Name = name,
                Meal = mealType + " (Extra)",
                Price = applicableRate.ToString("C"),
                Date = date,
                Time = time,
                ImageSource = imageSource
            });

            UpdateSerialNumbers(ExtraMealsPastEntriesListView);
        }

        private void ProcessRatesDataForSheet(ExcelWorksheet worksheet, Dictionary<string, Dictionary<string, double>> data)
        {
            for (int col = 2; col <= worksheet.Dimension.Columns; col++)
            {
                var day = worksheet.Cells[1, col].Text;

                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    var mealType = worksheet.Cells[row, 1].Text;
                    var rateString = worksheet.Cells[row, col].Text;

                    if (double.TryParse(rateString, out double rate))
                    {
                        if (!data.ContainsKey(mealType))
                        {
                            data[mealType] = new Dictionary<string, double>();
                        }

                        data[mealType][day] = rate;
                    }
                    else
                    {
                        // Handle the case where parsing fails
                        Console.WriteLine($"Failed to parse rate for {mealType} on {day}: {rateString}");
                    }
                }
            }
        }


        private void SaveExtraMealsStudentDetails(string rollNumber, string name, string amount, double applicableRate, string mealType, string date, bool isManualSelection)
        {
            var logEntry = $"Roll Number: {rollNumber} Name: {name} Amount: {amount} Rate: {applicableRate:C} Meal: {mealType} Date: {date} Time: {DateTime.Now} Manual: {isManualSelection}\n";
            File.AppendAllText("extra_meals_log.txt", logEntry);

            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    if (worksheet.Cells[row, 3].Text == rollNumber)
                    {
                        int col = 18;
                        while (!string.IsNullOrEmpty(worksheet.Cells[row, col].Text))
                        {
                            col++;
                        }
                        worksheet.Cells[row, col].Value = $"{applicableRate:C} - {mealType} - (Extra) {date}";
                        break;
                    }
                }

                package.Save();

            }
        }


        private void PopulateExtraMealOptions()
        {
            if (extraMealOptionsData != null && extraMealOptionsData.Count > 0)
            {
                ExtraMealsListBox.Items.Clear();

                foreach (var mealOption in extraMealOptionsData)
                {
                    ExtraMealsListBox.Items.Add(new ListBoxItem { Content = mealOption.Key });
                }
            }
        }


        private void ExtraMealsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ExtraMealsListBox.SelectedItem is ListBoxItem selectedItem)
            {
                lastSelectedOption = selectedItem.Content.ToString();

                // Set focus back to the RollNumberTextBoxExtraMeals
                SetFocusToRollNumberTextBox();
            }
        }

        private void SaveExtraMealButton_Click(object sender, RoutedEventArgs e)
        {
            var rollNumber = RollNumberTextBoxExtraMeals.Text;

            if (string.IsNullOrEmpty(rollNumber))
            {
                MessageBox.Show("Please enter a roll number.");
                return;
            }

            if (string.IsNullOrEmpty(lastSelectedOption))
            {
                MessageBox.Show("Please select an extra meal option.");
                return;
            }

            if (extraMealOptionsData.TryGetValue(lastSelectedOption, out var mealData))
            {
                var mealName = lastSelectedOption;
                var mealPrice = mealData.First().Value;

                var student = studentDataList.FirstOrDefault(s => s.RollNumber == rollNumber);
                if (student != null)
                {
                    var date = DateTime.Now.ToString("dd-MM-yyyy");
                    SaveExtraMealDetails(rollNumber, student.Name, student.Amount, mealPrice, mealName, date);
                    MessageBox.Show($"Extra meal saved for Roll Number: {rollNumber}, Meal: {mealName}, Price: {mealPrice:C}");
                }
                else
                {
                    MessageBox.Show("Student not found.");
                }
            }
        }


        private void SaveExtraMealDetails(string rollNumber, string name, string amount, double applicableRate, string mealType, string date)
        {
            var time = DateTime.Now.ToString("HH:mm:ss");
            var logEntry = $"Roll Number: {rollNumber} Name: {name} Amount: {amount} Rate: {applicableRate:C} Meal: {mealType} (Extra) Date: {date} Time: {time}\n";
            File.AppendAllText("extra_meal_log.txt", logEntry);

            // Update the Excel file
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    if (worksheet.Cells[row, 3].Text == rollNumber)
                    {
                        int col = 18; // Start from column R (18th column)
                        while (!string.IsNullOrEmpty(worksheet.Cells[row, col].Text))
                        {
                            col++;
                        }
                        worksheet.Cells[row, col].Value = $"{applicableRate:C} - (Extra) {mealType} - {date}";

                        // Apply formulas

                        break;
                    }
                }

                package.Save();
            }





            // Update the ListView
            Dispatcher.Invoke(() =>
            {
                var imageFilePath = imageFilePaths.FirstOrDefault(path => Path.GetFileNameWithoutExtension(path) == rollNumber);
                var imageSource = imageFilePath != null ? new BitmapImage(new Uri(imageFilePath)) : null;
                PastEntriesListView.Items.Insert(0, new ExtraMealPastEntry
                {
                    SerialNumber = PastEntriesListView.Items.Count + 1,
                    RollNumber = rollNumber,
                    Name = name,
                    Meal = mealType + " (Extra)",
                    Price = applicableRate.ToString("C"),
                    Date = date,
                    Time = DateTime.Now.ToString("HH:mm:ss"),
                    ImageSource = imageSource
                });

                UpdateSerialNumbers(PastEntriesListView);
            });
        }


        private void DisplayRatesData()
        {
            RatesListBox.Items.Clear();
            DisplayRatesDataForListBox(ratesData, RatesListBox);

        }
        private async Task AutoUploadFiles()
        {
            try
            {
                await UploadFilesAsync();
                // Optionally, you can add a status update here
                // UploadStatusLabel.Text = "Files automatically uploaded successfully.";
            }
            catch (Exception ex)
            {
                // Handle any errors that occur during upload
                MessageBox.Show($"Error during automatic upload: {ex.Message}");
            }
        }   

        private void DisplayRatesDataForListBox(Dictionary<string, Dictionary<string, double>> data, ListBox listBox)
        {
            int frameSize = 10;
            foreach (var mealType in data.Keys)
            {
                var framePanel = new StackPanel { Orientation = System.Windows.Controls.Orientation.Vertical, Margin = new Thickness(5) };
                framePanel.Children.Add(new TextBlock { Text = mealType, FontWeight = FontWeights.Bold });

                var dayRates = data[mealType].ToList();
                for (int i = 0; i < dayRates.Count; i += frameSize)
                {
                    for (int j = i; j < i + frameSize && j < dayRates.Count; j++)
                    {
                        var day = dayRates[j].Key;
                        var rate = dayRates[j].Value;
                        framePanel.Children.Add(new TextBlock { Text = $"{day}: {rate:C}" });
                    }
                    listBox.Items.Add(framePanel);
                }
            }
        }


        private async void RollNumberTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (RollNumberTextBox.Text.Length == 8)
            {
                string rollNumber = RollNumberTextBox.Text;
                if (HasPreviousEntry(rollNumber))
                {
                    ShowMealSelectionPopup(rollNumber);
                }
                else
                {
                    DisplayStudentDetails(rollNumber);
                }
                RollNumberTextBox.Clear();

                await AutoUploadFiles();
                RefocusInputBox();
            }
        }

        private bool HasPreviousEntry(string rollNumber)
        {
            return DashboardPastEntriesListView.Items.Cast<DashboardPastEntry>()
                .Any(entry => entry.RollNumber == rollNumber && entry.Date == DateTime.Now.ToString("dd-MM-yyyy"));
        }


        private void ShowMealSelectionPopup(string rollNumber)
        {
            var popup = new Popup
            {
                Width = 200,
                Height = 150,
                IsOpen = true,
                PlacementTarget = RollNumberTextBox,
                Placement = PlacementMode.Bottom,
                StaysOpen = false
            };

            var listBox = new ListBox();
            listBox.Items.Add(new ListBoxItem { Content = "Extra Meal" });
            listBox.Items.Add(new ListBoxItem { Content = "Another Meal" });
            listBox.Items.Add(new ListBoxItem { Content = "Cancel" });

            listBox.SelectionChanged += async (s, e) =>
            {
                if (listBox.SelectedItem is ListBoxItem selectedItem)
                {
                    popup.IsOpen = false;
                    ForceFocusOnRollNumberTextBox(); // Force focus immediately after selection

                    switch (selectedItem.Content.ToString())
                    {
                        case "Extra Meal":
                            await HandleExtraMeal(rollNumber);
                            break;
                        case "Another Meal":
                            DisplayStudentDetails(rollNumber);
                            break;
                        case "Cancel":
                            break;
                    }
                }
            };

            popup.Child = listBox;

            // Force focus on RollNumberTextBox after showing the popup
            popup.Opened += (s, e) => ForceFocusOnRollNumberTextBox();
        }

        private void ForceFocusOnRollNumberTextBox()
        {
            Dispatcher.InvokeAsync(() =>
            {
                RollNumberTextBox.Focus();
                RollNumberTextBox.SelectAll();

                // Simulate a click on the TextBox to ensure it's active
                var mouseDownEvent = new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left)
                {
                    RoutedEvent = Mouse.MouseDownEvent,
                    Source = RollNumberTextBox
                };
                RollNumberTextBox.RaiseEvent(mouseDownEvent);

                var mouseUpEvent = new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left)
                {
                    RoutedEvent = Mouse.MouseUpEvent,
                    Source = RollNumberTextBox
                };
                RollNumberTextBox.RaiseEvent(mouseUpEvent);
            }, System.Windows.Threading.DispatcherPriority.Input);
        }

        private async Task HandleExtraMeal(string rollNumber)
        {
            var student = studentDataList.FirstOrDefault(s => s.RollNumber == rollNumber);
            if (student != null)
            {
                double applicableRate = GetExtraMealApplicableRate();
                string mealType = GetMealType(DateTime.Now.TimeOfDay) + " (Extra)";
                var date = DateTime.Now.ToString("dd-MM-yyyy");
                var time = DateTime.Now.ToString("HH:mm:ss");

                SaveExtraMealDetails(rollNumber, student.Name, student.Amount, applicableRate, mealType, date, time);
                UpdateExtraMealsPastEntriesListView(rollNumber, student.Name, mealType, applicableRate, date, time);

                await Dispatcher.InvokeAsync(() =>
                {
                });
            }
            else
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    MessageBox.Show("Student not found.");
                });
            }
        }

        private async void RollNumberTextBoxExtraMeals_TextChanged(object sender, TextChangedEventArgs e)
        {
            var textBox = sender as TextBox;
            if (textBox != null)
            {
                string text = new string(textBox.Text.Where(char.IsDigit).ToArray());

                if (text.Length > 8)
                {
                    text = text.Substring(0, 8);
                }

                int cursorPosition = textBox.SelectionStart;
                textBox.Text = text;
                textBox.SelectionStart = Math.Min(cursorPosition, text.Length);

                if (text.Length == 8)
                {
                    ShowUpdatedSnackSelectionPopup(text);
                    await AutoUploadFiles();
                    textBox.Clear();
                    ForceFocusOnExtraMealsRollNumberTextBox();
                }
            }
        }

        private void RefocusInputBox()
        {
            Dispatcher.InvokeAsync(() =>
            {
                if (MainTabControl.SelectedItem is System.Windows.Controls.TabItem selectedTab)
                {
                    TextBox textBoxToFocus = null;

                    if (selectedTab.Name == "Dashboard")
                    {
                        textBoxToFocus = RollNumberTextBox;
                    }
                    else if (selectedTab.Name == "ExtraMeals")
                    {
                        textBoxToFocus = RollNumberTextBoxExtraMeals;
                    }

                    if (textBoxToFocus != null)
                    {
                        textBoxToFocus.Focus();
                        textBoxToFocus.SelectAll();

                        // Simulate a click on the TextBox
                        var mouseDownEvent = new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left)
                        {
                            RoutedEvent = Mouse.MouseDownEvent,
                            Source = textBoxToFocus
                        };
                        textBoxToFocus.RaiseEvent(mouseDownEvent);

                        var mouseUpEvent = new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left)
                        {
                            RoutedEvent = Mouse.MouseUpEvent,
                            Source = textBoxToFocus
                        };
                        textBoxToFocus.RaiseEvent(mouseUpEvent);
                    }
                }
            }, System.Windows.Threading.DispatcherPriority.Input);
        }


        private Popup currentPopup; // Add this as a class-level field

        private void ShowUpdatedSnackSelectionPopup(string rollNumber)
        {
            if (currentPopup != null)
            {
                currentPopup.IsOpen = false;
            }

            currentPopup = new Popup
            {
                Width = 200,
                Height = 200,
                IsOpen = true,
                StaysOpen = false,
                PlacementTarget = RollNumberTextBoxExtraMeals,
                Placement = PlacementMode.Bottom
            };

            var listBox = new ListBox
            {
                SelectionMode = SelectionMode.Single
            };

            foreach (var item in ExtraMealsListBox.Items)
            {
                if (item is ListBoxItem listBoxItem)
                {
                    listBox.Items.Add(new ListBoxItem { Content = listBoxItem.Content });
                }
            }

            listBox.SelectionChanged += async (s, e) =>
            {
                if (listBox.SelectedItem is ListBoxItem selectedItem)
                {
                    currentPopup.IsOpen = false;
                    ForceFocusOnExtraMealsRollNumberTextBox(); // Force focus immediately after selection

                    await SaveUpdatedExtraMealAutomaticallyAsync(rollNumber, selectedItem.Content.ToString());
                    RollNumberTextBoxExtraMeals.Clear();
                    await AutoUploadFiles();

                    if (currentPopup.Parent is Panel panel)
                    {
                        panel.Children.Remove(currentPopup);
                    }
                    currentPopup = null;
                }
            };

            currentPopup.Child = listBox;
            currentPopup.IsOpen = true;

            // Force focus on ExtraMealsRollNumberTextBox after showing the popup
            currentPopup.Opened += (s, e) => ForceFocusOnExtraMealsRollNumberTextBox();
        }

        private void ForceFocusOnExtraMealsRollNumberTextBox()
        {
            Dispatcher.InvokeAsync(() =>
            {
                RollNumberTextBoxExtraMeals.Focus();
                RollNumberTextBoxExtraMeals.SelectAll();

                // Simulate a click on the TextBox to ensure it's active
                var mouseDownEvent = new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left)
                {
                    RoutedEvent = Mouse.MouseDownEvent,
                    Source = RollNumberTextBoxExtraMeals
                };
                RollNumberTextBoxExtraMeals.RaiseEvent(mouseDownEvent);

                var mouseUpEvent = new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left)
                {
                    RoutedEvent = Mouse.MouseUpEvent,
                    Source = RollNumberTextBoxExtraMeals
                };
                RollNumberTextBoxExtraMeals.RaiseEvent(mouseUpEvent);
            }, System.Windows.Threading.DispatcherPriority.Input);
        }

        private void FocusInputAfterPopup()
        {
            Dispatcher.InvokeAsync(() =>
            {
                if (MainTabControl.SelectedItem is System.Windows.Controls.TabItem selectedTab)
                {
                    TextBox textBoxToFocus = null;

                    if (selectedTab.Name == "Dashboard")
                    {
                        textBoxToFocus = RollNumberTextBox;
                    }
                    else if (selectedTab.Name == "ExtraMeals")
                    {
                        textBoxToFocus = RollNumberTextBoxExtraMeals;
                    }

                    if (textBoxToFocus != null)
                    {
                        textBoxToFocus.Focus();
                        textBoxToFocus.CaretIndex = textBoxToFocus.Text.Length;
                    }
                }
            }, System.Windows.Threading.DispatcherPriority.Input);
        }
        private async Task SaveUpdatedExtraMealAutomaticallyAsync(string rollNumber, string selectedMeal)
        {
            if (extraMealOptionsData.TryGetValue(selectedMeal, out var mealData))
            {
                var mealName = selectedMeal;
                var mealPrice = mealData.First().Value;

                var student = studentDataList.FirstOrDefault(s => s.RollNumber == rollNumber);
                if (student != null)
                {
                    var date = DateTime.Now.ToString("dd-MM-yyyy");
                    var time = DateTime.Now.ToString("HH:mm:ss");
                    SaveExtraMealDetails(rollNumber, student.Name, student.Amount, mealPrice, mealName, date, time);
                    UpdatePastEntriesListView(rollNumber, student.Name, mealName, mealPrice, date, time);
                    await Dispatcher.InvokeAsync(() =>
                    {
                    });
                }
                else
                {
                    await Dispatcher.InvokeAsync(() =>
                    {
                        // Optional: Update a status label instead of showing a pop-up
                        StatusLabel.Text = $"Extra meal saved for Roll Number: {rollNumber}, Meal: {mealName}, Price: {mealPrice:C}";
                        RollNumberTextBoxExtraMeals.Focus();
                        SetFocusToActiveTextBox();
                    });
                }
            }
        }


        private void ShowSnackSelectionPopup(string rollNumber)
        {
            var popup = new Popup
            {
                Width = 200,
                Height = 200,
                IsOpen = true,
                PlacementTarget = RollNumberTextBoxExtraMeals,
                Placement = PlacementMode.Bottom
            };

            var listBox = new ListBox();
            foreach (var item in ExtraMealsListBox.Items)
            {
                listBox.Items.Add(item);
            }

            listBox.SelectionChanged += (s, e) =>
            {
                if (listBox.SelectedItem is ListBoxItem selectedItem)
                {
                    popup.IsOpen = false;
                    SaveExtraMealAutomatically(rollNumber, selectedItem.Content.ToString());
                    RollNumberTextBoxExtraMeals.Clear(); // Clear the textbox after selection
                }
            };

            popup.Child = listBox;
        }

        private void SaveExtraMealAutomatically(string rollNumber, string selectedMeal)
        {
            if (extraMealOptionsData.TryGetValue(selectedMeal, out var mealData))
            {
                var mealName = selectedMeal;
                var mealPrice = mealData.First().Value;

                var student = studentDataList.FirstOrDefault(s => s.RollNumber == rollNumber);
                if (student != null)
                {
                    var date = DateTime.Now.ToString("dd-MM-yyyy");
                    var time = DateTime.Now.ToString("HH:mm:ss");
                    SaveExtraMealDetails(rollNumber, student.Name, student.Amount, mealPrice, mealName, date, time);
                    UpdatePastEntriesListView(rollNumber, student.Name, mealName, mealPrice, date, time);
                    // Optional: Update a status label instead of showing a pop-up
                    StatusLabel.Text = $"Extra meal saved for Roll Number: {rollNumber}, Meal: {mealName}, Price: {mealPrice:C}";
                    RollNumberTextBoxExtraMeals.Focus();
                }
                else
                {
                    MessageBox.Show("Student not found.");
                }
            }
        }

        
        private void UpdatePastEntriesListView(string rollNumber, string name, string mealType, double price, string date, string time)
        {
            var imageFilePath = imageFilePaths.FirstOrDefault(path => Path.GetFileNameWithoutExtension(path) == rollNumber);
            var imageSource = imageFilePath != null ? new BitmapImage(new Uri(imageFilePath)) : null;

            PastEntriesListView.Items.Insert(0, new ExtraMealPastEntry
            {
                SerialNumber = PastEntriesListView.Items.Count + 1,
                RollNumber = rollNumber,
                Name = name,
                Meal = mealType,
                Price = price.ToString("C"),
                Date = date,
                Time = time,
                ImageSource = imageSource
            });

            UpdateSerialNumbers(PastEntriesListView);
        }

        



       


        private void DisplayStudentDetails(string rollNumber)
        {
            var student = studentDataList.FirstOrDefault(s => s.RollNumber == rollNumber);
            if (student != null)
            {
                double applicableRate;
                string mealType;

                applicableRate = GetApplicableRate();
                mealType = GetMealType(DateTime.Now.TimeOfDay);

                var date = DateTime.Now.ToString("dd-MM-yyyy");
                var entryDetails = $"{applicableRate:C} - {mealType} - {date}";

                StudentDetailsTextBlock.Text = $"Roll Number: {student.RollNumber}\nName: {student.Name}\nFathers Name: {student.Amount}\nRate: {applicableRate:C}";

                var imageFilePath = imageFilePaths.FirstOrDefault(path => Path.GetFileNameWithoutExtension(path) == rollNumber);
                if (imageFilePath != null)
                {
                    var bitmap = new BitmapImage(new Uri(imageFilePath));
                    StudentImage.Source = bitmap;
                }
                else
                {
                    StudentImage.Source = null;
                }

                RateDetailsTextBlock.Text = $"Applicable Rate: {applicableRate:C}";

                SaveStudentDetails(rollNumber, student.Name, student.Amount, applicableRate, mealType, date);
                UpdatePastEntriesListBox(rollNumber, student.Name, student.Amount, applicableRate, mealType, date);
            }
            else
            {
                StudentDetailsTextBlock.Text = "Student not found.";
                StudentImage.Source = null;
                RateDetailsTextBlock.Text = string.Empty;
            }
        }


        private void DisplayExtraMealStudentDetails(string rollNumber)
        {
            var student = studentDataList.FirstOrDefault(s => s.RollNumber == rollNumber);
            if (student != null)
            {
                double applicableRate = GetExtraMealApplicableRate();
                string mealType = GetMealType(DateTime.Now.TimeOfDay);
                var date = DateTime.Now.ToString("dd-MM-yyyy");


                var imageFilePath = imageFilePaths.FirstOrDefault(path => Path.GetFileNameWithoutExtension(path) == rollNumber);
                if (imageFilePath != null)
                {
                    var bitmap = new BitmapImage(new Uri(imageFilePath));
                }
                else
                {
                }


                SaveExtraMealDetails(rollNumber, student.Name, student.Amount, applicableRate, mealType, date, DateTime.Now.ToString("HH:mm:ss"));
                UpdateExtraMealsPastEntriesListView(rollNumber, student.Name, mealType, applicableRate, date, DateTime.Now.ToString("HH:mm:ss"));
            }
            else
            {
            }
        }

        private double GetApplicableRate()
        {
            var dayOfWeek = DateTime.Now.DayOfWeek.ToString();
            var currentTime = DateTime.Now.TimeOfDay;
            var mealType = GetMealType(currentTime);

            if (ratesData.ContainsKey(mealType) && ratesData[mealType].ContainsKey(dayOfWeek))
            {
                return ratesData[mealType][dayOfWeek];
            }
            else
            {
                // If the specific meal type is not found, try to find a default rate
                foreach (var meal in ratesData.Keys)
                {
                    if (ratesData[meal].ContainsKey(dayOfWeek))
                    {
                        return ratesData[meal][dayOfWeek];
                    }
                }
            }

            // If no rate is found, return a default value or throw an exception
            MessageBox.Show($"No rate found for {mealType} on {dayOfWeek}. Using default rate of 0.");
            return 0.0;
        }


        private string GetMealType(TimeSpan currentTime)
        {
            if (currentTime >= new TimeSpan(6, 0, 0) && currentTime <= new TimeSpan(12, 0, 0))
            {
                return "Breakfast";
            }
            if (currentTime >= new TimeSpan(12, 0, 0) && currentTime <= new TimeSpan(15, 0, 0))
            {
                return "Lunch";
            }
            if (currentTime >= new TimeSpan(15, 1, 0) && currentTime <= new TimeSpan(18, 0, 0))
            {
                return "Snacks";
            }
            if (currentTime >= new TimeSpan(18, 0, 0) && currentTime <= new TimeSpan(22, 0, 0))
            {
                return "Dinner";
            }
            return "Other";
        }

        private string GetExcelColumnName(int columnNumber)
        {
            string columnName = "";
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return columnName;
        }

        private void SaveStudentDetails(string rollNumber, string name, string amount, double applicableRate, string mealType, string date)
        {
            var logEntry = $"Roll Number: {rollNumber} Name: {name} Amount: {amount} Rate: {applicableRate:C} Meal: {mealType} Date: {date} Time: {DateTime.Now}\n";
            File.AppendAllText("student_log.txt", logEntry);

            // Update the Excel file
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    if (worksheet.Cells[row, 3].Text == rollNumber)
                    {
                        int col = 18; // Start from column R (18th column)
                        while (!string.IsNullOrEmpty(worksheet.Cells[row, col].Text))
                        {
                            col++;
                        }
                        worksheet.Cells[row, col].Value = $"{applicableRate:C} - {mealType} - {date}";



                    }
                }
                package.Save();
            }
        }

        private void ApplyFormulas_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    using (var package = new ExcelPackage(new FileInfo(openFileDialog.FileName)))
                    {
                        var worksheet = package.Workbook.Worksheets[0];

                        // Apply formulas
                        int lastRow = worksheet.Dimension.End.Row;

                        for (int row = 2; row <= lastRow; row++)
                        {
                            // Column I
                            worksheet.Cells[row, 9].Formula = "SUMPRODUCT(--(ISNUMBER(SEARCH(\"Breakfast\",R" + row + ":ZZ" + row + "))+ISNUMBER(SEARCH(\"Lunch\",R" + row + ":ZZ" + row + "))+ISNUMBER(SEARCH(\"Dinner\",R" + row + ":ZZ" + row + "))),--(NOT(ISNUMBER(SEARCH(\"Extra\",R" + row + ":ZZ" + row + ")))))";

                            // Column J
                            worksheet.Cells[row, 10].Value = double.Parse(ColumnJValue.Text);

                            // Column K
                            worksheet.Cells[row, 11].Formula = $"I{row}*J{row}";
                            // For column L
                            worksheet.Cells[row, 12].CreateArrayFormula("SUMPRODUCT(IF(LEN(R" + row + ":ZZ" + row + ") > 0, IF(ISNUMBER(SEARCH(\"Extra\", R" + row + ":ZZ" + row + ")), VALUE(MID(R" + row + ":ZZ" + row + ", FIND(\"₹\", R" + row + ":ZZ" + row + ") + 2, FIND(\" \", R" + row + ":ZZ" + row + " & \" \", FIND(\"₹\", R" + row + ":ZZ" + row + ") + 2) - FIND(\"₹\", R" + row + ":ZZ" + row + ") - 2)), 0), 0))");

                            // Column N
                            worksheet.Cells[row, 14].Value = double.Parse(ColumnNValue.Text);

                            // Column O
                            worksheet.Cells[row, 15].Formula = $"K{row}+L{row}+M{row}+N{row}";

                            // Column P
                            worksheet.Cells[row, 16].Value = double.Parse(ColumnPValue.Text);

                            // Column Q
                            worksheet.Cells[row, 17].Formula = $"P{row}-O{row}";
                        }

                        package.Save();
                    }

                    LogTextBox.Text += "Formulas applied successfully.\n";
                }
                catch (Exception ex)
                {
                    LogTextBox.Text += $"Error: {ex.Message}\n";
                }
            }
        }




        private void UpdatePastEntriesListBox(string rollNumber, string name, string amount, double applicableRate, string mealType, string date)
        {
            var imageFilePath = imageFilePaths.FirstOrDefault(path => Path.GetFileNameWithoutExtension(path) == rollNumber);
            var imageSource = imageFilePath != null ? new BitmapImage(new Uri(imageFilePath)) : null;

            var pastEntry = new DashboardPastEntry
            {
                SerialNumber = DashboardPastEntriesListView.Items.Count + 1,
                RollNumber = rollNumber,
                Name = name,
                Amount = amount,
                Rate = $"{applicableRate:C}",
                Meal = mealType,
                Date = date,
                Time = DateTime.Now.ToString("HH:mm:ss"),
                ImageSource = imageSource
            };

            DashboardPastEntriesListView.Items.Insert(0, pastEntry);
            UpdateSerialNumbers(DashboardPastEntriesListView);
        }

        private void UpdateSerialNumbers(ListView listView)
        {
            int count = listView.Items.Count;
            foreach (var item in listView.Items)
            {
                if (item is DashboardPastEntry dashboardEntry)
                {
                    dashboardEntry.SerialNumber = count--;
                }
                else if (item is ExtraMealPastEntry extraMealEntry)
                {
                    extraMealEntry.SerialNumber = count--;
                }
            }
            listView.Items.Refresh();
        }

        private void SetFocusToRollNumberTextBox()
        {
            RollNumberTextBoxExtraMeals.Focus();
            RollNumberTextBoxExtraMeals.CaretIndex = RollNumberTextBoxExtraMeals.Text.Length;
        }
        private void RollNumberTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && RollNumberTextBox.Text.Length == 8)
            {
                string rollNumber = RollNumberTextBox.Text;
                DisplayStudentDetails(rollNumber);
                RollNumberTextBox.Clear();

                // Set focus back to the TextBox
                RollNumberTextBox.Focus();
            }
        }

        private void DashboardPastEntriesListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }

 
    public class StudentData
    {
        public string RollNumber { get; set; }
        public string Name { get; set; }
        public string Amount { get; set; }
    }
}
public class DashboardPastEntry
{
    public int SerialNumber { get; set; }
    public string RollNumber { get; set; }
    public string Name { get; set; }
    public string Amount { get; set; }
    public string Rate { get; set; }
    public string Meal { get; set; }
    public string Date { get; set; }
    public string Time { get; set; }
    public BitmapImage ImageSource { get; set; }
}

public class ExtraMealPastEntry
{
    public int SerialNumber { get; set; }
    public string RollNumber { get; set; }
    public string Name { get; set; }
    public string Meal { get; set; }
    public string Price { get; set; }
    public string Date { get; set; }
    public string Time { get; set; }
    public BitmapImage ImageSource { get; set; }
}