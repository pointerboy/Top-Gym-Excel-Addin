using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Top_Gym_Excel
{
    public partial class ThisAddIn
    {
        // ConcurrentDictionary for thread-safe caching of users
        public ConcurrentDictionary<int, User> UserCache = new ConcurrentDictionary<int, User>();

        // List to store the marked users' information
        public List<string> MarkedUsersList = new List<string>();

        public ListBox markedUsersListBox; // This is a reference to the ListBox in your form

        // User control
        private UserControl _usr;
        // Custom task pane
        private Microsoft.Office.Tools.CustomTaskPane _myCustomTaskPane;

        private SerialPort _serialPort;
        private StringBuilder _dataBuffer = new StringBuilder();
        private string CurrentBarCode = string.Empty;
        private Thread _barcodeThread;

        private ConcurrentQueue<string> barcodeQueue = new ConcurrentQueue<string>(); // Queue for barcodes
        private bool isProcessingBarcodes = false; // To track whether barcode processing is in progress

        private static Dictionary<string, int> MonthColumnCache = new Dictionary<string, int>();
        private static Dictionary<string, int> DateColumnCache = new Dictionary<string, int>();

        List<string> scannedBarcodes = new List<string>();


        // Method to update the ListBox with the current list of marked users
        public void UpdateMarkedUsersListBox()
        {
            // Ensure the ListBox is available and clear it first
            if (markedUsersListBox != null)
            {
                markedUsersListBox.Items.Clear();

                // Add each marked user to the ListBox
                foreach (var user in MarkedUsersList)
                {
                    markedUsersListBox.Items.Add(user);
                }
            }
        }

        // Method to show the users form with the marked users
        public void ShowUsersForm()
        {
            // Create a new form
            Form usersForm = new Form
            {
                Text = "Marked Users",
                Width = 400,
                Height = 600
            };

            // Create a ListBox to display marked users
            markedUsersListBox = new ListBox
            {
                Dock = DockStyle.Fill
            };

            // Add the marked users to the ListBox
            UpdateMarkedUsersListBox();

            // Add the ListBox to the form
            usersForm.Controls.Add(markedUsersListBox);

            // Show the form
            usersForm.ShowDialog();
        }

        private void BarcodeScanningThread()
        {
            try
            {
                // Open the serial port
                _serialPort.Open();
                while (_serialPort.IsOpen)
                {
                    // Wait for data to be received
                    Thread.Sleep(100); // Sleep to avoid busy waiting, adjust as needed
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error initializing barcode scanner: " + ex.Message);
            }
        }

        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            var serialPort = (SerialPort)sender;
            var data = serialPort.ReadExisting();
            _dataBuffer.Append(data);

            // If the data ends with a carriage return (\r), treat it as a full barcode
            if (_dataBuffer.ToString().Contains('\r'))
            {
                var fullBarcode = _dataBuffer.ToString().Trim();
                _dataBuffer.Clear();

                // check for new year

                // Enqueue the barcode for processing
                barcodeQueue.Enqueue(fullBarcode);

                // Start barcode processing if it's not already running
                if (!isProcessingBarcodes)
                {
                    StartBarcodeProcessing();
                }
            }
        }


        // Start processing barcodes from the queue
        private void StartBarcodeProcessing()
        {
            if (isProcessingBarcodes) return;  // Avoid starting multiple tasks

            isProcessingBarcodes = true;

            // Start a task to process the barcodes in the queue
            Task.Run(async () =>
            {
                // Keep processing as long as there are barcodes in the queue
                while (barcodeQueue.Count > 0)
                {
                    if (barcodeQueue.TryDequeue(out var barcode))
                    {
                        await ProcessBarcodeScan(barcode);  // Process each barcode asynchronously
                        ShowBarcodeDebugDialog(barcode);   // Display the barcode (or handle it in your way)
                    }
                }

                isProcessingBarcodes = false;  // Reset processing flag when done
            });
        }

        private void ShowBarcodeDebugDialog(string barcode)
        {
            // This will show the barcode as a debug message
            MessageBox.Show("Scanned Barcode: " + barcode, "Debug Barcode", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // Make sure to properly close the serial port when your application is done
        public void CloseSerialPort()
        {
            if (_serialPort.IsOpen)
            {
                _serialPort.Close();
            }
        }


        public UserControl1 UserControl { get; private set; }
        private async void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
               
                // Create an instance of the user control
                UserControl = new UserControl1(); // Store the instance for later access
                _myCustomTaskPane = CustomTaskPanes.Add(UserControl, "Prijavljeni korisnici danas");
                _myCustomTaskPane.Visible = true;

                await WaitForExcelToLoad();
                Excel.Application excelApp = Globals.ThisAddIn.Application;
                Excel.Worksheet worksheet = excelApp.ActiveSheet as Excel.Worksheet;
                if (!worksheet.Name.StartsWith("Top"))
                {
                    ThisAddIn_Shutdown(sender, e);
                }


                _serialPort = new SerialPort("COM3", 9600, Parity.None, 8, StopBits.One);
                _serialPort.DataReceived += SerialPort_DataReceived;

                // Start a thread to handle the barcode scanning process
                _barcodeThread = new Thread(BarcodeScanningThread);
                _barcodeThread.Start();

                MessageBox.Show("Program je spreman za korišćenje", "Top Gym", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Critical error during startup: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // Method to find the month column (cached version)
        private int FindMonthColumn(dynamic worksheet, string month)
        {
            // Use cached value if available
            if (MonthColumnCache.ContainsKey(month))
            {
                return MonthColumnCache[month];
            }

            for (int col = 14; col <= worksheet.UsedRange.Columns.Count; col++) // Assuming month columns start at column 14
            {
                string cellValue = worksheet.Cells[2, col]?.Text?.ToString();
                if (!string.IsNullOrEmpty(cellValue) && cellValue.Equals(month, StringComparison.OrdinalIgnoreCase))
                {
                    // Cache the result for future calls
                    MonthColumnCache[month] = col;
                    return col;
                }
            }
            return -1;
        }

        // Method to find the date column (cached version)
        private int FindDateColumn(dynamic worksheet, int day, int startColumn)
        {
            string dayString = day.ToString("D2");

            // Use cached value if available
            if (DateColumnCache.ContainsKey(dayString))
            {
                return DateColumnCache[dayString];
            }

            for (int col = startColumn; col <= worksheet.UsedRange.Columns.Count; col++)
            {
                string cellValue = worksheet.Cells[3, col]?.Text?.ToString();
                if (!string.IsNullOrEmpty(cellValue) && cellValue == dayString)
                {
                    // Cache the result for future calls
                    DateColumnCache[dayString] = col;
                    return col;
                }
            }
            return -1;
        }

        private int lastMonthColumn = -1; // Cache for last month column

        private int GetUserRowFromBarcode(string barcode)
        {
            // Convert the barcode to an integer (removes leading zeros)
            if (int.TryParse(barcode.TrimStart('0'), out int userId))
            {
                // Add 5 to the ID to get the correct row number
                return userId + 5;
            }
            else
            {
                return -1;  // Return -1 if barcode is not valid
            }
        }

        private async Task ProcessBarcodeScan(string barcode)
        {
            // Get the Excel application object
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Worksheet worksheet = excelApp.ActiveSheet as Excel.Worksheet;

            if (worksheet != null)
            {
                // Find the row for the user based on the barcode (ID)
                int userRow = GetUserRowFromBarcode(barcode);

                if (userRow != -1)
                {
                    // Extract user information from the corresponding row (+5 offset)
                    string name = worksheet.Cells[userRow, 5].Text;  // E = Name
                    string lastName = worksheet.Cells[userRow, 6].Text;  // F = Lastname
                    string membershipStatus = worksheet.Cells[userRow, 12].Text;  // L = Membership Status

                    if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(lastName) || string.IsNullOrWhiteSpace(membershipStatus))
                    {
                        MessageBox.Show("Skeniran je nepostojeći korisnik!");
                        return;
                    }

                    // Check if the user has a valid membership
                    if (membershipStatus == "Isteklo")
                    {
                        // Show the flashing warning for expired membership
                        if (Globals.ThisAddIn.UserControl != null)
                        {
                            Globals.ThisAddIn.UserControl.FlashWarning($"Članarina istekla za korisnika {barcode} {name} {lastName}");
                            Globals.ThisAddIn.UserControl.Focus();
                        }
                    }

                    if (scannedBarcodes.Contains(barcode))
                    {
                        Globals.ThisAddIn.UserControl.FlashWarning("Korisnik " + barcode + " je skeniran 2 puta!");
                        return; // Exit early to prevent marking the arrival again
                    }
                    else
                    {
                        // Get the current date components (month and day)
                        DateTime currentDate = DateTime.Now;
                        string currentMonth = currentDate.ToString("MMMM");  // Full month name
                        int currentDay = currentDate.Day;

                        // Use cached month column if valid, otherwise find it
                        if (lastMonthColumn == -1 || worksheet.Cells[2, lastMonthColumn]?.Text?.ToString() != currentMonth)
                        {
                            lastMonthColumn = FindMonthColumn(worksheet, currentMonth);

                            if (lastMonthColumn == -1)
                            {
                                MessageBox.Show($"Month {currentMonth} not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }

                        // Find the date column
                        int dateColumn = FindDateColumn(worksheet, currentDay, lastMonthColumn);
                        if (dateColumn == -1)
                        {
                            MessageBox.Show($"Date {currentDay} not found for {currentMonth}.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        // Mark the arrival with "X"
                        Excel.Range arrivalCell = worksheet.Cells[userRow, dateColumn];
                        arrivalCell.Value = "X";

                        // Show success message
                       // MessageBox.Show($"Ručno upisan korisnik {name} {lastName} (ID: {barcode}) - {currentDate.ToShortDateString()}", "Arrival Marked", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // Add user to MarkedUsersList
                        if (!Globals.ThisAddIn.MarkedUsersList.Contains(barcode))
                        {
                            Globals.ThisAddIn.MarkedUsersList.Add(name + " " + lastName + " kod:" + barcode + " " + currentDate.Hour + ":" + currentDate.Minute
                               + ":" + currentDate.Millisecond);
                            scannedBarcodes.Add(barcode);
                            Globals.ThisAddIn.UserControl.UpdateArrivalList(Globals.ThisAddIn.MarkedUsersList);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("User not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Unable to access the Excel worksheet.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


         public void ShowMarkedUsersWindow()
        {
            Form markedUsersForm = new Form
            {
                Text = "Marked Users",
                Width = 400,
                Height = 600
            };

            ListBox markedUsersListBox = new ListBox
            {
                Dock = DockStyle.Fill
            };

            // Add the marked users to the list box
            foreach (var user in MarkedUsersList)
            {
                markedUsersListBox.Items.Add(user);
            }

            markedUsersForm.Controls.Add(markedUsersListBox);
            markedUsersForm.ShowDialog();
        }

        private async Task WaitForExcelToLoad()
        {
            while (Application.ActiveSheet == null)
            {
                await Task.Delay(100); // Wait for 100ms before checking again
            }
        }

        private void ShowExceptions(ConcurrentBag<Exception> exceptions)
        {
            StringBuilder errorMessages = new StringBuilder("The following errors occurred:\n\n");

            foreach (var exception in exceptions)
            {
                errorMessages.AppendLine(exception.Message);
            }

            MessageBox.Show(errorMessages.ToString(), "Processing Errors", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

      
        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        #endregion

        // User model
        public class User
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public string LastName { get; set; }
            public string Barcode { get; set; }
            public string MembershipStatus { get; set; }
            public int RowIndex { get; set; } // New property to store the row index
        }
    }
}

