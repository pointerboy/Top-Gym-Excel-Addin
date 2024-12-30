using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Top_Gym_Excel
{
    public partial class Ribbon1
    {
        private int lastMonthColumn = -1; // Cache for last month column
        List<string> scannedBarcodes = new List<string>();

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

        private async void button1_Click(object sender, RibbonControlEventArgs e)
        {
            // Open a dialog for barcode input
            string barcode = Microsoft.VisualBasic.Interaction.InputBox("Ručno unesite barkod:", "Barcode Input", "");

            // Validate input
            if (string.IsNullOrEmpty(barcode))
            {
                MessageBox.Show("Ovo nije validan barkod", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

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
                        MessageBox.Show($"Ručno upisan korisnik {name} {lastName} (ID: {barcode}) - {currentDate.ToShortDateString()}", "Arrival Marked", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        scannedBarcodes.Add(barcode);
                        // Add user to MarkedUsersList
                        if (!Globals.ThisAddIn.MarkedUsersList.Contains(barcode))
                        {
                            Globals.ThisAddIn.MarkedUsersList.Add(name + " " + lastName + " kod:" + barcode + " " + currentDate.Hour + ":" + currentDate.Minute
                                + ":" + currentDate.Millisecond);
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

        private int FindMonthColumn(Excel.Worksheet worksheet, string month)
        {
            for (int col = 14; col <= worksheet.UsedRange.Columns.Count; col++) // Assuming month columns start at column 14
            {
                string cellValue = worksheet.Cells[2, col]?.Text?.ToString();
                if (!string.IsNullOrEmpty(cellValue) && cellValue.Equals(month, StringComparison.OrdinalIgnoreCase))
                {
                    return col;
                }
            }
            return -1;
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        { }

        private int FindDateColumn(Excel.Worksheet worksheet, int day, int startColumn)
        {
            string dayString = day.ToString("D2");

            for (int col = startColumn; col <= worksheet.UsedRange.Columns.Count; col++)
            {
                string cellValue = worksheet.Cells[3, col]?.Text?.ToString();
                if (!string.IsNullOrEmpty(cellValue) && cellValue == dayString)
                {
                    return col;
                }
            }
            return -1;
        }
    }
}
