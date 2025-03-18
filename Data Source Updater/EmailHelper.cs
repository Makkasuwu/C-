using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using ClosedXML.Excel;
using MimeKit;
using MailKit.Net.Smtp;

class Program
{
    static void Main()
    {
        string filePath = ""; // Change this to your actual file path
        var groupedData = ReadAndGroupExcelData(filePath);
        
        foreach (var entry in groupedData)
        {
            string recipientEmail = entry.Key;
            string name = entry.Value.Item1;
            List<string> dataSources = entry.Value.Item2;
            
            string emailBody = $"Hello {name},\n\nHere are your assigned data sources:\n- {string.Join("\n- ", dataSources)}\n\nBest,\nYour Team";

            // Send the email
            SendEmail(recipientEmail, "Your Data Sources", emailBody);
            Console.WriteLine($"Email sent to {recipientEmail}");

            // Update Excel with timestamp
            UpdateExcelWithTimestamp(filePath, recipientEmail);
        }
    }

    static Dictionary<string, (string, List<string>)> ReadAndGroupExcelData(string filePath)
    {
        var groupedData = new Dictionary<string, (string, List<string>)>();

        using (var workbook = new XLWorkbook(filePath))
        {
            var worksheet = workbook.Worksheet(1);
            var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip header row

            foreach (var row in rows)
            {
                string name = row.Cell(1).GetString();
                string email = row.Cell(2).GetString();
                string dataSource = row.Cell(3).GetString();

                if (!groupedData.ContainsKey(email))
                {
                    groupedData[email] = (name, new List<string>());
                }

                groupedData[email].Item2.Add(dataSource);
            }
        }

        return groupedData;
    }

    static void SendEmail(string recipient, string subject, string body)
    {
        try
        {
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress("Marcus Bullock", "")); // your email
            message.To.Add(new MailboxAddress(recipient, recipient));
            message.Subject = subject;
            message.Body = new TextPart("plain") { Text = body };

            using (var client = new SmtpClient())
            {
                client.Connect("smtp.gmail.com", 587, MailKit.Security.SecureSocketOptions.StartTls);
                client.Authenticate("", ""); //your email and your token
                client.Send(message);
                client.Disconnect(true);
            }

            Console.WriteLine($"Email sent to {recipient} successfully!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error sending email to {recipient}: {ex.Message}");
        }
    }

    static void UpdateExcelWithTimestamp(string filePath, string email)
    {
        try
        {
            using (var workbook = new XLWorkbook(filePath))
{
    var worksheet = workbook.Worksheet(1);
    var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip header row

    // Find all rows corresponding to the email and update the timestamp
    foreach (var row in rows)
    {
        if (row.Cell(2).GetString() == email) // Adjust the column index based on your sheet
        {
            row.Cell(4).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"); // Column 4 is "Date Sent"
        }
    }

    // Save the updated workbook
    workbook.SaveAs(filePath);
    Console.WriteLine($"Timestamps updated for all occurrences of {email}");
}

        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error updating Excel: {ex.Message}");
        }
    }
}

