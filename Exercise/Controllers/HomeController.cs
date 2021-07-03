using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Exercise.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MailKit.Net.Smtp;
using MimeKit;
using Microsoft.AspNetCore.Hosting;

namespace Exercise.Controllers
{

  
    [Authorize]
    public class HomeController : Controller
    {
        //attach the attachment
        IHostingEnvironment env = null;
              
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger, IHostingEnvironment env)
        {
            this.env = env;
            _logger = logger;
        }

        [Authorize(Roles = "Admin")]
        public IActionResult Index()
        {
            return View();
        }

        [Authorize(Roles = "Guest")]
        public IActionResult Privacy()
        {
            return View();
        }

        [Authorize(Roles = "Admin")]
        public IActionResult ReadCreateExcel()
        {

            // initialise the array
            int[] values = new int[2];
            //set count to 0
            int i = 0;
            //multiplied value
            int multipliedValue = 0;

            _logger.LogInformation("logging is working");
            //read excel file
            try
            {
                
                //Lets open the existing excel file and read through its content . Open the excel using openxml sdk
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open("source.xlsx", false))
                {
                    
                    //create the object for workbook part  
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    Sheets thesheetcollection = workbookPart.Workbook.GetFirstChild<Sheets>();
                    StringBuilder excelResult = new StringBuilder();

                    //using for each loop to get the sheet from the sheetcollection  
                    foreach (Sheet thesheet in thesheetcollection)
                    {
                        excelResult.AppendLine("Excel Sheet Name : " + thesheet.Name);
                        excelResult.AppendLine("----------------------------------------------- ");
                        //statement to get the worksheet object by using the sheet id  
                        Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById(thesheet.Id)).Worksheet;

                        SheetData thesheetdata = (SheetData)theWorksheet.GetFirstChild<SheetData>();
                        foreach (Row thecurrentrow in thesheetdata)
                        {
                            foreach (Cell thecurrentcell in thecurrentrow)
                            {
                                //statement to take the integer value  
                                string currentcellvalue = string.Empty;
                                if (thecurrentcell.DataType != null)
                                {
                                    if (thecurrentcell.DataType == CellValues.SharedString)
                                    {
                                        int id;
                                        if (Int32.TryParse(thecurrentcell.InnerText, out id))
                                        {
                                            SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                            if (item.Text != null)
                                            {
                                                //code to take the string value  
                                                excelResult.Append(item.Text.Text + " ");
                                            }
                                            else if (item.InnerText != null)
                                            {
                                                currentcellvalue = item.InnerText;
                                            }
                                            else if (item.InnerXml != null)
                                            {
                                                currentcellvalue = item.InnerXml;
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    _logger.LogInformation("i = " + i.ToString());
                                    //storing the values in an array
                                    values[i] = Convert.ToInt16(thecurrentcell.InnerText);

                                    _logger.LogInformation("value array= " + values[i]);

                                    excelResult.Append(Convert.ToInt16(thecurrentcell.InnerText) + " ");
                                    //increment i
                                    i++;
                                }
                            }
                            excelResult.AppendLine();
                        }
                        excelResult.Append("");
                        _logger.LogInformation(excelResult.ToString());
                       
                    }// end foreach collection

                    //performing the multiplication
                   // int multipliedValue =0;
                    for (int j = 0; j < values.Length; j++) {

                        if (j == 0)
                            multipliedValue = values[j];
                        else
                            multipliedValue = values[j] * multipliedValue;
                    }

                     //  _logger.LogInformation("sum = ");
                     //_logger.LogInformation(multipliedValue.ToString());
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;

            }

           //create Excel file
            using (SpreadsheetDocument document = SpreadsheetDocument.Create("Result.xlsx", SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };

                sheets.Append(sheet);

                Row newRow = new Row();
                Cell cell = new Cell();
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(multipliedValue.ToString());
                newRow.AppendChild(cell);
                   
                sheetData.AppendChild(newRow);
                workbookPart.Workbook.Save();
            }


            // sending email            
            MimeMessage message = new MimeMessage();

            //sender email
            MailboxAddress from = new MailboxAddress("Admin",
            "sawarnabeel@gmail.com");
            message.From.Add(from);

            //receiver email
            MailboxAddress to = new MailboxAddress("Nabeel",
            "sawarnabeel@gmail.com");
            message.To.Add(to);

            message.Subject = "Attached multiplication result of two numbers";

            BodyBuilder bodyBuilder = new BodyBuilder();
            bodyBuilder.HtmlBody = "Hi Anoop, <br/> Please find attachehd result, regarding the small exercise you gave me. <br/>" +
                " Its working. I will be sending you the code by tomorrow. Please confirm if you have received this email. <br/>" +
                " **This is an autogenerated mail**";
            // bodyBuilder.TextBody = "Its working!";

            bodyBuilder.Attachments.Add(env.ContentRootPath + "\\Result.xlsx");

            message.Body = bodyBuilder.ToMessageBody();

            SmtpClient client = new SmtpClient();
            client.Connect("smtp.gmail.com", 465, true);
            client.Authenticate("sawarnabeel@gmail.com", "Enter your gmail password here");
            //send the email
            client.Send(message);
            client.Disconnect(true);
            client.Dispose();

            return View();
        }

       /* Separate view to send email separately
        * 
        * [Authorize(Roles = "Guest")]
        public IActionResult SendEmail()
        {
            //send email goes here
            MimeMessage message = new MimeMessage();

            MailboxAddress from = new MailboxAddress("Admin",
            "sawarnabeel@gmail.com");
            message.From.Add(from);

            MailboxAddress to = new MailboxAddress("Nabeel",
            "sawarnabeel@gmail.com");
            message.To.Add(to);

            message.Subject = "Attached multiplication result of two numbers";

            BodyBuilder bodyBuilder = new BodyBuilder();
            bodyBuilder.HtmlBody = "Hi Anoop, <br/> Please find attachehd result, regarding the small exercise you gave me. <br/>" +
                " Its working. I will be sending you the code by tomorrow. Please confirm if you have received this email. <br/>" +
                " **This is an autogenerated mail**";
           // bodyBuilder.TextBody = "Its working!";

            bodyBuilder.Attachments.Add(env.ContentRootPath + "\\Result.xlsx");

            message.Body = bodyBuilder.ToMessageBody();

            SmtpClient client = new SmtpClient();
            client.Connect("smtp.gmail.com", 465, true);
            client.Authenticate("sawarnabeel@gmail.com", "DERNIER+PWD11");
            //send the email
            client.Send(message);
            client.Disconnect(true);
            client.Dispose();


            return View();
        } */

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
