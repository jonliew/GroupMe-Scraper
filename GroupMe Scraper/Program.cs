using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using Newtonsoft.Json;
using OfficeOpenXml;

namespace GroupMe_Scraper
{
    class Program
    {
        private static readonly string groupId = "<INSERT YOUR GROUP ID HERE>";
        private static readonly string tokenId = "<INSERT YOUR GROUPME TOKEN ID HERE>";
        private static readonly string outputFileName = "output.xlsx";
        
        private static string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        private static readonly int limit = 100;

        static void Main(string[] args)
        {

  
            var lastId = "";
            var url = $"https://api.groupme.com/v3/groups/{groupId}/messages?token={tokenId}&limit={limit}&before_id={lastId}";

            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Messages");
                var worksheet = excel.Workbook.Worksheets["Messages"];
                var errorCode = 200;
                while (errorCode == 200)
                {
                    url = $"https://api.groupme.com/v3/groups/{groupId}/messages?token={tokenId}&limit={limit}&before_id={lastId}";
                    using (var wc = new WebClient())
                    {
                        var json = "";
                        try
                        {
                            json = wc.DownloadString(url);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                            break;
                        }
                        
                        //Console.WriteLine(json);
                        dynamic data = JsonConvert.DeserializeObject(json);
                        errorCode = data.meta.code;
                        if (errorCode != 200)
                        {
                            Console.WriteLine("Error code " + errorCode);
                            if (400 <= errorCode && errorCode < 500)
                            {
                                Console.WriteLine(data.meta.errors);
                            }
                        }
                        var messages = data.response.messages;
                        foreach (var message in messages)
                        {
                            string messageText = message.text;
                            foreach (var attachment in message.attachments)
                            {
                                if (attachment.type == "image")
                                {
                                    messageText += " " + attachment.url;
                                }
                            }

                            messageText = messageText.Replace("\n\n", " ");
                            var appendString = $"{message.id}|{message.name}|{message.sender_id}|{message.created_at}|{messageText}|{message.favorited_by.Count}";
                            lastId = message.id;


                            Console.WriteLine(appendString);

                            var messageContent = new List<string[]>()
                            {
                                new string[] {message.id, message.name, message.sender_id, message.created_at, messageText, message.favorited_by.Count.ToString()}
                            };

                            var row = worksheet.Dimension?.Rows ?? 0;
                            worksheet.Cells[row + 1, 1].LoadFromArrays(messageContent);
                            
                        }
                    }

                    System.Threading.Thread.Sleep(5000);
                }
                var excelFile = new FileInfo(desktopPath + "\\" + outputFileName);
                excel.SaveAs(excelFile);

            }
        }
    }
}
