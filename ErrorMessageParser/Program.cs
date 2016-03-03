using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace ErrorMessageParser
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            var ofd = new FolderBrowserDialog()
            {
                Description = "Select Messages Folder",
                ShowNewFolderButton = false
            };

            var result = ofd.ShowDialog();
            while (result != DialogResult.OK)
            {
                result = ofd.ShowDialog();
            }

            var folder = new DirectoryInfo(ofd.SelectedPath);
            var files = folder.GetFiles();
            var csvBuilder = new StringBuilder();
            csvBuilder.Append("EventDateTime\tPON\tLocationID\tID-Concat\tMessageName\tNote\tMessageID\tPayload\tFileName\r\n");
            foreach (var file in files)
            {
                string fullname = file.FullName;
                if (file.Length <= 0) continue;
                var xml = File.ReadAllText(file.FullName).Replace(((char)226).ToString(), "&#150;").Replace(((char)227).ToString(), "&#151;");
                var doc = XDocument.Parse(xml);

                var messageName = doc.Descendants().FirstOrDefault(d => d.Name.LocalName == "Name")?.Value;
                if (string.IsNullOrWhiteSpace(messageName))
                    messageName =
                        doc.Descendants()
                            .FirstOrDefault(d => d.Name.LocalName == "OutboundPharmacyMessageEvent")?
                            .Name.LocalName;
                if (messageName == "Unhandled Mpact Connector Service Exception Notification")
                {
                    Console.WriteLine("Ignoring error sending notification email: " + file.FullName);
                    continue;
                }
                if (doc.Root.Descendants().First().Name.LocalName == "BSSMPSRxConnectorServiceExceptionNotification")
                {
                    Console.WriteLine("Ignoring error sending notification email: " + file.FullName);
                    continue;
                }

                var payload = doc.Descendants().FirstOrDefault(d => d.Name.LocalName == "Payload" | d.Name.LocalName == "MessagePayload")?.Value;
                if (string.IsNullOrEmpty(payload))
                {
                    Console.WriteLine("Ignoring message with no payload: " + file.FullName);
                    continue;
                }

                var payloadXDoc = XDocument.Parse(payload);
                DateTime eventDate = DateTime.Parse(doc.Descendants().First(d => d.Name.LocalName == "EventDateTime").Value);
                string firstMessageId = payloadXDoc.Descendants().First(d => d.Name.LocalName == "MessageID").Value;

                string locationid = string.Empty;
                string pon = string.Empty;
                string payloadNote = string.Empty;
                switch (messageName)
                {
                    case "RxFill Command":
                        payloadNote = payloadXDoc.Descendants().FirstOrDefault(d => d.Name.LocalName == "Note")?.Value;
                        pon = payloadXDoc.Descendants().FirstOrDefault(d => d.Name.LocalName == "PrescriberOrderNumber")?.Value;
                        locationid = payloadXDoc.Descendants().FirstOrDefault(d => d.Name.LocalName == "To")?.Value;
                        break;
                    case "MPSRx Connector Service Invalid NCPDP Schema Notification":
                        payloadNote = doc.Descendants().FirstOrDefault(d => d.Name.LocalName == "ExceptionErrorMessage")?.Value;
                        locationid = payloadXDoc.Descendants("Facility").Descendants("NPI").FirstOrDefault()?.Value;
                        pon = doc.Descendants().FirstOrDefault(d => d.Name.LocalName == "OrderID")?.Value;
                        break;
                    case "OutboundPharmacyMessageEvent":
                        payloadNote = doc.Descendants().First(d => d.Name.LocalName == "MessageControlType")?.Value;
                        pon = payloadXDoc.Descendants().FirstOrDefault(d => d.Name.LocalName == "PrescriberOrderNumber")?.Value;
                        locationid = payloadXDoc.Descendants().FirstOrDefault(d => d.Name.LocalName == "From")?.Value;
                        break;
                }

                csvBuilder.AppendFormat("{0}\t{1}\t{2}\t{1}-{2}\t{3}\t{4}\t{5}\t{6}\t{7}\r\n", eventDate, pon, locationid, messageName, payloadNote, firstMessageId, payload, file.FullName);
            }

            Clipboard.SetData(DataFormats.StringFormat, csvBuilder.ToString());
            Console.WriteLine("CSV Copied to Clipboard! Press any key to exit.");
            Console.ReadKey();
        }
    }
}
