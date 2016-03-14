using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using CsvHelper;
using CsvHelper.Configuration;

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
                if (result == DialogResult.Cancel)
                {
                    Console.WriteLine("Canceled. Press any key to exit.");
                    Console.ReadKey();
                    return;
                }
                result = ofd.ShowDialog();
            }

            var interestingHeaders = new List<string>()
            {
                "NServiceBus.OriginatingMachine",
                "NServiceBus.OriginatingEndpoint",
                "NServiceBus.ProcessingMachine",
                "NServiceBus.ProcessingEndpoint",
                "NServiceBus.EnclosedMessageTypes",
                "NServiceBus.ExceptionInfo.StackTrace"
            };

            var folder = new DirectoryInfo(ofd.SelectedPath);
            var files = folder.GetFiles();

            //var exportStringBuilder = new StringBuilder();
            using (var exportStringBuilder = new StringWriter())
            {
                var csv = new CsvWriter(exportStringBuilder, new CsvConfiguration() { HasHeaderRecord = true, QuoteAllFields = true, });

                csv.WriteField("EventDateTime");
                csv.WriteField("PON");
                csv.WriteField("LocationID");
                csv.WriteField("MessageName");
                csv.WriteField("MessageType");
                csv.WriteField("Note");
                csv.WriteField("DeadletterAck");
                csv.WriteField("MessageID");
                csv.WriteField("Payload");
                csv.WriteField("StackTrace");
                csv.WriteField("FileName");
                csv.WriteField("Origin");
                csv.WriteField("Processor");
                csv.NextRecord();

                // Todo: Add support to process ServiceControl errors?
                Console.Write("Processing...");
                foreach (var file in files)
                {
                    if (file.Length <= 0) continue;
                    var xml =
                        File.ReadAllText(file.FullName)
                            .Replace(((char)226).ToString(), "&#150;")
                            .Replace(((char)227).ToString(), "&#151;");
                    XDocument doc = null;
                    try
                    {
                        doc = XDocument.Parse(xml);
                    }
                    catch (Exception)
                    {
                        Console.WriteLine("Can't parse XML content in file: " + file.FullName);
                        continue;
                    }
                    if (doc.Root == null)
                    {
                        Console.WriteLine("Can't parse message with no content: " + file.FullName);
                        continue;
                    }
                    XDocument message = null;
                    try
                    {
                        message = XDocument.Parse(doc.Descendants().First(d => d.Name.LocalName == "Content").Value);
                    }
                    catch (Exception)
                    {
                        Console.WriteLine("Can't parse XML inside Content element in file: " + file.FullName);
                        continue;
                    }

                    var headers =
                        doc.Root?.Element("Headers")
                            .Elements("Header")
                            .Where(d => interestingHeaders.Contains(d.Element("Name")?.Value))
                            .ToDictionary(d => d.Element("Name")?.Value,
                                d => d.Element("Value")?.Value);

                    var messageType = headers["NServiceBus.EnclosedMessageTypes"].Split(new char[] { ',' }).First();
                    if (messageType.StartsWith("Mps.TeRx.Exceptions.Notification.") || messageType == "Mps.TeRx.Exceptions.UnhandledMpsrxConnectorServiceRxFillCommandExceptionNotification")
                    {
                        Console.WriteLine("Ignoring error sending notifications: " + file.FullName);
                        continue;
                    }

                    var payload =
                        message.Descendants()
                            .FirstOrDefault(d => d.Name.LocalName == "Payload" | d.Name.LocalName == "MessagePayload")?
                            .Value;
                    if (string.IsNullOrEmpty(payload))
                    {
                        Console.WriteLine("Ignoring message with no payload: " + file.FullName);
                        continue;
                    }

                    var stackTrace = headers.ContainsKey("NServiceBus.ExceptionInfo.StackTrace")
                        ? headers["NServiceBus.ExceptionInfo.StackTrace"]
                        : "";

                    var processingEndpoint = headers.ContainsKey("NServiceBus.ProcessingEndpoint") ? headers["NServiceBus.ProcessingEndpoint"] : "";
                    var processingMachine = headers.ContainsKey("NServiceBus.ProcessingMachine") ? headers["NServiceBus.ProcessingMachine"] : "";

                    var payloadXDoc = XDocument.Parse(payload);
                    DateTime eventDate =
                        DateTime.Parse(message.Descendants().First(d => d.Name.LocalName == "EventDateTime").Value);
                    string firstMessageId =
                        payloadXDoc.Descendants().FirstOrDefault(d => d.Name.LocalName == "MessageID")?.Value;

                    string locationid = string.Empty;
                    string pon = string.Empty;
                    string messageNote = string.Empty;
                    string payloadType = string.Empty;
                    string deadletterAck = doc.Root.Element("Class")?.Value;
                    switch (messageType)
                    {
                        case "Mps.TeRx.Service.Command.RxFillCommand":
                            payloadType = "RxFill";
                            var notFilledElement = payloadXDoc.Descendants().FirstOrDefault(d => d.Name.LocalName == "NotFilled");
                            if (notFilledElement != null)
                            {
                                messageNote = payloadXDoc.Descendants().FirstOrDefault(d => d.Name.LocalName == "Note")?.Value;
                            }
                            pon = payloadXDoc.Descendants().FirstOrDefault(d => d.Name.LocalName == "PrescriberOrderNumber")?.Value;
                            locationid = payloadXDoc.Descendants().FirstOrDefault(d => d.Name.LocalName == "To")?.Value;
                            break;
                        case "Mps.TeRx.Service.Event.OutboundPharmacyMessage.OutboundPharmacyMessageEvent":
                            payloadType = message.Descendants().FirstOrDefault(d => d.Name.LocalName == "MessageControlType")?.Value;
                            pon = payloadXDoc.Descendants().FirstOrDefault(d => d.Name.LocalName == "PrescriberOrderNumber")?.Value;
                            locationid = payloadXDoc.Descendants().FirstOrDefault(d => d.Name.LocalName == "From")?.Value;
                            break;
                        case "Mps.TeRx.Service.Command.PayloadCommand":
                            deadletterAck = doc.Root.Element("Class")?.Value;
                            payloadType = payloadXDoc.Descendants().FirstOrDefault(d => d.Name.LocalName == "Body")?.Elements().FirstOrDefault()?.Name.LocalName ?? payloadXDoc.Root.Name.LocalName;
                            pon = payloadXDoc.Descendants().FirstOrDefault(d => d.Name.LocalName == "PrescriberOrderNumber")?.Value;
                            locationid = payloadXDoc.Descendants().FirstOrDefault(d => d.Name.LocalName == "FacilityID")?.Value;
                            break;
                        case "Mps.TeRx.Service.Command.OutboundPharmacyMessage.OutboundPharmacyMessageResponseQueuedForPersistenceCommand":
                            messageNote = doc.Root.Element("Class")?.Value;
                            payloadType = message.Descendants().FirstOrDefault(d => d.Name.LocalName == "MessageControlType")?.Value;
                            pon = payloadXDoc.Descendants().FirstOrDefault(d => d.Name.LocalName == "PrescriberOrderNumber")?.Value;
                            locationid = payloadXDoc.Descendants().FirstOrDefault(d => d.Name.LocalName == "From")?.Value;
                            break;
                        case "Mps.TeRx.Service.Command.OutboundPharmacyMessage.OutboundPharmacyMessageQueuedForEndpointDeliveryCommand":
                            messageNote = doc.Root.Element("Class")?.Value;
                            payloadType = message.Descendants().FirstOrDefault(d => d.Name.LocalName == "MessageControlType")?.Value;
                            pon = payloadXDoc.Descendants().FirstOrDefault(d => d.Name.LocalName == "PrescriberOrderNumber")?.Value;
                            locationid = payloadXDoc.Descendants().FirstOrDefault(d => d.Name.LocalName == "From")?.Value;
                            break;
                            // Todo: add parsing for notification messages??
                        default:
                            break;
                    }

                    csv.WriteField(eventDate);
                    csv.WriteField(pon);
                    csv.WriteField(locationid);
                    csv.WriteField(messageType);
                    csv.WriteField(payloadType);
                    csv.WriteField(messageNote);
                    csv.WriteField(deadletterAck);
                    csv.WriteField(firstMessageId);
                    csv.WriteField(payload.Length > 30000 ? payload.Remove(30000) : payload);
                    csv.WriteField(stackTrace);
                    csv.WriteField(file.FullName);
                    csv.WriteField($"{headers["NServiceBus.OriginatingEndpoint"]}@{headers["NServiceBus.OriginatingMachine"]}");
                    csv.WriteField($"{processingEndpoint}@{processingMachine}");
                    csv.NextRecord();
                }

                var outputFileDialog = new SaveFileDialog()
                {
                    Title = "Select Output File",
                    FileName = "Output.csv",
                    AddExtension = true,
                    Filter = "CSV|*.csv"
                };
                var outputFileResult = outputFileDialog.ShowDialog();
                while (outputFileResult != DialogResult.OK)
                {
                    if (outputFileResult == DialogResult.Cancel)
                    {
                        Console.WriteLine("Canceled. Press any key to exit.");
                        Console.ReadKey();
                        return;
                    }
                    outputFileResult = ofd.ShowDialog();
                }

                try
                {
                    using (var streamWriter = new StreamWriter(outputFileDialog.FileName))
                    {
                        streamWriter.Write(exportStringBuilder);
                    }
                }
                catch (Exception ex)
                {
                    var regColor = Console.ForegroundColor;
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    Console.WriteLine($"Error saving output: {ex.Message}");
                    Console.WriteLine("File output has been saved to your clipboard instead.");
                    Clipboard.SetData(DataFormats.Text, exportStringBuilder.ToString());
                    Console.ForegroundColor = regColor;
                }
            }

            Console.WriteLine("All done! Press any key to exit.");
            Console.ReadKey();
        }
    }
}
