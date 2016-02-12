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
            csvBuilder.Append("MessageName\tNote\tEventDateTime\tPayload\r\n");
            foreach (var file in files)
            {
                var doc = XDocument.Load(file.FullName);
                var messageName = doc.Descendants().First(d => d.Name.LocalName == "Name").Value;
                var payload = doc.Descendants().First(d => d.Name.LocalName == "Payload").Value;
                var payloadNote = XDocument.Parse(payload).Descendants().FirstOrDefault(d => d.Name.LocalName == "Note").Value;
                var eventDate = DateTime.Parse(doc.Descendants().First(d => d.Name.LocalName == "EventDateTime").Value);

                csvBuilder.AppendFormat("{0}\t{1}\t{2}\t{3}\r\n", messageName, payloadNote, eventDate, payload);
            }

            Clipboard.SetData(DataFormats.StringFormat, csvBuilder.ToString());
            Console.WriteLine("CSV Copied to Clipboard! Press any key to exit.");
            Console.ReadKey();
        }
    }
}
