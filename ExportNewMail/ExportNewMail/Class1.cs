using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

//using System.Windows.Forms;
using Fiddler;
using Microsoft.Office.Interop.Outlook;

[assembly: Fiddler.RequiredVersion("4.3.9.9")]


namespace ExportNewMail
{

    [ProfferFormat("Send Sessions with Outlook", "Export Sessions to Outlook Attachments")]
    

    public class AsMailTranscoder : ISessionExporter  // Ensure class is public, or Fiddler won't see it!
    {
        public void Dispose()
        {
        }

        public bool ExportSessions(string sExportFormat, Session[] oSessions, Dictionary<string, object> dictOptions, EventHandler<ProgressCallbackEventArgs> evtProgressNotifications)
        {

            IList<string> attachments = new List<string>();

            foreach (Session item in oSessions)
            {
                string file = Path.Combine(Path.GetTempPath(), item.id + ".txt");                    
                File.WriteAllText(file, item.ToString());
                attachments.Add(file);


                XmlDocument doc = new XmlDocument();
                try {
                    doc.Load(new MemoryStream(item.RequestBody));

                    file = Path.Combine(Path.GetTempPath(), "request_" + item.id + ".xml");
                    doc.Save(file);

                    attachments.Add(file);

                    doc.Load(new MemoryStream(item.ResponseBody));

                    file = Path.Combine(Path.GetTempPath(), "response_" + item.id + ".xml");
                    doc.Save(file);

                    attachments.Add(file);

                } catch
                {
                    //dann nicht :-);
                }

            }


            Application oApp = new Application();
            _MailItem oMailItem = (_MailItem)oApp.CreateItem(OlItemType.olMailItem);
            
            
            foreach (var item in attachments)
            {
                oMailItem.Attachments.Add((object)item,OlAttachmentType.olEmbeddeditem, 1, (object)"Attachment");
            }

            oMailItem.Display();

            return true;
        }
    }
}
