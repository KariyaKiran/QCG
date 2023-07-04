using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace QcGoldArchive
{
    public class XmlUtility
    {

        public static void WriteSettings(string[] settings,string path)   //write to settings.xml file
        {
            XmlDocument xmlDoc = new XmlDocument();
            XmlNode root;
            XmlNode node;

            LanguageManagement LM = LanguageManagement.CreateInstance();
            
            root = xmlDoc.CreateElement("Settings");
            xmlDoc.AppendChild(root);

            node = xmlDoc.CreateElement("Port");
            node.InnerText = settings[0];
            root.AppendChild(node);

            node = xmlDoc.CreateElement("Printer");
            node.InnerText = settings[1];
            root.AppendChild(node);

            node = xmlDoc.CreateElement("ExportFile");
            node.InnerText = settings[2];
            root.AppendChild(node);

            node = xmlDoc.CreateElement("AutoPrint");
            node.InnerText = settings[3];
            root.AppendChild(node);

            node = xmlDoc.CreateElement("usedefault");
            node.InnerText = settings[4];
            root.AppendChild(node);

            node = xmlDoc.CreateElement("DateFormat");
            node.InnerText = settings[5];
            root.AppendChild(node);

            node = xmlDoc.CreateElement("Language");
            node.InnerText = settings[6];
            root.AppendChild(node);

            node = xmlDoc.CreateElement("ResultsPrint");
            node.InnerText = settings[7];
            root.AppendChild(node);

            node = xmlDoc.CreateElement("HeaderSpace");
            node.InnerText = settings[8];
            root.AppendChild(node);

            node = xmlDoc.CreateElement("HeaderSpaceValue");
            node.InnerText = settings[9];
            root.AppendChild(node);

            node = xmlDoc.CreateElement("Digitalsignaturepath");
            node.InnerText = settings[10];
            root.AppendChild(node);

            node = xmlDoc.CreateElement("ReportType");
            node.InnerText = settings[11];
            root.AppendChild(node);

            node = xmlDoc.CreateElement("title");
            node.InnerText = settings[12];
            root.AppendChild(node);


            node = xmlDoc.CreateElement("pdfautoprint");
            node.InnerText = settings[13];
            root.AppendChild(node);


            node = xmlDoc.CreateElement("subtitleval");
            node.InnerText = settings[14];
            root.AppendChild(node);

            node = xmlDoc.CreateElement("optionalValue");
            node.InnerText = settings[15];
            root.AppendChild(node);

            node = xmlDoc.CreateElement("optionalValue2");
            node.InnerText = settings[16];
            root.AppendChild(node);

            node = xmlDoc.CreateElement("removeFooter");
            node.InnerText = settings[17];
            root.AppendChild(node);

            xmlDoc.Save(path);
        }

        public static string[] ReadSettings(string path)          //read from settings.xml file
        {
            string[] xmlLabels = { "Port", "Printer", "ExportFile", "AutoPrint", "usedefault", "DateFormat", "Language", "ResultsPrint", "HeaderSpace", "HeaderSpaceValue", "Digitalsignaturepath", "ReportType", "title", "pdfautoprint","subtitleval", "optionalValue","optionalValue2","removeFooter"};

            string[] settings = new string[18];

            XmlDocument xmlDoc = new XmlDocument();

            xmlDoc.Load(path);

            for (int i = 0; i < settings.Length; i++)
            {
                settings[i] = xmlDoc.GetElementsByTagName(xmlLabels[i])[0].InnerText;
            }

            return settings;
        }

       

    }
}
