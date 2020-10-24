using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mime;
using System.Reflection.Metadata;
using System.Text.Encodings.Web;
using System.Threading;
using System.Web.Providers.Entities;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Html;
using static DocumentFormat.OpenXml.Packaging.WordprocessingDocument;
using OpenXmlPowerTools;
using com.sun.org.apache.xerces.@internal.parsers;
using ikvm.extensions;

namespace CrmEmailGenerator
{
    public class CrmEmailGenerator
    {
        private static  string firstNamePlaceHolder = "{First Name (Attendee Contact (Contact))}";
       // private static string firstNameDynamicValue = "<span class=\"ms-crm-DataSlug\" style=\"DISPLAY: inline;background-color: #FFFF33; border: 1px solid #808080\" tabindex=\"-1\" value=\"<slugelement type=&quot;slug&quot;><slug type=&quot;dynamic&quot; value=&quot;contact.cpi_attendeecontactid.firstname&quot;/></slugelement>\" title=\"{First Name(Attendee Contact (Contact))}\">{First Name(Attendee Contact (Contact))}</span> &thinsp; </span>\r\n";
        static void Main(string[] args)
        {
            var fileNotFoundException = false;
            var directoryNotFoundException = false;
            Console.WriteLine("Enter the file name of the word document you'd like to convert to html: ");
            var fileName = Console.ReadLine();
            do
            {
                if (directoryNotFoundException)
                {
                    Console.WriteLine("Folder not found. A folder named WordToHtmlEmails needs to be saved directly on your c drive");
                    break;
                }
                if (fileNotFoundException)
                {
                    Console.WriteLine("Enter the file name of the word document you'd like to convert to html: ");
                    fileName = Console.ReadLine();
                }

                try
                {
                    WriteWordDocToHtmlFile(fileName);
                    fileNotFoundException = false;
                }
                catch (DirectoryNotFoundException e)
                {
                    directoryNotFoundException = true;
                    fileNotFoundException = false;
                }
                catch (FileNotFoundException e)
                {
                    fileNotFoundException = true;
                }

            } while(fileNotFoundException || directoryNotFoundException);
        }

        private static void WriteWordDocToHtmlFile(string fileName)
        {
            byte[] byteArray = File.ReadAllBytes($@"c:\CRM Emails\{fileName}");
            using (MemoryStream memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                WordprocessingDocument doc = Open(memoryStream, true);
                
                HtmlConverterSettings settings = new HtmlConverterSettings()
                {
                    PageTitle = "My Page Title"
                };
           
                XElement html = HtmlConverter.ConvertToHtml(doc, settings);

                var fileNameWithoutExtension = fileName.Substring(0, fileName.IndexOf('.'));
                File.WriteAllText($@"c:\CRM Emails\{fileNameWithoutExtension}.html", html.ToString());
                
            }
        }
    }
}
