using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using AdressBook_web_test;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using Newtonsoft.Json;
using Exel = Microsoft.Office.Interop.Excel;
using System.ComponentModel.Design;
using System.Text.RegularExpressions;

namespace Addressbook_test_data_generator { 

   class Progaram
{
        //qswqw
        static void Main(string[] args)
        {
            string formatData = args[0];
            int count = Convert.ToInt32(args[1]);
            string FileName = args[2];
            string format = args[3];
            if (formatData == "group")
            {
                List<GroupData> groups = new List<GroupData>();
                for (int i = 0; i < count; i++)
                {
                    groups.Add(new GroupData(TestBase.GenerateRandomString(10))
                    {
                        Header = TestBase.GenerateRandomString(12),
                        Footer = TestBase.GenerateRandomString(15)

                    });

                }
                if (format == "excel")
                {
                    writeToGroupExelfile(groups, FileName);
                }
                else
                {
                    StreamWriter writer = new StreamWriter(FileName);
                    if (format == "csv")
                    {
                        writeToGroupCSVfile(groups, writer);
                        writer.Close();
                    }
                    else if (format == "xml")
                    {
                        writeGroupToXMLfile(groups, writer);
                        writer.Close();


                    }
                    else if (format == "json")
                    {
                        writeGroupToJsonfile(groups, writer);
                        writer.Close();

                    }
                    else
                    {
                        System.Console.Out.WriteLine("Unrecognized format" + format);
                    }
                }



            }
            else if (formatData == "contact")

            {
                List<ContactData> contacts = new List<ContactData>();//о
                for (int i = 0; i < count; i++)
                {
                    contacts.Add(new ContactData(TestBase.GenerateRandomString(10), TestBase.GenerateRandomString(15))
                    {
                        Address = TestBase.GenerateRandomString(20),
                        Email2 = TestBase.GenerateRandomString(30),
                        WorkPhone = TestBase.GenerateRandomString(12)
                    });
                }

                if (format == "excel")
                {
                    writeToContactExelfile(contacts, FileName);
                }
                else
                {
                    StreamWriter writer = new StreamWriter(FileName);
                    if (format == "csv")
                    {
                        writeToContactCSVfile(contacts, writer);
                        writer.Close();
                    }
                    else if (format == "xml")
                    {
                        writeToContactXmlfile(contacts, writer);
                        writer.Close();

                    }
                    else if (format == "json")
                    {
                        writeToContactJsonfile(contacts, writer);
                        writer.Close();

                    }
                    else
                    {
                        System.Console.Out.WriteLine("Unrecognized format" + format);
                    }
                }



            }



            static void writeToContactExelfile(List<ContactData> contacts, string FileName)
            {
                Exel.Application app = new Exel.Application();
                app.Visible = true;
                Exel.Workbook wb = app.Workbooks.Add();
                Exel.Worksheet sheet = wb.ActiveSheet;

                int row = 1;
                foreach (ContactData contact in contacts)
                {
                    sheet.Cells[row, 1] = contact.Address;
                    sheet.Cells[row, 2] = contact.Email2;
                    sheet.Cells[row, 3] = contact.WorkPhone;
                    row++;
                }
                string fullPath = Path.Combine(Directory.GetCurrentDirectory(), FileName);
                File.Delete(fullPath);
                wb.SaveAs(fullPath);
                wb.Close();
                app.Visible = false;
                app.Quit();
                //уццу

            }

            static void writeToGroupExelfile(List<GroupData> groups, string FileName)
            {


                Exel.Application app = new Exel.Application();
                app.Visible = true;
                Exel.Workbook wb = app.Workbooks.Add();
                Exel.Worksheet sheet = wb.ActiveSheet;

                int row = 1;
                foreach (GroupData group in groups)
                {
                    sheet.Cells[row, 1] = group.Name;
                    sheet.Cells[row, 2] = group.Header;
                    sheet.Cells[row, 3] = group.Footer;
                    row++;
                }
                string fullPath = Path.Combine(Directory.GetCurrentDirectory(), FileName);
                File.Delete(fullPath);
                wb.SaveAs(fullPath);
                wb.Close();
                app.Visible = false;
                app.Quit();

                //throw new NotImplementedException();
            }

            static void writeToGroupCSVfile(List<GroupData> groups, StreamWriter writer)
            {
                foreach (GroupData group in groups)
                {
                    writer.WriteLine(String.Format("{0},{1},{2}",
                        group.Name, group.Header, group.Footer));
                }
            }




            static void writeGroupToXMLfile(List<GroupData> groups, StreamWriter writer)
            {
                new XmlSerializer(typeof(List<GroupData>)).Serialize(writer, groups);

            }

            static void writeGroupToJsonfile(List<GroupData> groups, StreamWriter writer)
            {
                writer.Write(JsonConvert.SerializeObject(groups, Newtonsoft.Json.Formatting.Indented));


            }

        }

         static void writeToContactJsonfile(List<ContactData> contacts, StreamWriter writer)
        {
            writer.Write(JsonConvert.SerializeObject(contacts, Newtonsoft.Json.Formatting.Indented));
        }

         static void writeToContactXmlfile(List<ContactData> contacts, StreamWriter writer)
        {
            new XmlSerializer(typeof(List<ContactData>)).Serialize(writer, contacts);
        }
        

        static void writeToContactCSVfile(List<ContactData> contacts, StreamWriter writer)
        {
            foreach (ContactData contact in contacts)
            {
                writer.WriteLine(String.Format("{0},{1},{2}",
                    contact.Address, contact.Email2, contact.WorkPhone));
            }
        }
    }

}


