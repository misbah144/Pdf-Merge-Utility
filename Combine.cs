using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace AnnualPack
{
    class Combine
    {
        string Statement = ConfigurationManager.AppSettings["Statement"].ToString(),
            Letter = ConfigurationManager.AppSettings["Letter"].ToString(),
            TaxCert = ConfigurationManager.AppSettings["TaxCert"].ToString(),
            PerformanceSumm = ConfigurationManager.AppSettings["PerformanceSummary"].ToString(),
            destination = ConfigurationManager.AppSettings["destination"].ToString();
        string Marketing = ConfigurationManager.AppSettings["Marketing"].ToString();
        
        public void Start()
        {
            string[] customer = new DirectoryInfo(Statement).GetFiles().Select(o => o.Name).ToArray();


            //foreach (var item in customer)
            //{
                using (PdfDocument finalPDF = new PdfDocument())
                {
                    ArrayList file = new ArrayList();
                    
                   // string[]myfile=
                    string[] myfiles = new DirectoryInfo(Statement).GetFiles().Select(o => o.FullName).ToArray();


                    //file.Add(Directory.GetFiles(PerformanceSumm, item));
                    //file.Add(Directory.GetFiles(Marketing, "Al Meezan Marketing Letter.pdf"));  // adding generic letter on demand
                    //file.Add(Directory.GetFiles(Letter, item));
                    file.Add(Directory.GetFiles(Statement));
                  //  file.Add(Directory.GetFiles(TaxCert, item));
                   // foreach (string[] filePath in file)

                    int last=myfiles.Count();

                    //string  first_name = myfiles[0].ToString();
                    //string  last_name = myfiles[last-1].ToString();

                    for (int i = 0; i < myfiles.Count();i++ )
                    {
                        string name = myfiles[i].ToString();
                        //  if (filePath.Length > 0)
                        //if (file[i].Length > 0)
                        //{
                        using (PdfDocument pdf = PdfReader.Open(myfiles[i].ToString(), PdfDocumentOpenMode.Import))
                        {
                            CopyPages(pdf, finalPDF);
                            //File.Delete(item.ToString());
                        }
                        //}
                    }
                    //if (File.Exists(destination+"merge" + ".pdf"))
                    //{
                    //    File.Delete(destination + "merge" +".pdf");
                    //}
                   // finalPDF.Save(destination + item + ".pdf");
                    finalPDF.Save(destination + "merge" +".pdf");                
                }
           // }                           
        }

        public void StatementCombine()
        {
            string[] customer = new DirectoryInfo(Statement).GetFiles().Select(o => o.Name.Substring(0, 6)).ToArray();
            IEnumerable<string> Uniqcustomer = customer.Distinct<string>();

            foreach (var item in Uniqcustomer)
            {
                using (PdfDocument finalPDF = new PdfDocument())
                {
                    ArrayList file = new ArrayList();
                    file.Add(Directory.GetFiles(Statement, item + "*"));
                    string[] obj = (string[])file[0];
                    foreach (string filePath in obj)
                    {
                        if (File.Exists(filePath))
                        {
                            using (PdfDocument pdf = PdfReader.Open(filePath.ToString(), PdfDocumentOpenMode.Import))
                            {

                                CopyPages(pdf, finalPDF);
                                File.Delete(filePath.ToString());
                            }
                        }
                    }
                    finalPDF.Save(Statement + item + ".pdf");
                }
            }
        }

        private void CopyPages(PdfDocument from, PdfDocument to)
        {
            for (int i = 0; i < from.PageCount; i++)
            {
                to.AddPage(from.Pages[i]);
            }
        }
    }
}
