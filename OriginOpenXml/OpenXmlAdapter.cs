using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace DocumentFormat.OpenXml
{
    [ComVisible(true)]
    [Guid("E2ADA7F8-63F6-4F05-9B99-F3A52BDF3D31")]
    public interface OpenXmlAdapter
    {
        void Dump();
    }

    [ComVisible(true)]
    [Guid("68CAE9A0-B50F-4706-9DE5-5A3B130487C7")]
    public class OpenXmlAdapterClass : OpenXmlAdapter
    {
        public void Dump()
        {
            using (MemoryStream ms = new MemoryStream())
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            using (SpreadsheetDocument doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookpart = doc.AddWorkbookPart();
                workbookpart.Workbook = new S.Workbook();
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new S.Worksheet(new S.SheetData());
                S.Sheets sheets = doc.WorkbookPart.Workbook.AppendChild<S.Sheets>(new S.Sheets());
                S.Sheet sheet = new S.Sheet()
                {
                    Id = doc.WorkbookPart.
                        GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "mySheet"
                };
                sheets.Append(sheet);
                workbookpart.Workbook.Save();

                OpenXmlValidator v = new OpenXmlValidator(FileFormatVersions.Office2013);
                var errs = v.Validate(doc);
                Console.WriteLine(errs.Count());
                //Assert.Equal(0, errs.Count());
            }

            Console.WriteLine("Hello");
        }
    }

}
