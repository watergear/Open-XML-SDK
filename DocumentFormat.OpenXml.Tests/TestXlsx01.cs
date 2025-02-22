﻿using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace DocumentFormat.OpenXml.Tests
{
    public class XlsxTests01
    {
        public static string s_TestFileLocation = "../../../TestFiles/";

        [Fact]
        public void X008_XlsxCreation_Package_Settings()
        {
            var fiSource = new FileInfo(Path.Combine(s_TestFileLocation, "Spreadsheet.xlsx"));
            var fiCopy = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, Guid.NewGuid().ToString() + ".xlsx"));
            File.Copy(fiSource.FullName, fiCopy.FullName);
            using (Package package = Package.Open(fiCopy.FullName, FileMode.Open, FileAccess.ReadWrite))
            {
                OpenSettings openSettings = new OpenSettings();
                openSettings.MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(MarkupCompatibilityProcessMode.ProcessAllParts, FileFormatVersions.Office2013);
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(package, openSettings))
                {
                    OpenXmlValidator v = new OpenXmlValidator(FileFormatVersions.Office2013);
                    var errs = v.Validate(doc);
                    Assert.Equal(1, errs.Count());
                }
            }
            if (TestUtil.DeleteTempFiles)
                fiCopy.Delete();
        }

        [Fact]
        public void X007_SpreadsheetDocument_Open()
        {
            var fiSource = new FileInfo(Path.Combine(s_TestFileLocation, "Spreadsheet.xlsx"));
            var fiCopy = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, Guid.NewGuid().ToString() + ".xlsx"));
            File.Copy(fiSource.FullName, fiCopy.FullName);
            OpenSettings openSettings = new OpenSettings();
            openSettings.MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(MarkupCompatibilityProcessMode.ProcessAllParts, FileFormatVersions.Office2013);
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fiCopy.FullName, true, openSettings))
            {
                OpenXmlValidator v = new OpenXmlValidator(FileFormatVersions.Office2013);
                var errs = v.Validate(doc);
                Assert.Equal(1, errs.Count());
            }
            if (TestUtil.DeleteTempFiles)
                fiCopy.Delete();
        }

        [Fact]
        public void X006_Xlsx_DeleteAdd_CoreExtendedProperties()
        {
            var docName = "Spreadsheet.xlsx";
            var ba = File.ReadAllBytes(s_TestFileLocation + docName);
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(ba, 0, ba.Length);
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(ms, true))
                {
                    var corePart = doc.CoreFilePropertiesPart;
                    var appPart = doc.ExtendedFilePropertiesPart;
                    doc.DeletePart(corePart);
                    doc.DeletePart(appPart);
                    doc.AddCoreFilePropertiesPart();
                    doc.AddExtendedFilePropertiesPart();
                    doc.AddCustomFilePropertiesPart();
                    doc.AddDigitalSignatureOriginPart();
                    doc.AddExtendedPart("relType", "contentType/xml", ".xml");
                    var tnPart = doc.AddThumbnailPart(ThumbnailPartType.Jpeg);
                    doc.DeletePart(tnPart);
                    tnPart = doc.AddThumbnailPart("image/jpg");
                    OpenXmlValidator v = new OpenXmlValidator(FileFormatVersions.Office2013);
                    var errs = v.Validate(doc);
                    Assert.Equal(1, errs.Count());
                }
            }
        }

        [Fact]
        public void X005_XlsxCreation_Package_Settings()
        {
            var fiSource = new FileInfo(Path.Combine(s_TestFileLocation, "Spreadsheet.xlsx"));
            var fiCopy = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, Guid.NewGuid().ToString() + ".xlsx"));
            File.Copy(fiSource.FullName, fiCopy.FullName);
            using (Package package = Package.Open(fiCopy.FullName, FileMode.Open, FileAccess.ReadWrite))
            {
                OpenSettings openSettings = new OpenSettings();
                openSettings.MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(MarkupCompatibilityProcessMode.ProcessAllParts, FileFormatVersions.Office2013);
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(package, openSettings))
                {
                    OpenXmlValidator v = new OpenXmlValidator(FileFormatVersions.Office2013);
                    var errs = v.Validate(doc);
                    Assert.Equal(1, errs.Count());
                }
            }
            if (TestUtil.DeleteTempFiles)
                fiCopy.Delete();
        }

        [Fact]
        public void X004_SpreadsheetDocument_Open()
        {
            var fiSource = new FileInfo(Path.Combine(s_TestFileLocation, "Spreadsheet.xlsx"));
            var fiCopy = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, Guid.NewGuid().ToString() + ".xlsx"));
            File.Copy(fiSource.FullName, fiCopy.FullName);
            OpenSettings openSettings = new OpenSettings();
            openSettings.MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(MarkupCompatibilityProcessMode.ProcessAllParts, FileFormatVersions.Office2013);
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fiCopy.FullName, true, openSettings))
            {
                OpenXmlValidator v = new OpenXmlValidator(FileFormatVersions.Office2013);
                var errs = v.Validate(doc);
                Assert.Equal(1, errs.Count());
            }
            if (TestUtil.DeleteTempFiles)
                fiCopy.Delete();
        }

        [Fact]
        public void X003_XlsxCreation_Stream()
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
                Assert.Equal(0, errs.Count());
            }
        }

        [Fact]
        public void X002_XlsxCreation()
        {
            FileInfo fi = new FileInfo(Path.Combine(s_TestFileLocation, Guid.NewGuid().ToString() + ".docx"));
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument doc = SpreadsheetDocument.Create(fi.FullName, SpreadsheetDocumentType.Workbook);
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
            Assert.Equal(0, errs.Count());

            doc.Close();

            if (TestUtil.DeleteTempFiles)
                fi.Delete();
        }

        [Fact]
        public void X001_XlsxValidation()
        {
            var docName = "Spreadsheet.xlsx";
            XlsxValidationHelper(docName, 2);
        }

        private static void XlsxValidationHelper(string docName, int expectedErrorCount)
        {
            var ba = File.ReadAllBytes(s_TestFileLocation + docName);
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(ba, 0, ba.Length);
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(ms, true))
                {
                    OpenXmlValidator validator = new OpenXmlValidator(FileFormatVersions.Office2007);
                    int cnt = validator.Validate(doc).Count();
                    validator = new OpenXmlValidator(FileFormatVersions.Office2010);
                    cnt += validator.Validate(doc).Count();
                    validator = new OpenXmlValidator(FileFormatVersions.Office2013);
                    cnt += validator.Validate(doc).Count();
                    Assert.Equal(expectedErrorCount, cnt);
                }
            }
        }
    }
}
