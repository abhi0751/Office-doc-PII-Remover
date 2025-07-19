using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using System.IO.Packaging;

namespace PPTMetadataCleaner
{
    class Program
    {
        static void Main(string[] args)
        {
            bool forceClean = false;

            Console.WriteLine("Enter the folder path to scan for office files:");
            string folderPath = Console.ReadLine();

            Console.WriteLine("Do you want a forece clean: Y/N");


            string forececleaning = Console.ReadLine();

      if (forececleaning == "Y" || forececleaning == "y")
                {
                forceClean = true;
            }

            

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine("Invalid folder path.");
                return;
            }

            string[] extensions = new[] { ".docx", ".xlsx", ".pptx" };

            var allFiles = Directory.EnumerateFiles(folderPath, "*.*", SearchOption.AllDirectories)
                                    .Where(file => extensions.Any(ext => file.EndsWith(ext, StringComparison.OrdinalIgnoreCase)))
                                    .ToArray();
            Console.WriteLine($"Found {allFiles.Length} files.");

            var pptApp = new Microsoft.Office.Interop.PowerPoint.Application();
            var wordApp = new Microsoft.Office.Interop.Word.Application();
            var excelApp = new Microsoft.Office.Interop.Excel.Application();

            
            pptApp.Visible = MsoTriState.msoCTrue;

            foreach (string file in Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories)
                .Where(f => f.EndsWith(".pptx") || f.EndsWith(".docx") || f.EndsWith(".xlsx")))
            {
                Console.WriteLine($"\nProcessing: {file}");

                bool shouldClean = false;

                try
                {
                    if (file.EndsWith(".pptx"))
                    {
                        shouldClean = CleanWithInterop_PPT(pptApp, file, forceClean);
                        CleanWithOpenXml(file);
                    }
                    else if (file.EndsWith(".docx"))
                    {
                        shouldClean = CleanWithInterop_Word(wordApp, file, forceClean);
                        CleanWithOpenXml(file);
                    }
                    else if (file.EndsWith(".xlsx"))
                    {
                        shouldClean = CleanWithInterop_Excel(excelApp, file, forceClean);
                        CleanWithOpenXml(file);
                    }



                    if (shouldClean)
                    {
                        CleanWithOpenXml(file);
                        Console.WriteLine("✅ Cleaned successfully.");
                    }
                    else
                    {
                        Console.WriteLine("⏭ Skipped.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error: {ex.Message}");
                }
            }

            pptApp.Quit();
            wordApp.Quit();
            excelApp.Quit();
            Marshal.ReleaseComObject(pptApp);
            Marshal.ReleaseComObject(wordApp);
            Marshal.ReleaseComObject(excelApp);

            Console.WriteLine("\nAll files processed. Press any key to exit.");
            Console.ReadKey();
        }

        static bool CleanWithInterop_PPT(Microsoft.Office.Interop.PowerPoint.Application app, string path, bool force)
        {
            Presentation pres = null;
            try
            {
                pres = app.Presentations.Open(path, WithWindow: MsoTriState.msoFalse);

                if (!force && pres.RemovePersonalInformation == MsoTriState.msoTrue)
                {
                    Console.WriteLine("File will not capture any personal info .   Skipping it");
                    return false;
                }

                pres.RemovePersonalInformation = MsoTriState.msoTrue;
                dynamic props = pres.BuiltInDocumentProperties;
                ClearProperty(props, "Author");
                ClearProperty(props, "Last Author");
                ClearProperty(props, "Company");
                ClearProperty(props, "Manager");
                ClearProperty(props, "Comments");
                pres.Save();
                return true;
            }
            finally
            {
                pres?.Close();
                Marshal.ReleaseComObject(pres);
            }
        }

        static bool CleanWithInterop_Word(Microsoft.Office.Interop.Word.Application app, string path, bool force)
        {
            Document doc = null;
            try
            {
                doc = app.Documents.Open(path, ReadOnly: false, Visible: false);
                if (!force && doc.RemovePersonalInformation)
                {
                    Console.WriteLine("File will not capture any personal info .   Skipping it");
                    return false;
                }
                doc.RemovePersonalInformation = true;
                dynamic props = doc.BuiltInDocumentProperties;
                ClearProperty(props, "Author");
                ClearProperty(props, "Last Author");
                ClearProperty(props, "Company");
                ClearProperty(props, "Manager");
                ClearProperty(props, "Comments");
                doc.Save();
                return true;
            }
            finally
            {
                doc?.Close(false);
                Marshal.ReleaseComObject(doc);
            }
        }

        static bool CleanWithInterop_Excel(Microsoft.Office.Interop.Excel.Application app, string path, bool force)
        {
            Workbook wb = null;
            try
            {
                wb = app.Workbooks.Open(path);
                if (!force && wb.RemovePersonalInformation)
                {
                    Console.WriteLine("File will not capture any personal info .   Skipping it");
                    return false;
                }
                wb.RemovePersonalInformation = true;
                dynamic props = wb.BuiltinDocumentProperties;
                ClearProperty(props, "Author");
                ClearProperty(props, "Last Author");
                ClearProperty(props, "Company");
                ClearProperty(props, "Manager");
                ClearProperty(props, "Comments");
                wb.Save();
                return true;
            }
            finally
            {
                wb?.Close(false);
                Marshal.ReleaseComObject(wb);
            }
        }
        

        static void ClearProperty(dynamic props, string name)
        {
            try
            {
                var prop = props[name];
                if (prop != null && prop.Value != null)
                    prop.Value = "";
            }
            catch { /* Ignore */ }
        }

        static void CleanWithOpenXml(string filePath)
        {
            if (filePath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            {
                using (var doc = WordprocessingDocument.Open(filePath, true))
                {
                    RemoveCustomXmlParts(doc);
                    RemoveCustomProperties(doc);
                    RemoveExtendedProperties(doc);
                    RemovePackageProperties(doc);
                }
            }
            else if (filePath.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
            {
                using (var doc = PresentationDocument.Open(filePath, true))
                {
                    RemoveCustomXmlParts(doc);
                    RemoveCustomProperties(doc);
                    RemoveExtendedProperties(doc);
                    RemovePackageProperties(doc);
                }
            }
            else if (filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                using (var doc = SpreadsheetDocument.Open(filePath, true))
                {
                    RemoveCustomXmlParts(doc);
                    RemoveCustomProperties(doc);
                    RemoveExtendedProperties(doc);
                    RemovePackageProperties(doc);
                }
            }
            else
            {
                Console.WriteLine("❌ Unsupported file type.");
            }
        }
        static void RemoveCustomXmlParts(OpenXmlPackage doc)
        {
            var package = GetPackage(doc);
            if (package == null) return;

            var parts = package.GetParts()
                .Where(p => p.ContentType == "application/xml" && p.Uri.OriginalString.StartsWith("/customXml/"))
                .ToList();

            foreach (var part in parts)
            {
                try { package.DeletePart(part.Uri); }
                catch (Exception ex)
                {
                    Console.WriteLine($"⚠ Failed to delete XML part: {ex.Message}");
                }
            }
        }

        static void RemoveCustomProperties(OpenXmlPackage doc)
        {
            if (doc is WordprocessingDocument wordDoc && wordDoc.CustomFilePropertiesPart != null)
            {
                wordDoc.DeletePart(wordDoc.CustomFilePropertiesPart);
            }
            else if (doc is SpreadsheetDocument excelDoc && excelDoc.CustomFilePropertiesPart != null)
            {
                excelDoc.DeletePart(excelDoc.CustomFilePropertiesPart);
            }
            else if (doc is PresentationDocument pptDoc && pptDoc.CustomFilePropertiesPart != null)
            {
                pptDoc.DeletePart(pptDoc.CustomFilePropertiesPart);
            }
        }

        static void RemoveExtendedProperties(OpenXmlPackage doc)
        {
            if (doc is WordprocessingDocument wordDoc && wordDoc.ExtendedFilePropertiesPart?.Properties != null)
            {
                var props = wordDoc.ExtendedFilePropertiesPart.Properties;
                props.Company = null;
                props.Manager = null;
                props.Save();
            }
            else if (doc is SpreadsheetDocument excelDoc && excelDoc.ExtendedFilePropertiesPart?.Properties != null)
            {
                var props = excelDoc.ExtendedFilePropertiesPart.Properties;
                props.Company = null;
                props.Manager = null;
                props.Save();
            }
            else if (doc is PresentationDocument pptDoc && pptDoc.ExtendedFilePropertiesPart?.Properties != null)
            {
                var props = pptDoc.ExtendedFilePropertiesPart.Properties;
                props.Company = null;
                props.Manager = null;
                props.Save();
            }
        }

        static void RemovePackageProperties(OpenXmlPackage doc)
        {
            if (doc.PackageProperties != null)
            {
                doc.PackageProperties.Creator = "";
                doc.PackageProperties.LastModifiedBy = "";
            }
        }

        static Package GetPackage(OpenXmlPackage doc)
        {
            var prop = typeof(OpenXmlPackage).GetProperty("Package", BindingFlags.NonPublic | BindingFlags.Instance);
            return prop?.GetValue(doc) as Package;
        }


        public static class PackageHelper
        {
            static Package GetPackage(OpenXmlPackage doc)
            {
                var packageField = typeof(OpenXmlPackage).GetProperty("Package", BindingFlags.NonPublic | BindingFlags.Instance);
                return packageField?.GetValue(doc) as Package;
            }
        }
    }
}
