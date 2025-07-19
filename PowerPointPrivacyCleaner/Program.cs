using DocumentFormat.OpenXml.Packaging;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace PPTMetadataCleaner
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Enter the folder path to scan for PPTX files:");
            string folderPath = Console.ReadLine();

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine("Invalid folder path.");
                return;
            }

            string[] pptFiles = Directory.GetFiles(folderPath, "*.pptx", SearchOption.AllDirectories);
            Console.WriteLine($"Found {pptFiles.Length} files.");

            var pptApp = new Microsoft.Office.Interop.PowerPoint.Application();
            pptApp.Visible = MsoTriState.msoCTrue;

            foreach (string file in pptFiles)
            {
                Console.WriteLine($"\nProcessing: {file}");
                try
                {
                    CleanWithInterop(pptApp, file);
                    CleanWithOpenXml(file);
                    Console.WriteLine("✅ Cleaned successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error: {ex.Message}");
                }
            }

            pptApp.Quit();
            Marshal.ReleaseComObject(pptApp);

            Console.WriteLine("\nAll files processed. Press any key to exit.");
            Console.ReadKey();
        }

        static void CleanWithInterop(Application pptApp, string filePath)
        {
            Presentation presentation = null;
            try
            {
                presentation = pptApp.Presentations.Open(filePath, WithWindow: MsoTriState.msoFalse);

                // Enable "Remove personal information on save"
                presentation.RemovePersonalInformation = MsoTriState.msoTrue;

                // Clear built-in properties
                dynamic props = presentation.BuiltInDocumentProperties;
                ClearProperty(props, "Author");
                ClearProperty(props, "Last Author");
                ClearProperty(props, "Manager");
                ClearProperty(props, "Company");
                ClearProperty(props, "Comments");

                presentation.Save();
            }
            finally
            {
                if (presentation != null)
                {
                    presentation.Close();
                    Marshal.ReleaseComObject(presentation);
                }
            }
        }

        static void ClearProperty(dynamic props, string propName)
        {
            try
            {
                var prop = props[propName];
                if (prop != null && prop.Value != null)
                {
                    prop.Value = "";
                }
            }
            catch
            {
                // Ignore missing properties
            }
        }

        static void CleanWithOpenXml(string filePath)
        {
            using (PresentationDocument ppt = PresentationDocument.Open(filePath, true))
            {
                RemoveCustomXmlParts(ppt);

                // Remove Custom Properties
                if (ppt.CustomFilePropertiesPart != null)
                {
                    ppt.DeletePart(ppt.CustomFilePropertiesPart);
                }

                // Remove Extended Properties
                if (ppt.ExtendedFilePropertiesPart != null)
                {
                    var props = ppt.ExtendedFilePropertiesPart.Properties;
                    if (props != null)
                    {
                        props.Company = null;
                        props.Manager = null;
                        props.Save();
                    }
                }

                // Remove Core Properties
                if (ppt.PackageProperties != null)
                {
                    ppt.PackageProperties.Creator = "";
                    ppt.PackageProperties.LastModifiedBy = "";
                }
            }
        }

        static void RemoveCustomXmlParts(PresentationDocument ppt)
        {
            var package = PackageHelper.GetPackage(ppt);
            if (package == null) return;

            var customXmlParts = package.GetParts()
                .Where(p => p.ContentType == "application/xml" && p.Uri.OriginalString.StartsWith("/customXml/"))
                .ToList();

            foreach (var part in customXmlParts)
            {
                try
                {
                    package.DeletePart(part.Uri);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"⚠ Failed to delete custom XML part {part.Uri}: {ex.Message}");
                }
            }
        }




        public static class PackageHelper
        {
            public static Package GetPackage(PresentationDocument doc)
            {
                // Use reflection to access the internal _package field
                var packageField = typeof(OpenXmlPackage).GetProperty("Package", BindingFlags.NonPublic | BindingFlags.Instance);
                return packageField?.GetValue(doc) as Package;
            }
        }
    }
}
