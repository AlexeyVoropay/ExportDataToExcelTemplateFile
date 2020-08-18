using System;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;

namespace ExportDataToExcelTemplate
{
    public static class FileHelper
    {   
        public static void CheckFile(String path)
        {
            if (String.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                throw new Exception(String.Format("Такого файла \"{0}\", не существует!", path));
            }
        }
        public static String CreateFile(string templateName)
        {
            string templateFolder = Path.GetFullPath(@"..\..\..\..\ExcelTemplates\Templates\");
            string fileExtention = ".xlsx";
            string resultFolder = @"C:\xlsx_repository\";
            if (!Directory.Exists(resultFolder))
            {
                Directory.CreateDirectory(resultFolder);
            }
            var templateFilePath = String.Format("{0}{1}{2}", templateFolder, templateName, fileExtention);
            var templateFolderPath = String.Format("{0}{1}", resultFolder, templateName);
            if (!File.Exists(String.Format("{0}{1}{2}", templateFolder, templateName, fileExtention)))
            {
                throw new Exception(String.Format("Не удалось найти шаблон документа \n\"{0}{1}{2}\"!", templateFolder, templateName, fileExtention));
            }

            //Если в пути шаблона (в templateName) присутствуют папки, то при выгрузке, тоже создаём папки
            var index = (templateFolderPath).LastIndexOf("\\", System.StringComparison.Ordinal);
            if (index > 0)
            {
                var directoryTest = (templateFolderPath).Remove(index, (templateFolderPath).Length - index);
                if (System.IO.Directory.Exists(directoryTest) == false)
                {
                    System.IO.Directory.CreateDirectory(directoryTest);
                }
            }

            var newFilePath = String.Format("{0}_{1}{2}", templateFolderPath, Regex.Replace((DateTime.Now.ToString(CultureInfo.InvariantCulture)), @"[^a-z0-9]+", ""), fileExtention);
            File.Copy(templateFilePath, newFilePath, true);
            return newFilePath;
        }
    }
}