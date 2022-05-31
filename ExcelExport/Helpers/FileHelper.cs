namespace ExcelExport.Helpers
{
    using System;
    using System.Globalization;
    using System.IO;
    using System.Text.RegularExpressions;

    public static class FileHelper
    {   
        public static void CheckFile(String path)
        {
            if (String.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                throw new Exception(String.Format("Такого файла \"{0}\", не существует!", path));
            }
        }
        public static String CreateResultFile(string templatePath)
        {            
            string resultFolder = @"C:\xlsx_repository\";
            if (!Directory.Exists(resultFolder))
            {
                Directory.CreateDirectory(resultFolder);
            }
            if (!File.Exists(templatePath))
                throw new Exception($"Не удалось найти шаблон документа \"{templatePath}\"!");
            //Если в пути шаблона (в templateName) присутствуют папки, то при выгрузке, тоже создаём папки
            var index = (templatePath).LastIndexOf("\\", StringComparison.Ordinal);
            if (index > 0)
            {
                var directoryTest = (templatePath).Remove(index, (templatePath).Length - index);
                if (!Directory.Exists(directoryTest))
                {
                    Directory.CreateDirectory(directoryTest);
                }
            }
            var templateFileNameWithoutExtension = Path.GetFileNameWithoutExtension(templatePath);
            var fileExtention = Path.GetExtension(templatePath);
            var suffix = Regex.Replace(DateTime.Now.ToString(CultureInfo.InvariantCulture), @"[^a-z0-9]+", "");
            var resultFilePath = $"{resultFolder}{templateFileNameWithoutExtension}_{suffix}{fileExtention}";
            File.Copy(templatePath, resultFilePath, true);
            return resultFilePath;
        }
    }
}