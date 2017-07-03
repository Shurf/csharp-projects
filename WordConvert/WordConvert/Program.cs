using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Aspose.Words;
using Aspose.Words.Saving;
using System.IO;

namespace WordConvert
{
    public class WordConverter
    {
        
        static void ConvertToHtml(string sourceFilePath, string outputDirectoryName)
        {
            var fileName = Path.GetFileNameWithoutExtension(sourceFilePath);

            Console.WriteLine(fileName);

            var document = new Document(sourceFilePath);
            var options = new HtmlSaveOptions(SaveFormat.Html);
            options.ExportImagesAsBase64 = true;
            document.Save(Path.Combine(outputDirectoryName, fileName + ".html"), options);
        }

        public static int ConvertPath(string inputDirectoryName, string outputDirectoryName)
        {
            var outputDirectory = new DirectoryInfo(outputDirectoryName);
            if (!outputDirectory.Exists)
                outputDirectory.Create();

            int result = 0;

            // convert all DOCXs
            foreach (var file in Directory.GetFiles(inputDirectoryName, "*.docx"))
            {
                ConvertToHtml(file, outputDirectoryName);
                result++;
            }

            // convert all DOCs
            foreach (var file in Directory.GetFiles(inputDirectoryName, "*.doc"))
            {
                ConvertToHtml(file, outputDirectoryName);
                result++;
            }

            return result;
        }
    }
}
