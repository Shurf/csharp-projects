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
    class Program
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

        static void Main(string[] args)
        {
            var inputDirectoryName = "g:\\projects\\books\\pisma_mahatm";
            var outputDirectoryName = inputDirectoryName + "_converted";

            var outputDirectory = new DirectoryInfo(outputDirectoryName);
            if (!outputDirectory.Exists)
                outputDirectory.Create();

            foreach (var file in Directory.GetFiles(inputDirectoryName, "*.docx"))
            {
                ConvertToHtml(file, outputDirectoryName);
            }

            /*foreach(var book_dir in Directory.GetDirectories("g:\\projects\\books\\roerich"))
            {
                var inputDirectoryName = book_dir;
                var outputDirectoryName = inputDirectoryName + "_converted";
                var outputDirectory = new DirectoryInfo(outputDirectoryName);
                if (!outputDirectory.Exists)
                    outputDirectory.Create();
                foreach (var file in Directory.GetFiles(inputDirectoryName, "*.docx"))
                {
                    ConvertToHtml(file, outputDirectoryName);
                }
            }*/

            
        }
    }
}
