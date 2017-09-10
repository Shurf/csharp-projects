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
        
        static void ConvertToHtml(string sourceFilePath, string outputDirectoryName, string outputFileName)
        {
            var fileName = Path.GetFileNameWithoutExtension(sourceFilePath);

            Console.WriteLine(fileName);

            var document = new Document(sourceFilePath);
            var options = new HtmlSaveOptions(SaveFormat.Html);
            options.ExportImagesAsBase64 = true;
            document.Save(Path.Combine(outputDirectoryName, outputFileName + ".html"), options);
        }

        public static int ConvertPath(string inputDirectoryName, string outputDirectoryName, bool rename=false)
        {
            var outputDirectory = new DirectoryInfo(outputDirectoryName);
            if (!outputDirectory.Exists)
                outputDirectory.Create();

            int result = 0;
            int count = 1;

            var files = new SortedSet<string>();

            // gather all DOCXs
            foreach (var file in Directory.GetFiles(inputDirectoryName, "*.docx").OrderBy(x => x))
            {
                if (!files.Add(file))
                {
                    Console.WriteLine("Warning! " + file + " will not be converted!");
                }
            }

            // gather all DOCs
            foreach (var file in Directory.GetFiles(inputDirectoryName, "*.doc").OrderBy(x => x))
            {
                if (!files.Add(file))
                {
                    Console.WriteLine("Warning! " + file + " will not be converted!");
                }
            }

            // convert
            foreach (var file in files)
            {
                var newName = rename ? count.ToString("D4") : file;
                ConvertToHtml(file, outputDirectoryName, newName);
                count++;
                result++;
            }

            return result;
        }
    }
}
