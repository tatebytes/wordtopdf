// Load a Word file from the local drive.
using Aspose.Words;

namespace aspose_wordtopdf
{
    class Program
    {
        static String input_path = "../../../../TestFiles/";

        static void BulkConvert(List<string> input_filenames)
        {
            foreach (string input_filename in input_filenames)
            {
                Document doc = new Document(input_path + input_filename);
                string output_filename = "Aspose_" + Path.GetFileNameWithoutExtension(input_filename) + ".pdf";
                // Save it to HTML format.
                doc.Save(input_path + output_filename);
                Console.WriteLine("Saved " + output_filename);
            }
        }

        static void Main(string[] args)
        {
            var startTime = DateTime.Now;// Record the start time
            try
            {
                List<string> filesToConvert = new List<string>
                    {
                        "15-MB-docx-file-download.docx",
                        "Airline_UseCases_Apryse_Features_latest.docx",
                        "invoice_template.docx"
                    };

                BulkConvert(filesToConvert);
            }
            catch (Exception e)
            {
                System.Console.WriteLine(e);
            }

            var endTime = DateTime.Now; // Record the end time
            var conversionTime = endTime - startTime; // Calculate the conversion time
            Console.WriteLine("Done in " + conversionTime);
            Console.ReadKey();
        }
    }
}