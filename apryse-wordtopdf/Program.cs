// Default namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// Majority of Apryse SDK can be used with these namespaces
using pdftron;
using pdftron.Common;
using pdftron.SDF;
using pdftron.PDF;
using pdftron.FDF;
using System.IO;

namespace apryse_wordtopdf
{
    class Program
    {
        // Required for AnyCPU implementation.
        private static PDFNetLoader loader = PDFNetLoader.Instance();
        static String input_path = "../../../TestFiles/";
        static String output_path = "../../../TestFiles/";

        static void BulkConvert(List<string> input_filenames)
        {
            foreach (string input_filename in input_filenames)
            {
                // Start with a PDFDoc (the conversion destination)
                using (PDFDoc pdfdoc = new PDFDoc())
                {
                    // Perform the conversion with no optional parameters
                    pdftron.PDF.Convert.OfficeToPDF(pdfdoc, input_path + input_filename, null);

                    // Generate the output filename based on the input filename
                    string output_filename = "Apryse_" + Path.GetFileNameWithoutExtension(input_filename) + ".pdf";

                    // Save the result
                    pdfdoc.Save(output_path + output_filename, SDFDoc.SaveOptions.e_linearized);

                    // Print the saved output filename
                    Console.WriteLine("Saved " + output_filename);
                }
            }
        }

        static void Main(string[] args)
        {
            var startTime = DateTime.Now;// Record the start time
            // Initialize PDFNet before using any Apryse related
            // classes and methods (some exceptions can be found in API)
            PDFNet.Initialize("demo:1684117218449:7daeec5f0300000000b410c2b2824e02d7e6f8d95281e6b48e2ea17c44");

            // Using PDFNet related classes and methods, must catch or throw PDFNetException
            try
            {
                List<string> filesToConvert = new List<string>
                    {
                        "15-MB-docx-file-download.docx",
                        "Airline_UseCases_Apryse_Features_latest.docx",
                        "invoice_template.docx"
                    };

                BulkConvert(filesToConvert);

                //using (PDFDoc doc = new PDFDoc())
                //{
                //    //doc.InitSecurityHandler();

                //    //// An example of creating a new page and adding it to
                //    //// doc's sequence of pages
                //    //Page newPg = doc.PageCreate();
                //    //doc.PagePushBack(newPg);

                //    //// Save as a linearized file which is most popular 
                //    //// and effective format for quick PDF Viewing.
                //    //doc.Save("linearized_output.pdf", SDFDoc.SaveOptions.e_linearized);

                //    //System.Console.WriteLine("Done. Results saved in linearized_output.pdf");

                //    // perform the conversion with no optional parameters
                //    pdftron.PDF.Convert.OfficeToPDF(doc, input_path + input_filename, null);

                //    // save the result
                //    doc.Save(output_path + output_filename, SDFDoc.SaveOptions.e_linearized);

                //    // And we're done!
                //    Console.WriteLine("Saved " + output_filename);
                //}
            }
            catch (PDFNetException e)
            {
                System.Console.WriteLine(e);
            }

            PDFNet.Terminate();
            var endTime = DateTime.Now; // Record the end time
            var conversionTime = endTime - startTime; // Calculate the conversion time
            Console.WriteLine("Done in "+ conversionTime);
            Console.ReadKey();
        }
    }
}