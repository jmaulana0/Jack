using PdfiumViewer;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataUtilities
{
    public class PdfToPngService
    {
        /// <summary>  
        /// Converts each page of the specified PDF into a PNG image represented as BinaryData.  
        /// </summary>  
        /// <param name="pdfFilePath">The file path of the PDF to convert.</param>  
        /// <param name="dpi">Dots per inch for the rendered image quality.</param>  
        /// <returns>A list of BinaryData objects, each representing a PNG image of a PDF page.</returns>  
        public List<BinaryData> ConvertPdfToPng(PdfDocument pdfDocument, int dpi = 300)
        {
            var binaryDataList = new List<BinaryData>();

            int pageCount = pdfDocument.PageCount;
            Console.WriteLine($"PDF loaded. Page count: {pageCount}");

            for (int page = 0; page < pageCount; page++)
            {
                Console.WriteLine($"Rendering page {page + 1}/{pageCount}...");

                // Render the PDF page to an image  
                using (var image = pdfDocument.Render(page, dpi, dpi, PdfRenderFlags.Annotations))
                {
                    using (var ms = new MemoryStream())
                    {
                        // Save the image to the memory stream in PNG format  
                        image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);

                        // Convert the memory stream to a byte array  
                        byte[] imageBytes = ms.ToArray();

                        // Create BinaryData from the byte array  
                        BinaryData binaryData = BinaryData.FromBytes(imageBytes);

                        // Add to the list  
                        binaryDataList.Add(binaryData);
                    }
                }
            }

            Console.WriteLine("PDF conversion completed.");
            return binaryDataList;
        }
    }
}
