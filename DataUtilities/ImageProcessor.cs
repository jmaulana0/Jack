using System;
using System.IO;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Png;
using SixLabors.ImageSharp.Processing;

namespace DataUtilities
{
    public class ImageProcessor
    {
        /// <summary>  
        /// Rotates a PNG image represented as BinaryData clockwise by 90 degrees.  
        /// </summary>  
        /// <param name="inputImage">The input PNG image as BinaryData.</param>  
        /// <returns>The rotated image as BinaryData.</returns>  
        public BinaryData RotatePngImageClockwise(BinaryData inputImage)
        {
            if (inputImage == null)
                throw new ArgumentNullException(nameof(inputImage), "Input image cannot be null.");

            // Convert BinaryData to byte array  
            byte[] imageBytes = inputImage.ToArray();

            using (var inputStream = new MemoryStream(imageBytes))
            {
                // Load the image using ImageSharp  
                using (Image image = Image.Load(inputStream))
                {
                    // Rotate the image 90 degrees clockwise  
                    image.Mutate(x => x.Rotate(RotateMode.Rotate90));

                    // Save the rotated image to a new memory stream in PNG format  
                    using (var outputStream = new MemoryStream())
                    {
                        image.SaveAsPng(outputStream);
                        byte[] rotatedBytes = outputStream.ToArray();

                        // Create BinaryData from the byte array  
                        return BinaryData.FromBytes(rotatedBytes);
                    }
                }
            }
        }
    }
}
