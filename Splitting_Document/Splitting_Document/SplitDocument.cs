
using System;
using Tesseract;

namespace Splitting_Document
{
    class SplitDocument
    {
        /// <summary>
        /// Get OCR extraction from image file
        /// </summary>
        /// <param name="imageFilePath">Image file path</param>
        /// <param name="extrationConfidence">Extraction confidence level</param>
        /// <returns>Text extracted</returns>
        public string GetExtractedText(string imageFilePath, ref float extrationConfidence)
        {
            string retrunValue = "";

            try
            {
                using (var engine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default))
                {
                    using (var img = Pix.LoadFromFile(imageFilePath))
                    {
                        using (var page = engine.Process(img))
                        {
                            var text = page.GetText();
                            extrationConfidence = page.GetMeanConfidence();
                            retrunValue = retrunValue + String.Format("Mean confidence: {0}", extrationConfidence);

                            retrunValue = retrunValue + String.Format("Text (GetText): \r\n{0}", text);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                //retrunValue = retrunValue + "Unexpected Error: " + e.Message;
                //retrunValue = retrunValue + " Details: ";
                //retrunValue = retrunValue + e.ToString();
            }

            return retrunValue;
        }
    }
}
