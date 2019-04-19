using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace Splitting_Document
{
    class Program
    {

        static void Main(string[] args)
        {
            List<string> listExpressionToFind = new List<string>();
            List<string> listExpressionToExclude = new List<string>();
            List<string> listTargetFolder = new List<string>();

            bool isClassified = true;
            int indexSplittingRules = 0;

            string multiplePagesFolder = args[0]; //@"C:\Users\Romain.Alexandre\Documents\Automation Anywhere Files\Automation Anywhere\My Docs\Split Document\Test Invoice\Multiple Pages";
            string UnclassifiedFolder = args[1];//@"C:\Users\Romain.Alexandre\Documents\Automation Anywhere Files\Automation Anywhere\My Docs\Split Document\Test Invoice\Unclassified";
            string splittingRulesCSV = args[2]; //@"C:\Users\Romain.Alexandre\Documents\Automation Anywhere Files\Automation Anywhere\My Docs\Split Document\Test Invoice\Splitting_Rules.csv";

            string OCRExtraction = "";

            int numberOfPages = 0;
            int pageIndex = 1;

            float extractionConfidence = 0;
            int indexFlip = 0;

            SplitDocument objSplitDocument = new SplitDocument();

            Console.WriteLine(string.Format("Classification in progress..."));

            if (File.Exists(multiplePagesFolder + "\\IMG\\SplittingComplete.txt"))
                File.Delete(multiplePagesFolder + "\\IMG\\SplittingComplete.txt");

            using (var reader = new StreamReader(splittingRulesCSV))
            {

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');

                    listExpressionToFind.Add(values[0]);
                    listExpressionToExclude.Add(values[1]);
                    listTargetFolder.Add(values[2]);
                }
            }

            listExpressionToFind.RemoveAt(0);
            listExpressionToExclude.RemoveAt(0);
            listTargetFolder.RemoveAt(0);

            DirectoryInfo directoryInfo = new DirectoryInfo(multiplePagesFolder + "\\IMG");
            numberOfPages = directoryInfo.GetFiles().Length;

            foreach (var file in directoryInfo.GetFiles("*.png"))
            {
                isClassified = true;
                indexSplittingRules = 0;
                extractionConfidence = 0.1F;
                indexFlip = 0;

                Console.WriteLine();
                Console.WriteLine(string.Format("Process page {0} / {1}", pageIndex, numberOfPages));

                while (extractionConfidence < 0.70F && extractionConfidence > 0 && indexFlip < 4)
                {
                    OCRExtraction = objSplitDocument.GetExtractedText(multiplePagesFolder + "\\IMG\\" + file.Name, ref extractionConfidence);

                    if (extractionConfidence < 0.70F && extractionConfidence>0)
                    {
                        FlipImage(multiplePagesFolder + "\\IMG\\" + file.Name);
                        indexFlip++;
                    }
                }

                if (extractionConfidence >= 0.70F)
                {

                    foreach (string expressionToFind in listExpressionToFind)
                    {
                        var listExpression1 = expressionToFind.Split('|');
                        foreach (string expression in listExpression1)
                        {
                            if (expression.Trim() != "" && !OCRExtraction.ToLower().Contains(expression.ToLower()))
                            {
                                isClassified = false;
                                break;
                            }
                            else
                                isClassified = true;
                        }

                        if (isClassified)
                        {
                            var listExpression2 = listExpressionToExclude[indexSplittingRules].Split('|');
                            foreach (string expression in listExpression2)
                            {
                                if (expression.Trim() != "" && OCRExtraction.ToLower().Contains(expression.ToLower()))
                                {
                                    isClassified = false;
                                    break;
                                }
                            }

                        }

                        if (isClassified)
                            break;

                        indexSplittingRules++;

                    }

                    if (isClassified)
                    {
                        Console.WriteLine(string.Format("Page classified!"));
                        if (File.Exists(listTargetFolder[indexSplittingRules] + "\\" + file.Name.Replace(".png", ".pdf")))
                            File.Delete(listTargetFolder[indexSplittingRules] + "\\" + file.Name.Replace(".png", ".pdf"));
                        Directory.Move(file.FullName.Replace("\\IMG", "").Replace(".png", ".pdf"), listTargetFolder[indexSplittingRules] + "\\" + (file.Name).Replace(".png", ".pdf"));
                        file.Delete();
                    }
                    else
                    {
                        Console.WriteLine(string.Format("Page not classified!"));
                        if (File.Exists(UnclassifiedFolder + "\\" + file.Name.Replace(".png", ".pdf")))
                            File.Delete(UnclassifiedFolder + "\\" + file.Name.Replace(".png", ".pdf"));
                        Directory.Move(file.FullName.Replace("\\IMG", "").Replace(".png", ".pdf"), UnclassifiedFolder + "\\" + file.Name.Replace(".png", ".pdf"));
                        file.Delete();
                    }

                }
                else
                {
                    Console.WriteLine(string.Format("Page not classified, extraction level too low!"));
                    if (File.Exists(UnclassifiedFolder + "\\" + file.Name.Replace(".png", ".pdf")))
                        File.Delete(UnclassifiedFolder + "\\" + file.Name.Replace(".png", ".pdf"));
                    Directory.Move(file.FullName.Replace("\\IMG", "").Replace(".png", ".pdf"), UnclassifiedFolder + "\\" + file.Name.Replace(".png", ".pdf"));
                    file.Delete();
                }

                pageIndex++;

            }

            if (File.Exists(multiplePagesFolder + "\\IMG\\SplittingComplete.txt"))
                File.Delete(multiplePagesFolder + "\\IMG\\SplittingComplete.txt");
            File.Create(multiplePagesFolder + "\\IMG\\SplittingComplete.txt");

            Console.WriteLine(string.Format("Classification complete!"));
            System.Threading.Thread.Sleep(1000);
        }


        /// <summary>
        /// Rotate image 90 degree clockwise
        /// </summary>
        /// <param name="imgFilePath">Image file to rotate</param>
        static void FlipImage(string imgFilePath)
        {
            System.Drawing.Image img = System.Drawing.Image.FromFile(imgFilePath);
            img.RotateFlip(RotateFlipType.Rotate90FlipNone);
            System.IO.File.Delete(imgFilePath);
            img.Save(imgFilePath);
        }
    }
}
