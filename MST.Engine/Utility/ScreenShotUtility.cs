using ImageMagick;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
using MST.Engine.Model.Initialization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing.Imaging;
using MST.Engine.Model.LineAnalyzer;
using NLog;

namespace MST.Engine.Utility
{
    class ScreenShotUtility
    {
        public static string finalImageFile;
        private static Logger logger = LogManager.GetCurrentClassLogger();
        public static string CompareScreen()
        {
            try
            {
                logger.Debug("------|Start To Compare Screen");
                finalImageFile = string.Empty;
                string templateName = Template.templateName;

                string baselinePath = InitializeParamter.ImageBaselinePath;
                string imagePath = baselinePath.Replace("Baseline", "");
                string actualPath = imagePath + @"Compare\Actual\";
                string diffPath = imagePath + @"Compare\Diff\";
                CreateDirectory(actualPath);
                CreateDirectory(diffPath);
                string date = DateTime.Now.ToString("yyyyMMdd");
                string todayActualPath = Path.Combine(actualPath, date);
                string todayDiffPath = Path.Combine(diffPath, date);
                CreateDirectory(todayActualPath);
                CreateDirectory(todayDiffPath);

                //Define all of the image 
                string baselineImageFullName = Path.Combine(baselinePath, templateName + ".png");
                string actualImageFullName = Path.Combine(todayActualPath, templateName + ".png");
                string diffImageFullName = Path.Combine(todayDiffPath, templateName + "_diff.png");
                string firstCombineImageFullName = Path.Combine(todayDiffPath, templateName + "_diff_Combine1.png");  //baseline + actual
                string secondCombineImageFullName = Path.Combine(todayDiffPath, templateName + "_diff_Combine2.png");  // first + diff            
                string finalDiffImageFullName = Path.Combine(todayDiffPath, templateName + "_diff_final.png"); // second + Watermark

                /*==========1. Get Actual =======================*/
                var driver = Initialize.Driver;
                Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
                ss.SaveAsFile(actualImageFullName, ImageFormat.Png);
                /*----------1 End--------------------------------*/

                /*==========2. Get Diff =======================*/
                //Temp Solution To Get baseline
                if (!File.Exists(baselineImageFullName))
                {
                    baselineImageFullName = actualImageFullName;
                }
                using (MagickImage mi_actual = new MagickImage(actualImageFullName))
                using (MagickImage mi_baseline = new MagickImage(baselineImageFullName))
                using (MagickImage mi_diff = new MagickImage())
                {
                    double distortion = mi_actual.Compare(mi_baseline, ErrorMetric.Absolute, mi_diff);
                    if (distortion != 0)
                    {
                        mi_diff.Write(diffImageFullName);
                        /*----------2 End--------------------------------*/
                        /*==========3. first Combine:baseline + actual===*/
                        using (MagickImageCollection collection = new MagickImageCollection())
                        {
                            collection.Add(new MagickImage(baselineImageFullName));
                            collection.Add(new MagickImage(actualImageFullName));
                            //MagickGeometry geo = new MagickGeometry("80x80%");
                            // Create a mosaic from both images
                            using (MagickImage result = collection.AppendHorizontally())
                            {
                                // Save the result
                                //result.Resize(geo);
                                //result.Scale(new Percentage(50));
                                //result.BorderColor = new MagickColor(System.Drawing.Color.Red);                           
                                result.Write(firstCombineImageFullName);
                            }
                        }
                        /*----------3 End--------------------------------*/
                        /*==========4. second Combine:first + diff ======*/
                        using (MagickImageCollection collection = new MagickImageCollection())
                        {
                            collection.Add(new MagickImage(firstCombineImageFullName));
                            collection.Add(new MagickImage(diffImageFullName));

                            //MagickGeometry geo = new MagickGeometry("150x150%");//wightxhight
                            using (MagickImage result = collection.AppendVertically())
                            {
                                //result.Resize(geo);
                                // Save the result                           
                                result.Write(secondCombineImageFullName);
                            }
                        }
                        /*----------4 End--------------------------------*/
                        /*==========5. final Combine:second + Watermark =*/
                        var imageBackgroundColor = new MagickColor("White");
                        var text = "Left: Baseline\nRight: Actual\nBottom: Different";
                        using (MagickImage image = new MagickImage(secondCombineImageFullName))
                        {
                            // Read the watermark that will be put on top of the image
                            using (MagickImage watermark = new MagickImage(imageBackgroundColor, 1000, 600))
                            {
                                watermark.Read("label:" + text);
                                // Draw the watermark in the bottom right corner
                                image.Composite(watermark, Gravity.Southeast, CompositeOperator.Over);
                                // Optionally make the watermark more transparent
                                watermark.Evaluate(Channels.Alpha, EvaluateOperator.Divide, 4);
                                // Or draw the watermark at a specific location
                                //image.Composite(watermark, 200, 50, CompositeOperator.Over);
                            }
                            // Save the result
                            image.Write(finalDiffImageFullName);
                        }
                        /*----------5 End--------------------------------*/
                        /*==========6. delete middle diff image==========*/
                        File.Delete(diffImageFullName);
                        File.Delete(firstCombineImageFullName);
                        File.Delete(secondCombineImageFullName);
                        // Rename final.png                   
                        string time = DateTime.Now.ToString("HHmmss");
                        string renameFinalImageFullName = finalDiffImageFullName.Replace("diff_final", time);
                        File.Move(finalDiffImageFullName, renameFinalImageFullName);
                        finalImageFile = renameFinalImageFullName;
                        /*----------6 End--------------------------------*/
                        return "Failure";
                    }
                    else
                    {
                        return "Pass";
                    }
                }
            }
            catch (Exception e)
            {
                logger.Fatal("Fail To Compare Screen");
                logger.Fatal(e.Message);
                return "Failure";
            }
        }

        private static void CreateDirectory(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }
    }
}
