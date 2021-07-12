using AForge.Imaging.Filters;
using Aspose.OCR;
using DG_Comprot_RpaML.Model;
using Microsoft.Azure.CognitiveServices.Vision.ComputerVision;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;

namespace DG.Comprot.Rpa
{
    class Program
    {
        private static string DriverPath =>
            $@"{Directory.GetCurrentDirectory()}\WebDrivers";

        static void Main(string[] args)
        {
            var comprot = new Comprot("70323117171");
            using (var driver = new ChromeDriver(DriverPath))
            {
                try
                {
                    //var credenciais = new ApiKeyServiceClientCredentials("9671205d08254b6391c0cd1b76ab47e6");

                    //using(var client = new ComputerVisionClient(credenciais) { Endpoint = "https://eastus.api.cognitive.microsoft.com" })
                    //{
                    //    var result = await client.RecognizePrintedTextInStreamAsync(true,);
                    //}

                    TelaInicialComprot(comprot, driver);

                    if (driver.Url.Contains("processo-consulta-lista"))
                    {
                        TelaConsultaLista(comprot, driver);
                    }
                }
                catch 
                {
                    driver.Dispose();
                }

            }
        }

        private static void TelaConsultaLista(Comprot comprot, ChromeDriver driver)
        {
            var totalLabel = driver.FindElementById("campo-total");

            comprot.QuantidadeProcessos = Convert.ToInt32(totalLabel.Text);
        }

        private static void TelaInicialComprot(Comprot comprot, ChromeDriver driver)
        {
            driver.Url = "https://comprot.fazenda.gov.br/comprotegov/site/index.html#ajax/processo-consulta.html";
            IWebElement cpfInput = driver.FindElementById("campo-cpf-cnpj");
            cpfInput.Clear();
            cpfInput.SendKeys($"   {comprot.Cpf}");

            var captchaResolved = false;

            while (!captchaResolved)
            {
                string captchaResult = string.Empty;
                do
                {
                    var captchaImg = driver.FindElementById("img_captcha_serpro_gov_br");
                    captchaImg.Click();

                    Task.Delay(1000).GetAwaiter().GetResult();
                    
                    captchaImg = driver.FindElementById("img_captcha_serpro_gov_br");

                    var src = captchaImg.GetAttribute("src");

                    byte[] data = Convert.FromBase64String(src.Split(',').LastOrDefault());
                    Bitmap bitmap;

                    using (MemoryStream stream = new MemoryStream(data))
                    {
                        bitmap = new Bitmap(stream);
                        //bitmap = ReconhecerImagem(stream);
                    }

                    //var newStream = new MemoryStream();

                    //bitmap.Save(newStream, ImageFormat.Png);
                    
                    captchaResult = ValidarCaptcha(bitmap);

                    //var captchaCleanResult = ValidarCaptchaClean(bitmap);



                    bitmap.Save(@$"E:\MachineLearning\NewCaptchas\result_{captchaResult}_{Guid.NewGuid().ToString()}.png");

                    //using (var stream2 = new MemoryStream())
                    //{
                    //    bitmap.Save(stream2, ImageFormat.Png);

                    //    captchaResult = libVar.RecognizeLine(stream2);
                    //}
                }
                while (true);

                var captchaInput = driver.FindElementById("txtTexto_captcha_serpro_gov_br");

                captchaInput.SendKeys(captchaResult);

                var submit = driver.FindElementsByClassName("btn-primary");

                submit.First().Click();

                var errors = driver.FindElementsByClassName("label-danger");

                if (!errors.Any(error => error.Displayed))
                    captchaResolved = true;
            }
        }

        private static string ValidarCaptcha(Bitmap image)
        {
            List<Bitmap> bitmaps = new List<Bitmap>();
            var widthToIgnore = image.Width / 14;
            var officialWidth = ((image.Width - widthToIgnore) / 6);
            for (int i = 0; i < 6; i++)
            {
                var newImage = new Bitmap(officialWidth, image.Height);

                var widthToJump = widthToIgnore;
                if (i > 0)
                    widthToJump += officialWidth * i;

                using (Graphics g = Graphics.FromImage(newImage))
                {
                    //var toCrop = new Rectangle(0, 0, image.Width, image.Height);
                    g.DrawImageUnscaled(image, new Rectangle(new Point(-widthToJump, 0), new Size(officialWidth, newImage.Height)));
                    //pictureBox.Image = image;
                    //pictureBox1.Image = newImage;
                }

                bitmaps.Add(newImage);
            }

            string final = string.Empty;
            int counter = 1;
            foreach (var bitmap in bitmaps)
            {
                //PredictionScore bestChoice = SimplePrediction(counter, bitmap);
                PredictionScore bestChoice = ComplexPrediction(counter, bitmap);

                if (bestChoice.Score > 2.0F)
                    bitmap.Save(@$"E:\MachineLearning\Comprot\{bestChoice.Prediction}\BestChoice_{bestChoice.Prediction}_{bestChoice.Score}%_{Guid.NewGuid().ToString()}.png");
                else
                    bitmap.Save(@$"E:\MachineLearning\NotClassified\{bestChoice.Prediction}_{bestChoice.Score}%_{Guid.NewGuid().ToString()}.png");

                final += bestChoice.Prediction;

                counter++;
            }

            
            return final;
        }

        private static PredictionScore SimplePrediction(int counter, Bitmap bitmap)
        {
            var bitmap100rgb = CleanImage(bitmap, 100);
            var bitmap150rgb = CleanImage(bitmap, 150);

            ModelOutput predictionResult100rgb = RunPrediction(counter, bitmap100rgb);
            ModelOutput predictionResult150rgb = RunPrediction(counter, bitmap150rgb);
            ModelOutput predictionResult = RunPrediction(counter, bitmap);

            List<PredictionScore> predictions = new List<PredictionScore>();

            var scoreStandard = predictionResult.Score.Max();
            var score100Rgb = predictionResult100rgb.Score.Max();
            var score150Rgb = predictionResult150rgb.Score.Max();

            predictions.Add(new PredictionScore(predictionResult.Prediction, scoreStandard));
            predictions.Add(new PredictionScore(predictionResult100rgb.Prediction, score100Rgb));
            predictions.Add(new PredictionScore(predictionResult150rgb.Prediction, score150Rgb));

            //if (standardScore < 0.45)
            //    bitmap.Save(@$"E:\MachineLearning\NotClassified\{predictionResult.Prediction}_{standardScore}%_{Guid.NewGuid().ToString()}.png");
            //else if (standardScore > 0.75)
            //    bitmap.Save(@$"E:\MachineLearning\Comprot\{predictionResult.Prediction}\generated_2_{predictionResult.Prediction}_{standardScore}%_{Guid.NewGuid().ToString()}.png");
            //else
            //    bitmap.Save(@$"E:\MachineLearning\WellClassified\wellclassified_{predictionResult.Prediction}_{standardScore}%_{Guid.NewGuid().ToString()}.png");

            var bestChoice = predictions.OrderByDescending(prediction => prediction.Score).First();
            return bestChoice;
        }

        private static PredictionScore ComplexPrediction(int counter, Bitmap bitmap)
        {
            var prediction = new FinalPrediction();
            var bitmap100rgb = CleanImage(bitmap, 100);
            var bitmap150rgb = CleanImage(bitmap, 150);

            var bitmapRotate180 = RotateImage(bitmap, RotateFlipType.Rotate180FlipNone);
            var bitmapRotate90 = RotateImage(bitmap, RotateFlipType.Rotate90FlipNone);
            var bitmapRotate270 = RotateImage(bitmap, RotateFlipType.Rotate270FlipNone);

            var bitmap100Rotate180 = RotateImage(bitmap100rgb, RotateFlipType.Rotate180FlipNone);
            var bitmap100Rotate90 = RotateImage(bitmap100rgb, RotateFlipType.Rotate90FlipNone);
            var bitmap100Rotate270 = RotateImage(bitmap100rgb, RotateFlipType.Rotate270FlipNone);

            var bitmap150Rotate180 = RotateImage(bitmap150rgb, RotateFlipType.Rotate180FlipNone);
            var bitmap150Rotate90 = RotateImage(bitmap150rgb, RotateFlipType.Rotate90FlipNone);
            var bitmap150Rotate270 = RotateImage(bitmap150rgb, RotateFlipType.Rotate270FlipNone);

            prediction.AddPrediction(RunPrediction(counter, bitmap100rgb));
            prediction.AddPrediction(RunPrediction(counter, bitmap150rgb));
            prediction.AddPrediction(RunPrediction(counter, bitmap));

            prediction.AddPrediction(RunPrediction(counter, bitmapRotate180), 0.5F);
            prediction.AddPrediction(RunPrediction(counter, bitmapRotate90), 0.5F);
            prediction.AddPrediction(RunPrediction(counter, bitmapRotate270), 0.5F);

            prediction.AddPrediction(RunPrediction(counter, bitmap100Rotate180), 0.5F);
            prediction.AddPrediction(RunPrediction(counter, bitmap100Rotate90), 0.5F);
            prediction.AddPrediction(RunPrediction(counter, bitmap100Rotate270), 0.5F);

            prediction.AddPrediction(RunPrediction(counter, bitmap150Rotate180), 0.5F);
            prediction.AddPrediction(RunPrediction(counter, bitmap150Rotate90), 0.5F);
            prediction.AddPrediction(RunPrediction(counter, bitmap150Rotate270), 0.5F);


            var bestChoice = prediction.BestChoice();

            return bestChoice;
        }

        public class FinalPrediction
        {
            public Dictionary<string, PredictionScore> Predictions { get; set; } = new Dictionary<string, PredictionScore>();

            public FinalPrediction AddPrediction(PredictionScore prediction, Single weight = 1.0F)
            {
                if (!Predictions.ContainsKey(prediction.Prediction))
                    Predictions.Add(prediction.Prediction, prediction);
                else
                {
                    Predictions[prediction.Prediction].Count += 1;
                    Predictions[prediction.Prediction].Scores.Add((prediction.Score, weight)); 
                }

                return this;
            }

            public FinalPrediction AddPrediction(ModelOutput model, float weight = 1.0F) => AddPrediction(new PredictionScore(model, weight), weight);

            public PredictionScore BestChoice()
            {
                Single betterWeight = Single.MinValue;
                PredictionScore bestChoice = Predictions.FirstOrDefault().Value;

                foreach (var prediction in Predictions)
                {
                    var weight = prediction.Value.CalculateWeight();

                    if (weight > betterWeight)
                    {
                        betterWeight = weight;
                        bestChoice = prediction.Value;
                        bestChoice.Score = betterWeight;
                    }
                }

                return bestChoice;
            }
        }

        public class PredictionScore
        {
            public PredictionScore(string prediction, float score, Single weight = 1.0F)
            {
                Prediction = prediction;
                Score = score;
                Scores = new List<(Single score, Single weightScore)>();
                Scores.Add((score, weight));
                Count = 1;
            }

            public PredictionScore(ModelOutput prediction, Single weight = 1.0F)
            {
                Prediction = prediction.Prediction;
                Score = prediction.Score.Max();
                Scores = new List<(Single score, Single weightScore)>();
                Scores.Add((prediction.Score.Max(), weight));
                Count = 1;
            }

            public string Prediction { get; set; }
            public int Count { get; set; }
            public Single Score { get; set; }
            public List<(Single score, Single weightScore)> Scores { get; set; }

            public Single CalculateWeight()
            {
                var averageScore = Scores.Sum(s => s.score * s.weightScore) / Scores.Count;
                return Count * averageScore;
            }
        }

        private static ModelOutput RunPrediction(int counter, Bitmap bitmap)
        {
            var filePath = @$".\{counter.ToString()}.png";
            bitmap.Save(filePath);
            ModelInput sampleData = new ModelInput()
            {
                ImageSource = filePath,
            };

            // Make a single prediction on the sample data and print results
            var predictionResult = ConsumeModel.Predict(sampleData);
            return predictionResult;
        }
               
        static Bitmap RotateImage(Bitmap image, RotateFlipType rotateFlipType)
        {
            var bitmap = image.Clone(new Rectangle(0, 0, image.Width, image.Height), System.Drawing.Imaging.PixelFormat.Format24bppRgb);
            
            bitmap.RotateFlip(rotateFlipType);

            return bitmap;
            
        }

        static Bitmap CleanImage(Bitmap imagem, int minColorRange)
        {
            //Bitmap imagem = new Bitmap(img);
            imagem = imagem.Clone(new Rectangle(0, 0, imagem.Width, imagem.Height), System.Drawing.Imaging.PixelFormat.Format24bppRgb);
            Erosion erosion = new Erosion();
            Dilatation dilatation = new Dilatation();
            Invert inverter = new Invert();
            ColorFiltering cor = new ColorFiltering();
            cor.Blue = new AForge.IntRange(minColorRange, 255);
            cor.Red = new AForge.IntRange(minColorRange, 255);
            cor.Green = new AForge.IntRange(minColorRange, 255);
            Opening open = new Opening();
            BlobsFiltering bc = new BlobsFiltering();
            Closing close = new Closing();
            GaussianSharpen gs = new GaussianSharpen();
            ContrastCorrection cc = new ContrastCorrection();
            bc.MinHeight = 10;
            FiltersSequence seq = new FiltersSequence(gs, inverter, open, inverter, bc, inverter, open, cc, cor, bc, inverter);

            var bitmap = seq.Apply(imagem);

            return bitmap;
        }
    }

    public class Comprot
    {
        public Comprot()
        {
            Processos = new List<ProcessoComprot>();
        }

        public Comprot(string cpf) : this()
        {
            Cpf = cpf;
        }

        public string Cpf { get; set; }
        public string Nome { get; set; }
        public int QuantidadeProcessos { get; set; }
        public List<ProcessoComprot> Processos { get; set; }
        public class ProcessoComprot
        {
            public DateTime DataProtocolo { get; set; }
            public string Processo { get; set; }
        }
    }
}
