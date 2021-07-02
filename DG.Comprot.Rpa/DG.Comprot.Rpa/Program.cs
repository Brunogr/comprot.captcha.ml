using AForge.Imaging.Filters;
using Aspose.OCR;
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

                    var src = captchaImg.GetAttribute("src");
                    var libVar = new AsposeOcr("ABCDEFGHIJKLMNOPQRSTUVXYZabcdefghijklmnopqrstuvxyz123456789");

                    byte[] data = Convert.FromBase64String(src.Split(',').LastOrDefault());
                    Bitmap bitmap;

                    using (MemoryStream stream = new MemoryStream(data))
                        bitmap = new Bitmap(stream);

                    bitmap.Save(@$"E:\MachineLearning\Captchas\{Guid.NewGuid().ToString()}.png");

                    using (var stream2 = new MemoryStream())
                    {
                        bitmap.Save(stream2, ImageFormat.Png);

                        captchaResult = libVar.RecognizeLine(stream2);
                    }
                }
                while ((captchaResult.Length != 6) || captchaResult.Contains(' '));

                var captchaInput = driver.FindElementById("txtTexto_captcha_serpro_gov_br");

                captchaInput.SendKeys(captchaResult);

                var submit = driver.FindElementsByClassName("btn-primary");

                submit.First().Click();

                var errors = driver.FindElementsByClassName("label-danger");

                if (!errors.Any(error => error.Displayed))
                    captchaResolved = true;
            }
        }

        static Bitmap ReconhecerImagem(MemoryStream img)
        {
            Bitmap imagem = new Bitmap(img);
            imagem = imagem.Clone(new Rectangle(0, 0, imagem.Width, imagem.Height), System.Drawing.Imaging.PixelFormat.Format24bppRgb);
            Erosion erosion = new Erosion();
            Dilatation dilatation = new Dilatation();
            Invert inverter = new Invert();
            ColorFiltering cor = new ColorFiltering();
            cor.Blue = new AForge.IntRange(200, 255);
            cor.Red = new AForge.IntRange(200, 255);
            cor.Green = new AForge.IntRange(200, 255);
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
