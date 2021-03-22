using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ConvertPDF.Controllers
{

    public class ConvertPostBody
    {
        public string Base64 { get; set; }
    }
    [Route("api/[controller]")]
    [ApiController]
    public class ConvertPDFController : ControllerBase
    {
        public static string rawFolder = @"C:\Users\dev\Desktop\tmp\raw";
        public static string convertedFolder = @"C:\Users\dev\Desktop\tmp\converted";
        public static string execWordPath = "C:/\"Program Files\"/LibreOffice/program/swriter.exe";
        public static string execExcelPath = "C:/\"Program Files\"/LibreOffice/program/scalc.exe";

        [HttpPost]
        public async Task<ActionResult> PostAsync([FromBody] ConvertPostBody data)
        {
            string converted = await ConvertPDFAsync(data.Base64);
            return Ok(converted);
        }

        public static async Task<string> ConvertPDFAsync(string base64)
        {
            Directory.CreateDirectory(rawFolder);
            Directory.CreateDirectory(convertedFolder);

            string now = DateTime.Now.ToString("FFFFFFF");
            string fileName = now;
            string rawFilePath = Path.Join(rawFolder, fileName);
            string convertedFilePath = Path.Join(convertedFolder, fileName + ".pdf");

            await System.IO.File.WriteAllBytesAsync(rawFilePath, Convert.FromBase64String(base64));

            byte[] convertedBytes;
            try
            {
                await RunProcessAsync("/C " + execWordPath + " --headless --convert-to pdf " + rawFilePath + " --outdir " + convertedFolder);

                convertedBytes = await System.IO.File.ReadAllBytesAsync(convertedFilePath);
            }
            catch (Exception)
            {
                Console.WriteLine("fail");
                await RunProcessAsync("/C del /f " + rawFilePath);
                return null;
            }

            // delete tmp
            await RunProcessAsync("/C del /f " + convertedFilePath);
            await RunProcessAsync("/C del /f " + rawFilePath);

            Console.WriteLine("done");
            return Convert.ToBase64String(convertedBytes);
        }

        static Task<int> RunProcessAsync(string command)
        {
            var tcs = new TaskCompletionSource<int>();

            var process = new Process
            {
                StartInfo = { WindowStyle = ProcessWindowStyle.Hidden, FileName = "cmd.exe", Arguments = command },
                EnableRaisingEvents = true
            };

            process.Exited += (sender, args) =>
            {
                tcs.SetResult(process.ExitCode);
                process.Dispose();
            };

            process.Start();

            return tcs.Task;
        }
    }
}
