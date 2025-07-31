using System.Drawing;
using System.Drawing.Imaging;
using MailKit.Net.Smtp;
using MimeKit;
using System.Diagnostics;

namespace FileWatcher
{
    public class Worker : BackgroundService // Es una clase que se esta heredando de .NET
    {
        private readonly ILogger<Worker> _logger; // Estamos creando la variable local
        private readonly string folderPath = @"C:\Users\Ian\Documents\Visual Studio 2022\FW_Test"; // Path for folder
        private readonly string emailRecipient = "naiseyer@gmail.com"; // Email that receives the confirmation
        private readonly string emailSender = "naiseyer@gmail.com"; // Email that sends the file ** MUST ADD DI\
        private readonly string emailPassword = "groqirgtjwwwdgko"; // Secure with Gmail 2step Auth
        private readonly string smtpServer = "smtp.gmail.com"; // Verify provider
        private readonly int smtpPort = 587; // PORT VERIFY PORT

        private DateTime lastCheck = DateTime.MinValue;


        public Worker(ILogger<Worker> logger)
        {
            _logger = logger; //variable local siendo instanciada por la variable del DI creando el objeto para ser utilizado a nivel global
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken) // Signature del metodo //Los tres reyes magos con await
        {
            if (_logger.IsEnabled(LogLevel.Information))
            {
                _logger.LogInformation("FileWatcher running."); // feedback user
            }

            while (!stoppingToken.IsCancellationRequested) // WHILE LOOP TO VERIFY EVERY 10 SECONDS FOR NEW FILES
            {
                try
                {
                    var newFiles = Directory.GetFiles(folderPath)
                        .Select(f => new FileInfo(f))
                        .Where(f => f.CreationTime > lastCheck)
                        .ToList();

                    if (newFiles.Any())
                    {
                        string screenshotPath = await TakeFolderScreenshotAsync();

                        await SendEmailWithScreenshot(screenshotPath);

                        File.Delete(screenshotPath); // cleans up path
                    }

                    lastCheck = DateTime.Now;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error in file watcher loop.");
                }

                await Task.Delay(TimeSpan.FromSeconds(10), stoppingToken);

            }
        }


        private async Task<string> TakeFolderScreenshotAsync() // FUNCTION TO TAKE SCREENSHOT
        {
            _logger.LogInformation("Taking Screenshot...");

            Process.Start("explorer.exe", folderPath); // OPENS FILE EXPLORER

            await Task.Delay(3000);

            string fileName = Path.Combine(Path.GetTempPath(), $"screenshot_{Guid.NewGuid()}.png");

            _logger.LogInformation("This will be the file name " + fileName);

            var screenBounds = Screen.PrimaryScreen.Bounds;

            using (Bitmap bmp = new Bitmap(screenBounds.Width, screenBounds.Height))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.CopyFromScreen(0, 0, 0, 0, bmp.Size);
                }
                bmp.Save(fileName, ImageFormat.Png);
            }

            return fileName;
    }

        private async Task SendEmailWithScreenshot(string screenshotPath) // FUNCTION TO SEND EMAIL VIA GMAIL SERVICE
        {
            _logger.LogInformation("Sending email...");

            var message = new MimeMessage();
            message.From.Add(MailboxAddress.Parse(emailSender));
            message.To.Add(MailboxAddress.Parse(emailRecipient));
                message.Subject = $"New file detected in folder at {DateTime.Now}";

            var builder = new BodyBuilder
            {
                TextBody = "A new file has been detected. Screenshot attached.",
            };
            builder.Attachments.Add(screenshotPath);

            message.Body = builder.ToMessageBody();

            using var client = new SmtpClient();

            client.ServerCertificateValidationCallback = (s, c, h, e) => true;

            await client.ConnectAsync(smtpServer, smtpPort, MailKit.Security.SecureSocketOptions.StartTls); // MailKit.Security.SecureSocketOptions.StartTls); <== ensures a secure connection.
            await client.AuthenticateAsync(emailSender, emailPassword); 
            await client.SendAsync(message);
            await client.DisconnectAsync(true);

            _logger.LogInformation("Email sent...");
        }

    }
}
