#region Using Statements
using System.Drawing.Imaging;
using MailKit.Net.Smtp;
using MimeKit;
using System.Diagnostics;
using SHDocVw;
using Microsoft.Data.SqlClient;
#endregion

namespace FileWatcher
{
    public class Worker : BackgroundService // Es una clase que se esta heredando de .NET
    {
        #region Variables
        private readonly ILogger<Worker> _logger; // Estamos creando la variable local
        private readonly string folderPath = @"C:\Users\Ian\Documents\Visual Studio 2022\FW_Test"; // Path for folder

        private readonly string emailRecipient = "naiseyer@gmail.com"; // Email that receives the confirmation
        private readonly string emailSender = "naiseyer@gmail.com"; // Email that sends the file ** MUST ADD DI\
        private readonly string emailPassword = "groqirgtjwwwdgko"; // Secure with Gmail 2step Auth
        private readonly string smtpServer = "smtp.gmail.com"; // Verify provider
        private readonly int smtpPort = 587; // PORT VERIFY PORT

        private readonly string _connectionString;
        private DateTime lastCheck = DateTime.MinValue;
        #endregion

        #region Constructor

        public Worker(ILogger<Worker> logger, IConfiguration config)
        {
            _logger = logger; //variable local siendo instanciada por la variable del DI creando el objeto para ser utilizado a nivel global
            _connectionString = config.GetConnectionString("FileWatcherDB");
        }

        #endregion

        #region Protected Overridable Methods
        
        protected override async Task ExecuteAsync(CancellationToken stoppingToken) // Signature del metodo //Los tres reyes magos con await
        {
            _logger.LogInformation("FileWatcher started at: {time}", DateTimeOffset.Now); // feedback user


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
                        _logger.LogInformation("{count} new file(s) detected.", newFiles.Count);

                        await OpenOrReuseExplorerAsync(folderPath);


                        string screenshotPath = await TakeFolderScreenshotAsync();
                        _logger.LogInformation("Screenshot saved to: {path}", screenshotPath);

                        await SendEmailWithScreenshot(screenshotPath);

                        await LogEmailSentAsync(screenshotPath, emailRecipient, newFiles.Count);

                        File.Delete(screenshotPath); // cleans up path
                        _logger.LogInformation("Temporary screenshot deleted.");
                    }

                    lastCheck = DateTime.Now;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error in file watcher loop.");
                }

                await Task.Delay(TimeSpan.FromSeconds(10), stoppingToken);

            }
            _logger.LogInformation("FileWatcher stopping at: {time}", DateTimeOffset.Now);
        }

        #endregion

        #region Private Methods

        private async Task OpenOrReuseExplorerAsync(string folderPath)
        {
            folderPath = Path.GetFullPath(folderPath).TrimEnd('\\');
            ShellWindows shellWindows = new ShellWindows();

            foreach (InternetExplorer window in shellWindows)
            {
                string location = window.LocationURL.Replace("file:///", "").Replace("/", "\\");
                location = Uri.UnescapeDataString(location).TrimEnd('\\');

                if (string.Equals(location, folderPath, StringComparison.OrdinalIgnoreCase))
                {
                    _logger.LogInformation("Explorer window already open for {folder}, bringing to front.", folderPath);
                    window.Visible = true;
                    window.FullScreen = false;
                    window.Left = 0;
                    window.Top = 0;
                    return;
                }
            }

            _logger.LogInformation("Opening new Explorer window for {folder}", folderPath);
            Process.Start("explorer.exe", folderPath);
            await Task.Delay(2000);
        }

        private async Task<string> TakeFolderScreenshotAsync()
        {
            _logger.LogInformation("Taking screenshot...");
            string fileName = Path.Combine(Path.GetTempPath(), $"screenshot_{Guid.NewGuid()}.png");

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

            _logger.LogInformation("Email sent to {recipient}", emailRecipient);
        }
        
        private async Task LogEmailSentAsync(string fileName, string recipient, int filesDetected)
        {
            try
            {
                using var conn = new SqlConnection(_connectionString);
                await conn.OpenAsync();

                string sql = @"
                    INSERT INTO EmailLog (FileName, RecipientEmail, FilesDetectedCount, FolderPath)
                    VALUES (@FileName, @RecipientEmail, @FilesDetectedCount, @FolderPath);";

                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@FileName", Path.GetFileName(fileName));
                cmd.Parameters.AddWithValue("@RecipientEmail", recipient);
                cmd.Parameters.AddWithValue("@FilesDetectedCount", filesDetected);
                cmd.Parameters.AddWithValue("@FolderPath", folderPath);

                await cmd.ExecuteNonQueryAsync();

                _logger.LogInformation("Email log saved to database for file: {file}", fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving email log to database.");
            }

        }

        #endregion
    }
}
