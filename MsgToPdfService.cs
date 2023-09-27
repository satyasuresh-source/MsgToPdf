namespace Integrity.Services.MsgToPdf
{
    using System;
    using System.IO;
    using System.Threading;
    using MsgReader;
    using System.Threading.Tasks;
    using System.Diagnostics;
    using Microsoft.Extensions.Hosting;
    using Microsoft.Extensions.Logging;
    using Azure.Storage.Files.Shares;
    using Azure.Storage.Files.Shares.Models;
    using Azure;
    using SendGrid.Helpers.Mail;
    using System.Linq;
    using System.Text.RegularExpressions;
    using SendGrid;
    using Microsoft.Extensions.Configuration;
    using Microsoft.EntityFrameworkCore;
    using Integrity.Services.MsgToPdf.Models;
    public class MsgToPdfService : IHostedService
    {
        /// <summary>
        /// Represents a set of key/value application configuration properties.
        /// </summary>
        private readonly IConfiguration configuration;

        /// <summary>
        /// A generic interface for logging where the category name is derived from the specified.
        /// </summary>
        private readonly ILogger<MsgToPdfService> logger;

        /// <summary>
        /// Represents azure storage connection string. 
        /// </summary>
        private string connectionString { get; set; }

        /// <summary>
        /// Represents conversion path.
        /// </summary>
        private string conversionPath { get; set; }

        /// <summary>
        /// Represents backup fax delivery receipts.
        /// </summary>
        private string backupFaxDeliveryReceiptsPath { get; set; }

        /// <summary>
        /// Represents fax delivery receipts storage path.
        /// </summary>
        private string faxDeliveryReceipts { get; set; }

        /// <summary>
        /// Represents file share name.
        /// </summary>
        private string fileshareName { get; set; }

        /// <summary>
        /// Represents finished receipts.
        /// </summary>
        private string finishedReceipts { get; set; }        

        /// <summary>
        /// Represents share service client.
        /// </summary>
        private ShareServiceClient shareServiceClient { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public MsgToPdfService()
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="configuration">Represents a set of key/value application configuration properties.</param>
        /// <param name="logger">A generic interface for logging where the category name is derived
        /// from the specified.</param>
        /// <exception cref="ArgumentNullException">The ArgumentException is thrown when an argument
        /// is null when it shouldn't be.</exception>
        public MsgToPdfService(IConfiguration configuration, ILogger<MsgToPdfService> logger)
        {
            this.configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
            this.logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }


        /// <summary>
        /// Method to convert the .msg file to pdf.
        /// </summary>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Represents an asynchronous operation.</returns>
        protected async Task RunAsync(CancellationToken cancellationToken)
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                try
                {
                    connectionString = configuration["AzureFileStorage:ConnectionString"];
                    fileshareName = configuration["AzureFileStorage:FileShareName"];
                    faxDeliveryReceipts = configuration["AzureFileStorage:FaxDeliveryReceiptPath"];
                    conversionPath = configuration["AzureFileStorage:ConversionPath"];
                    finishedReceipts = configuration["AzureFileStorage:FinishedReceipts"];
                    backupFaxDeliveryReceiptsPath = configuration["AzureFileStorage:BackupFaxDeliveryReceipts"];

                    shareServiceClient = new ShareServiceClient(connectionString);
                    var shareClient = shareServiceClient.GetShareClient(fileshareName);
                    var directoryClient = shareClient.GetDirectoryClient(faxDeliveryReceipts);
                    var subject = string.Empty;


                    await foreach (ShareFileItem fileItem in directoryClient.GetFilesAndDirectoriesAsync())
                    {
                        if (fileItem.FileSize > 0)
                        {
                            string fileName = Path.GetFileName(fileItem.Name);
                            string fileNameGuid = Guid.NewGuid().ToString().Replace("-", "");
                            logger.LogInformation(DateTime.UtcNow.ToString() + " Copying File");
                            logger.LogInformation(fileItem.Name);

                            var fileClient = directoryClient.GetFileClient(fileItem.Name);
                            var downloadMsgFile = await fileClient.OpenReadAsync();

                            var folderName = Guid.NewGuid().ToString().Replace("-", "");

                            var str1 = Path.Combine(
                                Environment.CurrentDirectory,
                                "bin",
                                "PitCrew",
                                "Fax Delivery Receipts",
                                "ConversionFolder",
                                fileItem.Name);

                            using (var fs = File.Create(str1))
                            {
                                downloadMsgFile.CopyTo(fs);
                                fs.Dispose();
                            }

                            string str2 = Path.Combine(
                                Environment.CurrentDirectory, "bin", "PitCrew", "Fax Delivery Receipts", "ConversionFolder", "MsgToPdf");

                            logger.LogInformation(DateTime.UtcNow.ToString() + " Extracting File");

                            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                            Reader reader = new Reader();
                            string[] toFolder = reader.ExtractToFolder(str1, str2, true);

                            // Specify the path to the wkhtmltopdf executable
                            string wkhtmltopdfPath = Path.Combine(Environment.CurrentDirectory, "bin", "wkhtmltopdf");


                            // Prepare the command to execute
                            string command = $"{wkhtmltopdfPath} - -"; // Read from standard input and write to standard output

                            // Create a process start info
                            var startInfo = new ProcessStartInfo
                            {
                                FileName = "cmd.exe",
                                Arguments = $"/C {command}",
                                UseShellExecute = false,
                                RedirectStandardInput = true,
                                RedirectStandardOutput = true,
                                RedirectStandardError = true, // If you want to capture error output
                                WorkingDirectory = Environment.CurrentDirectory, // Set the working directory if necessary
                            };

                            CovertHtmlToPdf(toFolder);

                            var backupFaxDeliveryReceiptsLocalPath = Path.Combine(
                                Environment.CurrentDirectory,
                                "bin",
                                "PitCrew",
                                "BackupFax Delivery Receipts",
                                fileItem.Name);

                            File.Delete(toFolder[0]);
                            File.Delete(backupFaxDeliveryReceiptsLocalPath);

                            int result = 0;
                            if (int.TryParse(Regex.Match(toFolder[0].ToString(), "(P[0-9]+)").Value.TrimStart('P'), out result))
                            {
                                // Read the connection string from appsettings.json
                                var connectionString = configuration.GetConnectionString("DefaultConnection");

                                // Configure DbContextOptionsBuilder
                                var optionsBuilder = new DbContextOptionsBuilder<dbAgencyContext>();

                                optionsBuilder.UseSqlServer(connectionString);

                                // Create an instance of your DbContext using the configured options
                                using (dbAgencyContext dbAgencyEntities = new dbAgencyContext(optionsBuilder.Options))
                                {
                                    var policyinfo = dbAgencyEntities.tblPolicies.Where(i => i.ID == result)
                                        .Select(i => new
                                        {
                                            i.ID,
                                            i.CustomerID,
                                            i.Profile.Level.License.CompanyID,
                                            i.Profile.Level.ProductID
                                        }).FirstOrDefault();

                                    try
                                    {
                                        int id = policyinfo.CompanyID;
                                    }
                                    catch (Exception ex)
                                    {
                                        SendGridMessage message = new SendGridMessage();
                                        message.Subject = "Policy Not found";
                                        message.SetFrom(new EmailAddress("connect.service@integritymarketing.com", "Test"));
                                        message.AddContent("text/plain", connectionString);
                                        message.AddTo("suresh.palepu@integritymarketing.com");
                                        var SendGridClient = new SendGridClient("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");
                                        var response = await SendGridClient.SendEmailAsync(message);
                                    }
                                    if (policyinfo == null)
                                    {
                                        //MoveFile(fileItem, str1, toFolder);
                                        await UploadPdfToFinishedDeliveryReceiptsBlob(fileItem, toFolder, startInfo, cancellationToken);
                                        await MoveFileAsync(backupFaxDeliveryReceiptsPath, fileName);
                                        await MoveFileAsync(conversionPath, fileName);
                                        await DeleteFileFromAzureFileShareAsync(
                                            this.connectionString, fileshareName, "PitCrew/Fax Delivery Receipts", fileName);
                                        continue;
                                    }
                                    tblDocument entity = new tblDocument();
                                    entity.AgentID = (int?)null;
                                    entity.AppDate = new DateTime?();
                                    entity.AssignedToTeamID = null;
                                    entity.AssignedToUserID = null;
                                    entity.ChangeForm = new bool?(false);
                                    entity.Committed = true;
                                    entity.CompanyID = policyinfo.CompanyID;
                                    entity.Compressed = false;
                                    entity.CreatedBy = new int?(434);
                                    entity.CreatedDate = new DateTime?(DateTime.UtcNow);
                                    entity.CustomerID = policyinfo.CustomerID;
                                    entity.Date = new DateTime?(DateTime.UtcNow);
                                    entity.ErrorCount = 0;
                                    entity.File512Hash = (string)null;
                                    entity.Filename = Guid.NewGuid().ToString().Replace("-", "") + ".pdf";
                                    entity.FileType = new int?(28);
                                    entity.Flag = (string)null;
                                    entity.FName = null;
                                    entity.LicenseID = new int?(0);
                                    entity.LName = null;
                                    entity.MasterDocID = null;
                                    entity.Notes = (string)null;
                                    entity.OwnerID = new int?(434);
                                    entity.Paperport = true;
                                    entity.Path = (string)null;
                                    entity.Policy_No = null;
                                    entity.PolicyID = new int?(policyinfo.ID);
                                    entity.ProductID = policyinfo.ProductID;
                                    entity.SSN = null;
                                    entity.SubmissionReceived = false;
                                    entity.TypeID = new int?(28);
                                    entity.Updated = false;
                                    entity.Upload = true;
                                    entity.Uploaded = true;
                                    entity.Volume = "";
                                    entity.WorkflowStatusID = 0;
                                    dbAgencyEntities.tblDocuments.Add(entity);
                                    string str3 = entity.ID.ToString();

                                    var imagingPath = Path.Combine(
                                        Environment.CurrentDirectory,
                                        "bin",
                                        "Imaging");

                                    string destFileName = string.Concat(new object[4]
                                    {
                                        (object) Path.Combine(Environment.CurrentDirectory, "bin", "Imaging") + "\\",
                                        (object) str3.ToCharArray()[str3.Length - 1],
                                        (object) "\\",
                                        (object) entity.Filename
                                    });


                                    //File.Move(fi.FullName, "\\\\psm-fil-02\\PitCrew\\BackupFax Delivery Receipts\\" + Path.GetFileName(fi.FullName));
                                    await MoveFileAsync(backupFaxDeliveryReceiptsPath, fileName);
                                    await MoveFileAsync(conversionPath, fileName);
                                    await DeleteFileFromAzureFileShareAsync(
                                        this.connectionString, fileshareName, "PitCrew/Fax Delivery Receipts", fileName);

                                    File.Move(str1, backupFaxDeliveryReceiptsLocalPath);

                                    var path = Path.Combine(Environment.CurrentDirectory, "bin", "Imaging") + "\\" + str3.ToCharArray()[str3.Length - 1];

                                    if (!Directory.Exists(path)) 
                                    {
                                        Directory.CreateDirectory(path);
                                    }

                                    File.Move(Path.ChangeExtension(toFolder[0], ".pdf"), destFileName);

                                    await UploadFileToImagingBlob(
                                        entity.Filename, destFileName, str3.ToCharArray()[str3.Length - 1].ToString());


                                    using (MemoryStream ms = new MemoryStream(File.ReadAllBytes(destFileName)))
                                    {
                                        BlobUtility.UploadFile(BlobUtility.PolicyDocuments, entity.Guid.Value, ms);
                                    }
                                    dbAgencyEntities.SaveChanges();

                                    File.Delete(backupFaxDeliveryReceiptsLocalPath);
                                    File.Delete(destFileName);
                                }
                            }
                            else
                            {
                                SendGridMessage message = new SendGridMessage();
                                message.Subject = "Regular Experssion Failed";
                                message.SetFrom(new EmailAddress("connect.service@integritymarketing.com", "Test"));
                                message.AddContent("text/plain", "Regular Experssion Failed");
                                message.AddTo("suresh.palepu@integritymarketing.com");
                                var SendGridClient = new SendGridClient("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");
                                var response = await SendGridClient.SendEmailAsync(message);
                                await UploadPdfToFinishedDeliveryReceiptsBlob(fileItem, toFolder, startInfo, cancellationToken);
                                await MoveFileAsync(backupFaxDeliveryReceiptsPath, fileName);
                                await MoveFileAsync(conversionPath, fileName);
                                await DeleteFileFromAzureFileShareAsync(
                                        this.connectionString, fileshareName, "PitCrew/Fax Delivery Receipts", fileName);
                                //MoveFile(fi, toFolder);
                                File.Delete(str1);
                                //Directory.Delete(str1, true);
                                logger.LogInformation(DateTime.UtcNow.ToString() + " Finished File");
                            }


                        }
                    }
                }
                catch (Exception ex)
                {
                    logger.LogInformation(ex.InnerException.ToString());

                    SendGridMessage message = new SendGridMessage();
                    message.Subject = "Error";
                    message.SetFrom(new EmailAddress("connect.service@integritymarketing.com", "Test"));
                    message.AddContent("text/plain", ex.StackTrace);
                    message.AddTo("suresh.palepu@integritymarketing.com");
                    var SendGridClient = new SendGridClient("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");
                    var response = await SendGridClient.SendEmailAsync(message);
                }

                await Task.Delay(TimeSpan.FromMilliseconds(10000), cancellationToken);
            }
        }

        private async Task UploadPdfToFinishedDeliveryReceiptsBlob(ShareFileItem fileItem, string[] toFolder, ProcessStartInfo startInfo, CancellationToken cancellationToken)
        {
            using (var process = new Process { StartInfo = startInfo })
            {
                process.Start();

                var htmlContent = File.ReadAllText(toFolder[0]);
                // Write the HTML content to the process's standard input
                await process.StandardInput.WriteLineAsync(htmlContent);
                process.StandardInput.Close(); // Close the input stream to signal completion

                // Read the PDF output from the process's standard output
                var pdfStream = new MemoryStream();
                await process.StandardOutput.BaseStream.CopyToAsync(pdfStream);
                pdfStream.Seek(0, SeekOrigin.Begin);

                await CreateFileAsync(
                   Path.ChangeExtension(fileItem.Name, ".pdf"),
                   finishedReceipts,
                   pdfStream,
                   fileshareName,
                   cancellationToken);
            }
        }

        private static void CovertHtmlToPdf(string[] toFolder)
        {
            Process processHtmlToPdf = new Process();
            processHtmlToPdf.StartInfo = new ProcessStartInfo()
            {
                FileName = "wkhtmltopdf.exe",
                Arguments = "\"" + toFolder[0] + "\" \"" + Path.ChangeExtension(toFolder[0], ".pdf") + "\"",
                CreateNoWindow = false,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            processHtmlToPdf.Start();
            processHtmlToPdf.WaitForExit();
        }

        public async Task StartAsync(CancellationToken cancellationToken)
        {
            logger.LogInformation("Service Started.");
            await RunAsync(cancellationToken);
        }

        public Task StopAsync(CancellationToken cancellationToken)
        {
            logger.LogInformation("Service Stopped.");            
            return Task.CompletedTask;
        }

        public async Task CreateFileAsync(
            string fileName, string folder, MemoryStream content, string shareName, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var fileShareClient = await GetFileShareClientAsync(
                fileName, folder, shareName, cancellationToken);

            await fileShareClient.CreateAsync(
                content.Length, null, null, null, null, null, cancellationToken);

            await fileShareClient.UploadRangeAsync(
                new HttpRange(0, content.Length),
                content,
                null,
                null,
                null,
                cancellationToken);

            logger.LogInformation("PDF uploaded successfully!");
        }

        private async Task<ShareFileClient> GetFileShareClientAsync(
            string fileName, string folder, string shareName, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var shareClient = new ShareClient(connectionString, shareName);

            if (!await shareClient.ExistsAsync(cancellationToken))
            {
                await shareClient.CreateAsync(null, null, cancellationToken);
            }

            var directory = shareClient.GetDirectoryClient(folder);

            if (!await directory.ExistsAsync(cancellationToken))
            {
                await directory.CreateAsync(null, null, null, cancellationToken);
            }

            return directory.GetFileClient(fileName);
        }

        private async Task MoveFileAsync(string destinationFilePath, string fileName)
        {
            var serviceClient = new ShareServiceClient(connectionString);

            var shareClient = serviceClient.GetShareClient(fileshareName);

            string sourceFilePath = string.Format(
                "{0}/{1}", configuration["AzureFileStorage:FaxDeliveryReceiptPath"], fileName);

            var sourceDirectory = shareClient.GetDirectoryClient(sourceFilePath);
            //var fileClient = sourceDirectory.GetFileClient(fileName);

            var destinationDirectory = shareClient.GetDirectoryClient(destinationFilePath);

            if (!destinationDirectory.Exists())
            {
                destinationDirectory.Create();
            }

            var destinationFile = destinationDirectory.GetFileClient(Path.GetFileName(sourceFilePath));

            try
            {
                destinationFile.StartCopy(sourceDirectory.Uri);
                //fileClient.Delete(); // Optionally, delete the source file after copying
                //await DeleteFileFromAzureFileShareAsync(
                //    connectionString, fileshareName, "PitCrew/Fax Delivery Receipts", fileName);
                logger.LogInformation("File moved successfully.");
            }
            catch (RequestFailedException ex)
            {
                logger.LogInformation($"Error moving the file: {ex.Message}");
            }

        }

        public async Task DeleteFileFromAzureFileShareAsync(
            string connectionString, 
            string shareName,
            string directoryName, 
            string fileName)
        {
            ShareServiceClient serviceClient = new ShareServiceClient(connectionString);
            ShareClient shareClient = serviceClient.GetShareClient(shareName);
            ShareDirectoryClient directoryClient = shareClient.GetDirectoryClient(directoryName);
            ShareFileClient fileClient = directoryClient.GetFileClient(fileName);



            if (await fileClient.ExistsAsync())
            {
                await fileClient.DeleteAsync();
                Console.WriteLine($"File '{fileName}' deleted successfully.");
            }
            else
            {
                Console.WriteLine($"File '{fileName}' does not exist in the Azure File Share.");
            }
        }

        private async Task UploadFileToImagingBlob(
            string fileName, string localPath, string destinationFilePath)
        {
            var shareClient = shareServiceClient.GetShareClient(
                configuration["AzureFileStorage:PremierInformationTechnologyFileShareName"]);

            //var directoryClient = shareClient.GetDirectoryClient(
            //    configuration["AzureFileStorage:Imaging"]);            

            var destinationDirectory = shareClient.GetDirectoryClient(
                configuration["AzureFileStorage:Imaging"] +"/" + destinationFilePath);

            if (!destinationDirectory.Exists())
            {
                destinationDirectory.Create();
            }

            var fileClient = destinationDirectory.GetFileClient(fileName);
            using (var fileStream = File.OpenRead(localPath))
            {
                await fileClient.CreateAsync(fileStream.Length);
                await fileClient.UploadRangeAsync(new HttpRange(0, fileStream.Length), fileStream);
            }
        }

        public static void MoveFile(ShareFileItem fi, string sourceFile, string[] htmFile)
        {
            var backupFaxDeliveryReceiptsPath = Path.Combine(
                Environment.CurrentDirectory,
                "bin",
                "PitCrew",
                "BackupFax Delivery Receipts",
                fi.Name);

            var finishedReceipts = Path.Combine(
                Environment.CurrentDirectory,
                "bin",
                "PitCrew",
                "Fax Delivery Receipts",
                "Finished Receipts");

            //File.Move(fi.FullName, "\\\\psm-fil-02\\PitCrew\\BackupFax Delivery Receipts\\" + Path.GetFileName(fi.FullName));
            File.Move(sourceFile, backupFaxDeliveryReceiptsPath);

            //File.Delete("\\\\psm-fil-02\\PitCrew\\Fax Delivery Receipts\\Finished Receipts\\" + Path.GetFileName(Path.ChangeExtension(htmFile[0], ".pdf")));
            File.Delete(finishedReceipts + "\\" + Path.GetFileName(Path.ChangeExtension(htmFile[0], ".pdf")));

            //File.Move(Path.ChangeExtension(htmFile[0], ".pdf"), "\\\\psm-fil-02\\PitCrew\\Fax Delivery Receipts\\Finished Receipts\\" + Path.GetFileName(Path.ChangeExtension(htmFile[0], ".pdf")));
            File.Move(Path.ChangeExtension(htmFile[0], ".pdf"), finishedReceipts + "\\" + Path.GetFileName(Path.ChangeExtension(htmFile[0], ".pdf")));
        }
    }
}
