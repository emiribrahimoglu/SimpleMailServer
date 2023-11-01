using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Data;
using Google.Apis.Auth;
using Google.Apis.Oauth2;
using Google.Apis.Auth.OAuth2;
using System.Threading;
using SmtpClient = System.Net.Mail.SmtpClient;
using MailKit.Security;
using MailKit.Net.Smtp;
using Google.Apis.Gmail.v1;
using MimeKit;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Util.Store;
using Google.Apis.Util;
using Microsoft.Identity.Client;
using System.Security.Principal;

namespace SimpleMailServer
{
    internal class Program
    {
        #region Configuration Declarations
        static string sqlConnectionString = string.Empty;
        static string appDirectory = AppDomain.CurrentDomain.BaseDirectory;
        static SmtpClient smtpClient = null;

        static string smtpHost = string.Empty;
        static int smtpPort = 0;
        static string smtpUsername = string.Empty;
        static string smtpPassword = string.Empty;
        static bool smtpEnableTLS = false;
        static int tryCount = 0;
        static int mailsPerSession = 0;
        static string smtpDefaultMailFrom = string.Empty;

        static string googleClientId = string.Empty;
        static string googleClientSecret = string.Empty;
        static string googleAccessToken = string.Empty;

        static string microsoftTenantId = string.Empty;
        static string microsoftClientId = string.Empty;
        static string microsoftAccessToken = string.Empty;

        static bool smtpIsLogginEnabled = false;

        static bool isMailSent = false;
        #endregion

        #region Variable Declarations
        static DataTable MailsToSendDT = null;
        static DataTable mailAttachmentsDT = null;

        static DateTime runDate = DateTime.Now;
        //string logFileName = "SimpleMailServer-"+runDate.Year + "_" + runDate.Month + "_" + runDate.Day + "-" + runDate.Hour + "_" + runDate.Minute + "_" + runDate.Second + ".txt";
        static string logFileName = "SimpleMailServer-" + runDate.Year + "_" + runDate.Month + "_" + runDate.Day + ".txt";
        static string subfolders = "Logs" + "\\" + runDate.Year + "\\" + runDate.Month + "\\" + runDate.Day + "\\";
        static string logFilePath = Path.Combine(appDirectory + subfolders, logFileName);
        static string configDirectory = Path.Combine(appDirectory, "MailConfigs\\");
        static string mailFrom = string.Empty;
        static string mailTO = string.Empty;
        static string subject = string.Empty;
        static string body = string.Empty;
        static string mailId = string.Empty;
        static string tempTryCount = string.Empty;

        static DateTime currentTime = DateTime.Now;

        static MailMessage msg = null;
        static Attachment att = null;
        static MemoryStream attStream = null;
        static byte[] bytes = null;
        static string attFilename = string.Empty;

        static bool isBasicAuth = false;
        static bool isMicrosotOauth2 = false;
        static bool isGoogleOauth2 = false;

        static UserCredential googleCredentials = null;
        static MailKit.Net.Smtp.SmtpClient mailkitClient = null;
        static MimeMessage mimeMsg = null;
        #endregion

        static void Main(string[] args)
        {
            try
            {
                sqlConnectionString = File.ReadAllText(configDirectory + "SqlConnectionString.txt");
            }
            catch (Exception ee)
            {

                throw new Exception("SqlConnectionString.txt dosyası uygulamanın dizininde bulunamadı veya dosyaya erişilemedi.\n" + ee.Message);
            }

            LoadConfigurations();
            GetAuthType();
            GetTryCountAndMailsPerSession();
            if (isBasicAuth)
            {
                InitializeBasicSMTP();
                BasicAuthenticationSmtpSend();
            }
            else if (isGoogleOauth2)
            {
                GetGoogleOauthToken();
                SendMailGoogleOauth2();
            }
            else
            {
                /*
                 * GetMicrosoftOauthToken();
                 * SendMailMicrosoftOauth2();
                */
                AddToLog(DateTime.Now, logFilePath, "Microsoft Oauth2 yöntemini kullanmaya çalıştınız.\nMicrosoft Oauth2 kodu yazıldı ancak test edilmediği için aktif değil. Lütfen başka bir yöntem kullanın.");
            }
            
            
        }

        static void LoadConfigurations()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(configDirectory));

            smtpHost = File.ReadAllText(configDirectory + "SmtpHost.txt");
            smtpPort = int.Parse(File.ReadAllText(configDirectory + "SmtpPort.txt"));
            smtpUsername = File.ReadAllText(configDirectory + "SmtpUsername.txt");
            smtpPassword = File.ReadAllText(configDirectory + "SmtpPassword.txt");
            smtpEnableTLS = bool.Parse(File.ReadAllText(configDirectory + "SmtpTls.txt"));
            smtpDefaultMailFrom = File.ReadAllText(configDirectory + "SmtpDefaultMailFrom.txt");
            smtpIsLogginEnabled = bool.Parse(File.ReadAllText(configDirectory + "SmtpIsLoggingEnabled.txt"));

            googleClientId = File.ReadAllText(configDirectory + "GoogleOauth2ClientId.txt");
            googleClientSecret = File.ReadAllText(configDirectory + "GoogleOauth2ClientSecret.txt");

            microsoftTenantId = File.ReadAllText(configDirectory + "MicrosoftOauth2TenantId.txt");
            microsoftClientId = File.ReadAllText(configDirectory + "MicrosoftOauth2ClientId.txt");
        }

        static async void GetMicrosoftOauthToken()
        {
            if (string.IsNullOrEmpty(File.ReadAllText(configDirectory + "MicrosoftOauth2AccessToken.txt")))
            {
                var options = new PublicClientApplicationOptions
                {
                    ClientId = microsoftClientId,
                    TenantId = microsoftTenantId,

                    // Use "https://login.microsoftonline.com/common/oauth2/nativeclient" for apps using
                    // embedded browsers or "http://localhost" for apps that use system browsers.
                    RedirectUri = "http://localhost"
                };

                var publicClientApplication = PublicClientApplicationBuilder
                    .CreateWithApplicationOptions(options)
                    .Build();

                var scopes = new string[] {
                    "email",
                    "offline_access",
                    //"https://outlook.office.com/IMAP.AccessAsUser.All", // Only needed for IMAP
                    //"https://outlook.office.com/POP.AccessAsUser.All",  // Only needed for POP
                    "https://outlook.office.com/SMTP.Send" // Only needed for SMTP
                };

                var microsoftOauthToken = await publicClientApplication.AcquireTokenInteractive(scopes).ExecuteAsync();

                File.Create(configDirectory + "MicrosoftOauth2AccessToken.txt").Dispose();

                using (TextWriter tw = new StreamWriter(configDirectory + "MicrosoftOauth2AccessToken.txt"))
                {
                    tw.WriteLine(microsoftOauthToken.AccessToken);
                    microsoftAccessToken = microsoftOauthToken.AccessToken;
                }
            }
            else
            {
                microsoftAccessToken = File.ReadAllText(configDirectory + "MicrosoftOauth2AccessToken.txt");
            }
        }

        static void SendMailMicrosoftOauth2()
        {
            MailsToSendDT = GetMailsToSend();
            foreach (DataRow row in MailsToSendDT.Rows)
            {
                mailFrom = row["MAILFROM"].ToString();
                mailTO = row["MAILTO"].ToString();
                subject = row["SUBJECT"].ToString();
                body = row["MAILBODY"].ToString();
                tempTryCount = row["TRYCOUNT"].ToString();
                mailId = row["ID"].ToString();

                mailAttachmentsDT = GetMailAttachments(mailId);

                if (string.IsNullOrEmpty(mailFrom))
                {
                    if (!string.IsNullOrEmpty(mailTO))
                    {
                        mimeMsg = new MimeMessage();
                        mimeMsg.From.Add(MailboxAddress.Parse(smtpDefaultMailFrom));
                        mimeMsg.To.Add(new MailboxAddress(mailTO, mailTO));
                        mimeMsg.Subject = subject;
                        BodyBuilder builder = new BodyBuilder();
                        builder.HtmlBody = body;
                        foreach (DataRow row2 in mailAttachmentsDT.Rows)
                        {
                            bytes = (byte[])row2["DATA"];
                            attFilename = row2["FILENAME"].ToString();
                            builder.Attachments.Add(attFilename, bytes);
                        }
                        mimeMsg.Body = builder.ToMessageBody();
                    }

                }
                else
                {
                    if (!string.IsNullOrEmpty(mailTO))
                    {
                        mimeMsg = new MimeMessage();
                        mimeMsg.From.Add(MailboxAddress.Parse(mailFrom));
                        mimeMsg.To.Add(new MailboxAddress(mailTO, mailTO));
                        mimeMsg.Subject = subject;
                        BodyBuilder builder = new BodyBuilder();
                        builder.HtmlBody = body;
                        foreach (DataRow row2 in mailAttachmentsDT.Rows)
                        {
                            bytes = (byte[])row2["DATA"];
                            attFilename = row2["FILENAME"].ToString();
                            builder.Attachments.Add(attFilename, bytes);
                        }
                        mimeMsg.Body = builder.ToMessageBody();
                    }
                }

                try
                {
                    using (mailkitClient = new MailKit.Net.Smtp.SmtpClient())
                    {
                        mailkitClient.ServerCertificateValidationCallback = (s, c, h, e) => true;
                        mailkitClient.Connect(smtpHost, smtpPort, SecureSocketOptions.StartTls);
                        SaslMechanismOAuth2 oauth2 = new SaslMechanismOAuth2(smtpUsername, microsoftAccessToken);
                        mailkitClient.Authenticate(oauth2);
                        mailkitClient.Send(mimeMsg);
                        mailkitClient.Disconnect(true);
                    }

                    isMailSent = true;
                }
                catch (Exception ee)
                {
                    isMailSent = false;
                    currentTime = DateTime.Now;
                    AddToLog(currentTime, mailId, logFilePath, "Mail gönderilirken hata alındı\n" + ee.Message + "\nMail To: " + mailTO + "(Bu parantezin solu boşsa mail bu yüzden iletilememiştir)");
                }

                try
                {
                    if (isMailSent)
                    {
                        currentTime = DateTime.Now;
                        UpdateSentDate(currentTime, mailId);
                        IncreaseTryCount(mailId, tempTryCount);
                        AddToLog(currentTime, mailId, logFilePath, "Mail başarıyla gönderildi\n");
                        isMailSent = false;
                    }
                    else
                    {
                        IncreaseTryCount(mailId, tempTryCount);
                        currentTime = DateTime.Now;
                        AddToLog(currentTime, mailId, logFilePath, "Mail gönderilemedi. Veritabanında trycount değeri güncellendi\n");
                    }

                }
                catch (Exception ee)
                {
                    currentTime = DateTime.Now;
                    AddToLog(currentTime, mailId, logFilePath, "Mail başarıyla gönderildi ancak veritabanı tarafında güncelleme ya da işlemin loglanması sırasında hata oluştu\n" + ee.Message);
                    isMailSent = false;
                }

                mimeMsg = null;
                bytes = null;
                attFilename = String.Empty;
            }
        }

        static void GetAuthType()
        {
            string[] auths = File.ReadAllLines(configDirectory + "SmtpAuthType.txt");
            for (int i = 0; i < auths.Length; i++)
            {
                switch (auths[i].Split('=')[0].ToLower())
                {
                    case "basicauth":
                        isBasicAuth = bool.Parse(auths[i].Split('=')[1].ToLower());
                        break;
                    case "microsoftoauth2":
                        isMicrosotOauth2 = bool.Parse(auths[i].Split('=')[1].ToLower());
                        break;
                    case "googleoauth2":
                        isGoogleOauth2 = bool.Parse(auths[i].Split('=')[1].ToLower());
                        break;
                    default:
                        AddToLog(currentTime, logFilePath, "SmtpAuthType text dosyasının formatlamasında bir hata mevcut. Dosya okunamıyor.");
                        throw new Exception("SmtpAuthType text dosyasının formatlamasında bir hata mevcut. Dosya okunamıyor.");
                        break;
                    
                }
            }
        }

        static async void GetGoogleOauthToken()
        {
            var clientSecrets = new ClientSecrets
            {
                ClientId = googleClientId,
                ClientSecret = googleClientSecret
            };

            var codeFlow = new GoogleAuthorizationCodeFlow(new GoogleAuthorizationCodeFlow.Initializer
            {
                DataStore = new FileDataStore("CredentialCacheFolder", false),
                Scopes = new[] { "https://mail.google.com/" },
                ClientSecrets = clientSecrets
            });

            // Note: For a web app, you'll want to use AuthorizationCodeWebApp instead.
            var codeReceiver = new LocalServerCodeReceiver();
            var authCode = new AuthorizationCodeInstalledApp(codeFlow, codeReceiver);

            googleCredentials = await authCode.AuthorizeAsync(smtpUsername, CancellationToken.None);

            if (googleCredentials.Token.IsExpired(SystemClock.Default))
                await googleCredentials.RefreshTokenAsync(CancellationToken.None);

            File.Create(configDirectory + "GoogleOauth2Token.txt").Dispose();

            using (TextWriter tw = new StreamWriter(configDirectory + "GoogleOauth2Token.txt"))
            {
                tw.WriteLine(googleCredentials.Token.AccessToken);
                googleAccessToken = googleCredentials.Token.AccessToken;
            }
            
        }

        static void SendMailGoogleOauth2()
        {
            MailsToSendDT = GetMailsToSend();
            foreach (DataRow row in MailsToSendDT.Rows)
            {
                mailFrom = row["MAILFROM"].ToString();
                mailTO = row["MAILTO"].ToString();
                subject = row["SUBJECT"].ToString();
                body = row["MAILBODY"].ToString();
                tempTryCount = row["TRYCOUNT"].ToString();
                mailId = row["ID"].ToString();

                mailAttachmentsDT = GetMailAttachments(mailId);

                if (string.IsNullOrEmpty(mailFrom))
                {
                    if (!string.IsNullOrEmpty(mailTO))
                    {
                        mimeMsg = new MimeMessage();
                        mimeMsg.From.Add(MailboxAddress.Parse(smtpDefaultMailFrom));
                        mimeMsg.To.Add(new MailboxAddress(mailTO, mailTO));
                        mimeMsg.Subject = subject;
                        BodyBuilder builder = new BodyBuilder();
                        builder.HtmlBody = body;
                        foreach (DataRow row2 in mailAttachmentsDT.Rows)
                        {
                            bytes = (byte[])row2["DATA"];
                            attFilename = row2["FILENAME"].ToString();
                            builder.Attachments.Add(attFilename, bytes);
                        }
                        mimeMsg.Body = builder.ToMessageBody();
                    }

                }
                else
                {
                    if (!string.IsNullOrEmpty(mailTO))
                    {
                        mimeMsg = new MimeMessage();
                        mimeMsg.From.Add(MailboxAddress.Parse(mailFrom));
                        mimeMsg.To.Add(new MailboxAddress(mailTO, mailTO));
                        mimeMsg.Subject = subject;
                        BodyBuilder builder = new BodyBuilder();
                        builder.HtmlBody = body;
                        foreach (DataRow row2 in mailAttachmentsDT.Rows)
                        {
                            bytes = (byte[])row2["DATA"];
                            attFilename = row2["FILENAME"].ToString();
                            builder.Attachments.Add(attFilename, bytes);
                        }
                        mimeMsg.Body = builder.ToMessageBody();
                    }
                }

                try
                {
                    using (mailkitClient = new MailKit.Net.Smtp.SmtpClient())
                    {
                        mailkitClient.ServerCertificateValidationCallback = (s, c, h, e) => true;
                        mailkitClient.Connect(smtpHost, smtpPort, SecureSocketOptions.StartTls);
                        SaslMechanismOAuth2 oauth2 = new SaslMechanismOAuth2(smtpUsername, googleAccessToken);
                        mailkitClient.Authenticate(oauth2);
                        mailkitClient.Send(mimeMsg);
                        mailkitClient.Disconnect(true);
                    }
                    
                    isMailSent = true;
                }
                catch (Exception ee)
                {
                    isMailSent = false;
                    currentTime = DateTime.Now;
                    AddToLog(currentTime, mailId, logFilePath, "Mail gönderilirken hata alındı\n" + ee.Message + "\nMail To: " + mailTO + "(Bu parantezin solu boşsa mail bu yüzden iletilememiştir)");
                }

                try
                {
                    if (isMailSent)
                    {
                        currentTime = DateTime.Now;
                        UpdateSentDate(currentTime, mailId);
                        IncreaseTryCount(mailId, tempTryCount);
                        AddToLog(currentTime, mailId, logFilePath, "Mail başarıyla gönderildi\n");
                        isMailSent = false;
                    }
                    else
                    {
                        IncreaseTryCount(mailId, tempTryCount);
                        currentTime = DateTime.Now;
                        AddToLog(currentTime, mailId, logFilePath, "Mail gönderilemedi. Veritabanında trycount değeri güncellendi\n");
                    }

                }
                catch (Exception ee)
                {
                    currentTime = DateTime.Now;
                    AddToLog(currentTime, mailId, logFilePath, "Mail başarıyla gönderildi ancak veritabanı tarafında güncelleme ya da işlemin loglanması sırasında hata oluştu\n" + ee.Message);
                    isMailSent = false;
                }

                mimeMsg = null;
                bytes = null;
                attFilename = String.Empty;
            }
            
        }

        static void InitializeBasicSMTP()
        {
            smtpClient = new SmtpClient()
            {
                Host = smtpHost,
                Port = smtpPort,
                Credentials = new NetworkCredential(smtpUsername, smtpPassword),
                EnableSsl = smtpEnableTLS
            };
        }

        static void GetTryCountAndMailsPerSession()
        {
            tryCount = int.Parse(File.ReadAllText(configDirectory + "MailTryCount.txt"));
            mailsPerSession = int.Parse(File.ReadAllText(configDirectory + "MailsPerSession.txt"));
        }

        static DataTable GetMailsToSend()
        {
            DataTable dt = new DataTable();
            using (SqlConnection con = new SqlConnection(sqlConnectionString))
            {
                string query = "SELECT TOP (@MailsPerSession) * FROM MAILS WHERE SENTDATE IS NULL AND TRYCOUNT != @TryCount ORDER BY ID DESC";
                SqlCommand cmd = new SqlCommand(query, con);
                cmd.Parameters.AddWithValue("@TryCount", tryCount.ToString());
                cmd.Parameters.AddWithValue("@MailsPerSession", mailsPerSession);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
            }
            return dt;
        }

        static void IncreaseTryCount(string MailID, string MailTryCount)
        {
            using (SqlConnection con = new SqlConnection(sqlConnectionString))
            {
                MailTryCount = (int.Parse(MailTryCount) + 1).ToString();
                string query = "UPDATE MAILS SET TRYCOUNT = @TryCount WHERE ID = @Id";
                SqlCommand cmd = new SqlCommand(query, con);
                cmd.Parameters.AddWithValue("@TryCount", MailTryCount);
                cmd.Parameters.AddWithValue("@Id", MailID);
                cmd.Connection.Open();
                cmd.ExecuteNonQuery();
                cmd.Connection.Close();
            }
        }

        static void UpdateSentDate(DateTime SentDate, string MailID)
        {
            using (SqlConnection con = new SqlConnection(sqlConnectionString))
            {
                string query = "UPDATE MAILS SET SENTDATE = @DateTime WHERE ID = @Id";
                SqlCommand cmd = new SqlCommand(query, con);
                cmd.Parameters.AddWithValue("@DateTime", SentDate);
                cmd.Parameters.AddWithValue("@Id", MailID);
                cmd.Connection.Open();
                cmd.ExecuteNonQuery();
                cmd.Connection.Close();
            }
        }

        static void AddToLog(DateTime now, string MailID, string logFilePath, string Message)
        {
            if (smtpIsLogginEnabled)
            {
                if (!File.Exists(logFilePath))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(logFilePath));
                    File.Create(logFilePath).Dispose();

                    using (TextWriter tw = new StreamWriter(logFilePath))
                    {
                        tw.WriteLine(now.ToString() + "\nMail ID: " + MailID + "\n" + Message);
                        tw.WriteLine();
                        tw.WriteLine();
                        tw.WriteLine();
                    }

                }
                else if (File.Exists(logFilePath))
                {
                    using (TextWriter tw = new StreamWriter(logFilePath, true))
                    {
                        tw.WriteLine(now.ToString() + "\nMail ID: " + MailID + "\n" + Message);
                        tw.WriteLine();
                        tw.WriteLine();
                        tw.WriteLine();
                    }
                }
            }
        }

        static void AddToLog(DateTime now, string logFilePath, string Message)
        {
            if (smtpIsLogginEnabled)
            {
                if (!File.Exists(logFilePath))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(logFilePath));
                    File.Create(logFilePath).Dispose();

                    using (TextWriter tw = new StreamWriter(logFilePath))
                    {
                        tw.WriteLine(now.ToString() + "\n" + Message);
                        tw.WriteLine();
                        tw.WriteLine();
                        tw.WriteLine();
                    }

                }
                else if (File.Exists(logFilePath))
                {
                    using (TextWriter tw = new StreamWriter(logFilePath, true))
                    {
                        tw.WriteLine(now.ToString() + "\n" + Message);
                        tw.WriteLine();
                        tw.WriteLine();
                        tw.WriteLine();
                    }
                }
            }
        }

        static DataTable GetMailAttachments(string mailId)
        {
            DataTable dt = new DataTable();
            using (SqlConnection con = new SqlConnection(sqlConnectionString))
            {
                string query = "SELECT FILENAME, DATA FROM MAILATTACHMENTS WHERE MAILID = @MailId";
                SqlCommand cmd = new SqlCommand(query, con);
                cmd.Parameters.AddWithValue("@MailId", int.Parse(mailId));
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
            }
            return dt;

        }

        static void BasicAuthenticationSmtpSend()
        {
            MailsToSendDT = GetMailsToSend();
            foreach (DataRow row in MailsToSendDT.Rows)
            {
                mailFrom = row["MAILFROM"].ToString();
                mailTO = row["MAILTO"].ToString();
                subject = row["SUBJECT"].ToString();
                body = row["MAILBODY"].ToString();
                tempTryCount = row["TRYCOUNT"].ToString();
                mailId = row["ID"].ToString();

                mailAttachmentsDT = GetMailAttachments(mailId);

                if (string.IsNullOrEmpty(mailFrom))
                {
                    if (!string.IsNullOrEmpty(mailTO))
                    {
                        msg = new MailMessage(smtpDefaultMailFrom, mailTO, subject, body);
                        msg.IsBodyHtml = true;
                        foreach (DataRow row2 in mailAttachmentsDT.Rows)
                        {
                            bytes = (byte[])row2["DATA"];
                            attFilename = row2["FILENAME"].ToString();
                            attStream = new MemoryStream(bytes);
                            att = new System.Net.Mail.Attachment(attStream, attFilename);
                            msg.Attachments.Add(att);
                        }
                    }

                }
                else
                {
                    if (!string.IsNullOrEmpty(mailTO))
                    {
                        msg = new MailMessage(mailFrom, mailTO, subject, body);
                        msg.IsBodyHtml = true;
                        foreach (DataRow row2 in mailAttachmentsDT.Rows)
                        {
                            bytes = (byte[])row2["DATA"];
                            attFilename = row2["FILENAME"].ToString();
                            attStream = new MemoryStream(bytes);
                            att = new System.Net.Mail.Attachment(attStream, attFilename);
                            msg.Attachments.Add(att);
                        }
                    }
                }

                try
                {
                    smtpClient.Send(msg);
                    isMailSent = true;
                }
                catch (Exception ee)
                {
                    isMailSent = false;
                    currentTime = DateTime.Now;
                    AddToLog(currentTime, mailId, logFilePath, "Mail gönderilirken hata alındı\n" + ee.Message + "\nMail To: " + mailTO + "(Bu parantezin solu boşsa mail bu yüzden iletilememiştir)");
                }

                try
                {
                    if (isMailSent)
                    {
                        currentTime = DateTime.Now;
                        UpdateSentDate(currentTime, mailId);
                        IncreaseTryCount(mailId, tempTryCount);
                        AddToLog(currentTime, mailId, logFilePath, "Mail başarıyla gönderildi\n");
                        isMailSent = false;
                    }
                    else
                    {
                        IncreaseTryCount(mailId, tempTryCount);
                        currentTime = DateTime.Now;
                        AddToLog(currentTime, mailId, logFilePath, "Mail gönderilemedi. Veritabanında trycount değeri güncellendi\n");
                    }

                }
                catch (Exception ee)
                {
                    currentTime = DateTime.Now;
                    AddToLog(currentTime, mailId, logFilePath, "Mail başarıyla gönderildi ancak veritabanı tarafında güncelleme ya da işlemin loglanması sırasında hata oluştu\n" + ee.Message);
                    isMailSent = false;
                }

                msg = null;
                attStream = null;
                bytes = null;
                att = null;
                attFilename = String.Empty;
            }
        }
    }
}
