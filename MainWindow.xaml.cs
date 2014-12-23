using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Net;
using System.Net.Sockets;
using System.Net.Mail;
using System.Windows.Threading;
using System.Reflection;

using Google;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v2;
using Google.Apis.Drive.v2.Data;
using Google.Apis.Services;

using Json = Newtonsoft.Json;

namespace FindIpAddress
{
   /// <summary>
   /// Interaction logic for MainWindow.xaml
   /// </summary>
   public partial class MainWindow : Window
   {
      private System.Windows.Forms.NotifyIcon notifyIcon;
      private WindowState currentWindowState = WindowState.Minimized;
      private Assembly _assembly;

      public string CurrentIP { get; set; }
      private DriveService GoogleAPI { get; set; }
      private string FileID { get; set; }

      private class MailSecrets
      {
         public string Password { get; set; }
         public string Email { get; set; }
         public string Name { get; set; }

         public MailSecrets()
         {
            this.Password = string.Empty;
            this.Email = string.Empty;
            this.Name = string.Empty;
         }

         public MailSecrets(string email, string pwd, string name)
         {
            this.Email = email;
            this.Password = pwd;
            this.Name = name;
         }
      }

      public MainWindow()
      {
         InitializeComponent();
         try
         {
            _assembly = Assembly.GetExecutingAssembly();
         }
         catch
         {
            MessageBox.Show("Error accessing resources!");
         }
         InitializeIconTray();
         this.DataContext = this;
         this.GoogleAPI = GetGoogleService();
         GetFileId();
         DispatcherTimer dispatcherTimer = new DispatcherTimer();
         dispatcherTimer.Tick += dispatcherTimer_Tick;
         dispatcherTimer.Interval = new TimeSpan(1, 0, 0);
         dispatcherTimer.Start();
      }

      private DriveService GetGoogleService()
      {
         //https://console.developers.google.com/project/named-dialect-796/apiui/credential
         //https://developers.google.com/drive/web/quickstart/quickstart-cs
         //https://developers.google.com/console/help/new/#usingkeys
         //https://developers.google.com/drive/web/examples/dotnet#additional_resources

         UserCredential credential;
         //using(var stream = new System.IO.FileStream("client_secrets.json", System.IO.FileMode.Open, System.IO.FileAccess.Read))
         using(var stream = _assembly.GetManifestResourceStream("FindIpAddress.client_secrets.json"))
         {
            credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                GoogleClientSecrets.Load(stream).Secrets,
                new[] { DriveService.Scope.Drive },
                "user", CancellationToken.None).Result;
         }

         // Create the service.
         var service = new DriveService(new BaseClientService.Initializer()
         {
            HttpClientInitializer = credential,
            ApplicationName = "GetMyIpAddress",
         });

         return service;
      }

      private void GetFileId()
      {
         List<File> result = new List<File>();
         FilesResource.ListRequest listRequest = this.GoogleAPI.Files.List();
         listRequest.Q = "title = 'MyIpAddress.txt'";
         do
         {
            try
            {
               FileList files = listRequest.Execute();

               result.AddRange(files.Items);
               listRequest.PageToken = files.NextPageToken;
            }
            catch(Exception e)
            {
               MessageBox.Show("An error occurred: " + e.Message, "GetFileId error", MessageBoxButton.OK, MessageBoxImage.Error);
               listRequest.PageToken = null;
            }
         } while(!String.IsNullOrEmpty(listRequest.PageToken));

         if(result.Count == 1)
         {
            this.FileID = result[0].Id;
         }
         else
         {
            if(result.Count == 0)
            {
               // File's metadata.
               string mimeType = "text/plain";//"application/vnd.google-apps.document"; //"application/vnd.google-apps.document"
               File body = new File();
               body.Title = "MyIpAddress.txt";
               body.Description = "The updated IP address of my laptop on Internet";
               body.MimeType = mimeType;

               //// Set the parent folder.
               //if (!String.IsNullOrEmpty(parentId)) {
               //  body.Parents = new List<ParentReference>()
               //      {new ParentReference() {Id = parentId}};
               //}

               // File's content.
               System.IO.MemoryStream stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(this.CurrentIP));
               try 
               {
                  FilesResource.InsertMediaUpload request = this.GoogleAPI.Files.Insert(body, stream, mimeType);
                  request.Upload();

                  File file = request.ResponseBody;

                  this.FileID = file.Id;
               } 
               catch (Exception e) 
               {
                  MessageBox.Show("An error occurred: " + e.Message, "File creation error", MessageBoxButton.OK, MessageBoxImage.Error);
               }
            }
            else
            {
               this.FileID = string.Empty;
            }
         }
      }

      void dispatcherTimer_Tick(object sender, EventArgs e)
      {
         string IP = CheckIpAddress();
         if(this.CurrentIP != IP)
         {
            this.notifyIcon.BalloonTipText = this.CurrentIP;
            SendMail();
            UpdateGoogleDocument();
         }         
      }

      private void InitializeIconTray()
      {
         this.notifyIcon = new System.Windows.Forms.NotifyIcon();
         System.Windows.Forms.ContextMenu contextMenu = new System.Windows.Forms.ContextMenu();

         this.CurrentIP = CheckIpAddress();
         this.notifyIcon.BalloonTipText = this.CurrentIP;
         this.notifyIcon.BalloonTipTitle = string.Empty;
         this.notifyIcon.Text = this.CurrentIP;
         this.notifyIcon.Icon = new System.Drawing.Icon(@"Images\iRC_icon.ico");
         this.notifyIcon.Visible = true;
         this.notifyIcon.DoubleClick += new EventHandler(notifyIcon_DoubleClick);
         this.notifyIcon.ContextMenu = contextMenu;

         // Contextual Menu
         // CheckIP
         var menuItem = new System.Windows.Forms.MenuItem();
         menuItem.Index = 0;
         menuItem.Text = @"CheckIP";
         menuItem.Click += menuCheckIPItem_Click;
         contextMenu.MenuItems.Add(menuItem);
         // Send Mail
         menuItem = new System.Windows.Forms.MenuItem();
         menuItem.Index = 0;
         menuItem.Text = @"Send Mail";
         menuItem.Click += menuSendMailItem_Click;
         contextMenu.MenuItems.Add(menuItem);
         // Exit
         menuItem = new System.Windows.Forms.MenuItem();
         menuItem.Index = 0;
         menuItem.Text = @"Exit";
         menuItem.Click += menuExitItem_Click;
         contextMenu.MenuItems.Add(menuItem);
      }

      private void menuExitItem_Click(object sender, EventArgs e)
      {
         this.Close();
      }

      private void menuCheckIPItem_Click(object sender, EventArgs e)
      {
         CheckIpAddress();
         this.notifyIcon.ShowBalloonTip(1000);
      }

      private void menuSendMailItem_Click(object sender, EventArgs e)
      {
         SendMail();
      }

      private string CheckIpAddress()
      {
         WebClient webClient = new WebClient();
         string IP = webClient.DownloadString("http://icanhazip.com");
         return IP;
      }

      private void ButtonSendMail_Click(object sender, RoutedEventArgs e)
      {
         SendMail();
      }

      private void SendMail()
      {
         MailSecrets mailSecret = null;

         using(var stream = _assembly.GetManifestResourceStream("FindIpAddress.mail_secrets.json"))
         {
            Json.JsonSerializer serializer = new Json.JsonSerializer();
            System.IO.StreamReader streamReader = new System.IO.StreamReader(stream);
            Json.JsonTextReader reader = new Json.JsonTextReader(streamReader);
            mailSecret = serializer.Deserialize<MailSecrets>(reader);
         }

         var fromAddress = new MailAddress(mailSecret.Email, mailSecret.Name);
         var toAddress = new MailAddress(mailSecret.Email, mailSecret.Name);
         const string subject = "Adresse";
         string body = CheckIpAddress();

         using(var smtp = new SmtpClient())
         {
            smtp.Host = "smtp.gmail.com";
            smtp.Port = 587;
            smtp.EnableSsl = true;
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtp.UseDefaultCredentials = false;
            smtp.Timeout = 10000;
            smtp.Credentials = new NetworkCredential(fromAddress.Address, mailSecret.Password);

            using(var message = new MailMessage(fromAddress, toAddress))
            {
               message.Subject = subject;
               message.Body = body;
               message.IsBodyHtml = true;
               try
               {
                  smtp.Send(message);
                  if(this.currentWindowState == System.Windows.WindowState.Normal)
                  {
                     MessageBox.Show("eMail sent", "", MessageBoxButton.OK, MessageBoxImage.Information);
                  }
               }
               catch(Exception ep)
               {
                  MessageBox.Show("Exception Occured:" + ep.Message, "Send Mail Error", MessageBoxButton.OK, MessageBoxImage.Error);
               }
            }
         }
      }

      private void OnClose(object sender, System.ComponentModel.CancelEventArgs e)
      {
         notifyIcon.Dispose();
         notifyIcon = null;
      }

      private void Window_StateChanged(object sender, EventArgs e)
      {
         if(WindowState == WindowState.Minimized)
         {
            Hide();
            if(notifyIcon != null)
            {
               notifyIcon.ShowBalloonTip(500);
            }
         }
         else
         {
            currentWindowState = WindowState;
         }
      }

      private void Window_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
      {
         CheckTrayIcon();
      }

      private void notifyIcon_DoubleClick(object sender, EventArgs e)
      {
         Show();
         this.WindowState = System.Windows.WindowState.Normal;
         this.currentWindowState = System.Windows.WindowState.Normal;
      }

      private void CheckTrayIcon()
      {
         ShowTrayIcon(!IsVisible);
      }

      private void ShowTrayIcon(bool show)
      {
         if(notifyIcon != null)
         {
            notifyIcon.Visible = show;
         }
      }

      private void ButtonFind_Click(object sender, RoutedEventArgs e)
      {
         this.CurrentIP = CheckIpAddress();
      }

      private void Window_Loaded(object sender, RoutedEventArgs e)
      {
         this.WindowState = System.Windows.WindowState.Minimized;
      }

      private void UpdateGoogleDocument()
      {
         try 
         {
            // First retrieve the file from the API.
            File file = this.GoogleAPI.Files.Get(this.FileID).Execute();
            
            // File's new content.
            System.IO.MemoryStream stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(this.CurrentIP));

            // Send the request to the API.
            FilesResource.UpdateMediaUpload request = this.GoogleAPI.Files.Update(file, this.FileID, stream, file.MimeType);
            request.NewRevision = true;
            request.Upload();

            //File updatedFile = request.ResponseBody;
            //return updatedFile;
         } 
         catch (Exception e) 
         {
            MessageBox.Show("An error occurred: " + e.Message, "Update document error", MessageBoxButton.OK, MessageBoxImage.Error);
         }
      }
   }
}
