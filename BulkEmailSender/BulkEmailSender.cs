using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Amib.Threading;

namespace BulkEmailSender
{
    public class BulkEmailSender
    {
        #region constants and variables 

        private string newsLetterFilePath = @"C:\email\pup.htm";
        private string recipientEmailsFilePath = @"C:\email\email.txt";
        private string unsubscribedEmailsFilePath = @"C:\email\unsubscribed.txt";
        private string validSubscribedRecipientsFilePath = @"C:\email\FinalEmailList.txt";
        private string logFilePath = @"C:\email\pup_log.txt";
        private string smtpHost = @"news.palmerpartynews.com.au";
        private string senderEmailAddress = @"Clive.Palmer@news.palmerpartynews.com.au";
        private string mailSubject = @"A Message From Clive Palmer";

        private StreamWriter logStreamer;        
        private int totalRecipientCount = 0;
        private int currRecipientIndex = 0;
        private int totalSentCount = 0;
        private int recipientIndexForDoneEvent = 0;

        private ManualResetEvent doneEvent;
        // [ThreadStatic]
        // public static string recipientEmailAddress; // each thread has one its value

        #endregion

        /// <summary>
        /// set values from a configuration file
        /// </summary>
        private void SetValuesFromConfigFile()
        {
            newsLetterFilePath = ConfigurationSettings.AppSettings["NewsLetterFilePath"].ToString();
            recipientEmailsFilePath = ConfigurationSettings.AppSettings["RecipientEmailsFilePath"].ToString();
            unsubscribedEmailsFilePath = ConfigurationSettings.AppSettings["UnsubscribedEmailsFilePath"].ToString();
            validSubscribedRecipientsFilePath = ConfigurationSettings.AppSettings["ValidSubscribedRecipientsFilePath"].ToString();
            logFilePath = ConfigurationSettings.AppSettings["LogFilePath"].ToString();
            smtpHost = ConfigurationSettings.AppSettings["SmtpHost"].ToString();
            senderEmailAddress = ConfigurationSettings.AppSettings["SenderEmailAddress"].ToString();
            mailSubject = ConfigurationSettings.AppSettings["MailSubject"].ToString();                                                            
        }

        /// <summary>
        /// get a news letter from a file
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private string GetNewsLetter(string filePath)
        {
            string newsLetter = string.Empty;

            if (File.Exists(filePath))
            {
                using (StreamReader streamReader = new StreamReader(filePath))
                {
                    newsLetter = streamReader.ReadToEnd();
                    newsLetter.Replace("images", "http://www.alacritytech.com.au/pup/contents/images");
                    streamReader.Close();
                }
            }            

            return newsLetter;
        }

        /// <summary>
        /// get unsubscribed email list from a file 
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private List<string> GetUnsubscribedEmails(string filePath)
        {
            List<string> unsubscribedEmails = new List<string>();

            if (File.Exists(filePath))
            {
                using (StreamReader streamReader = new StreamReader(filePath))
                {
                    string line = string.Empty;
                    while ((line = streamReader.ReadLine()) != null)
                    {
                        unsubscribedEmails.Add(line);
                    }
                    streamReader.Close();
                }
            }

            return unsubscribedEmails;
        }

        /// <summary>
        /// get valid and subscribed recipient email addresses from a file
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private List<string> GetValidSubscribedRecipients(string filePath)
        {
            List<string> recipients = new List<string>();

            if (File.Exists(filePath))
            {
                Console.WriteLine("Started reading an email list - checking validation and unsubscribed emails.");
                Console.WriteLine("...");

                using (StreamReader streamReader = new StreamReader(filePath))
                {
                    List<string> unsubscribedEmails = GetUnsubscribedEmails(unsubscribedEmailsFilePath);

                    string line = string.Empty;
                    while ((line = streamReader.ReadLine()) != null)
                    {
                        if (IsValidEmail(line)) // check if its a valid email 
                        {
                            if (IsUnsubscribedEmail(line, unsubscribedEmails) == false) // check if its a unsubscribed email
                            {
                                recipients.Add(line);
                            }                            
                        }                        
                    }
                    streamReader.Close();
                }

                Console.WriteLine("Finished reading the email list.");
            }            

            return recipients;
        }

        /// <summary>
        /// check if its a unsubscribed email from a unsubscribed email list
        /// </summary>
        /// <param name="emailAddress"></param>
        /// <param name="unsubscribedEmails"></param>
        /// <returns></returns>
        private bool IsUnsubscribedEmail(string emailAddress, List<string> unsubscribedEmails)
        {
            return unsubscribedEmails.Exists(v => v.Equals(emailAddress));
        }

        /// <summary>
        /// write valid and subscribed recipient email addresses to a file
        /// </summary>
        private void WriteValidSubscribedRecipients(List<string> recipients, string filePath)
        {
            using (StreamWriter streamWriter = new StreamWriter(filePath))
            {
                foreach (var recipient in recipients)
                {
                    streamWriter.WriteLine(recipient);                   
                }
                streamWriter.Close();
            }
        }

        /// <summary>
        /// check whether it's a valid email address
        /// </summary>
        /// <param name="emailAddress"></param>
        /// <returns></returns>
        private bool IsValidEmail(string emailAddress)
        {            
            string emailPattern = @"^(([^<>()[\]\\.,;:\s@\""]+"
                        + @"(\.[^<>()[\]\\.,;:\s@\""]+)*)|(\"".+\""))@"
                        + @"((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}"
                        + @"\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+"
                        + @"[a-zA-Z]{2,}))$";

            Regex regex = new Regex(emailPattern);

            return regex.IsMatch(emailAddress);                           
        }

        /// <summary>
        /// send bulk emails
        /// </summary>
        public void SendBulkEmails()
        {
            SetValuesFromConfigFile();
            logStreamer = new StreamWriter(logFilePath, true);
            List<string> recipients = GetValidSubscribedRecipients(recipientEmailsFilePath);
            string newsLetter = GetNewsLetter(newsLetterFilePath);

            WriteValidSubscribedRecipients(recipients, validSubscribedRecipientsFilePath); // TODO this function is really neccessary?            
            SendBulkEmails(recipients, newsLetter);            
        }

        /// <summary>
        /// send bulk emails 
        /// </summary>
        /// <param name="recipients">recipient list</param>
        /// <param name="newsLetter">news letter</param>
        private void SendBulkEmails(List<string> recipients, string newsLetter)
        {
            if (recipients.Count == 0)
            {
                Console.WriteLine("There is no any recipient.");
                return;
            }

            if (newsLetter == string.Empty)
            {
                Console.WriteLine("There is no news letter.");
                return;
            }                        
                                              
            Console.WriteLine("Started sending emails.");

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            totalRecipientCount = recipients.Count();                                              
            int recipientCountPerLoop = 1000;

            try
            {
                while (currRecipientIndex < totalRecipientCount)
                {
                    int startRecipientIndex = currRecipientIndex;
                    int endRecipientIndex = currRecipientIndex + recipientCountPerLoop;
                    if (endRecipientIndex > totalRecipientCount)
                    {
                        endRecipientIndex = totalRecipientCount;
                    }
                    recipientIndexForDoneEvent = endRecipientIndex;

                    STPStartInfo stpStartInfo = new STPStartInfo();
                    // stpStartInfo.StartSuspended = true;
                    stpStartInfo.AreThreadsBackground = false;
                    SmartThreadPool smartThreadPool = new SmartThreadPool(stpStartInfo);

                    doneEvent = new ManualResetEvent(false);
                    ParallelOptions parallelOptions = new ParallelOptions();
                    parallelOptions.MaxDegreeOfParallelism = 100;
                    Parallel.For(startRecipientIndex, endRecipientIndex, parallelOptions, (recipientIndex) =>
                    {
                        MailMessage mailMessage = new MailMessage(senderEmailAddress, recipients[recipientIndex]);
                        mailMessage.Subject = mailSubject;
                        mailMessage.Body = newsLetter;
                        mailMessage.IsBodyHtml = true;

                        // ThreadPool.QueueUserWorkItem(new WaitCallback(SendEmailAsync), mailMessage);                    
                        smartThreadPool.QueueWorkItem(new WorkItemCallback(this.SendEmailAsync), mailMessage, WorkItemPriority.Highest);
                    });

                    smartThreadPool.Start();

                    //// Wait for the completion of all work items
                    //smartThreadPool.WaitForIdle();
                    //smartThreadPool.Shutdown();

                    doneEvent.WaitOne(); // wait until all threads are completed                
                }
            }
            catch (Exception ex)
            {
                logStreamer.WriteLine(ex);
                Console.WriteLine(ex);
            }

            stopwatch.Stop();
            Console.WriteLine("Finished sending emails.");
            Console.WriteLine(String.Format("{0} of {1} emails were sent in {2}.", totalSentCount, totalRecipientCount, stopwatch.Elapsed.ToString()));
            Console.WriteLine("Please hit enter to exit application.");
            Console.ReadLine();            
        }        

        /// <summary>
        /// send an email in async mode
        /// </summary>
        /// <param name="mailMessage"></param>
        private object SendEmailAsync(Object mailMessageObject)
        {
            MailMessage mailMessage = mailMessageObject as MailMessage;
            if (mailMessage == null)
            {
                throw new ArgumentException("Argument must be of type MailMessage");
            }

            // create a smtp client 
            //SmtpClient smtpClient = new SmtpClient
            //{
            //    Host = "smtp.gmail.com",
            //    Port = 587,
            //    EnableSsl = true,
            //    UseDefaultCredentials = true,
            //    Credentials = new System.Net.NetworkCredential("", "")
            //};

            SmtpClient smtpClient = new SmtpClient
            {
                Host = smtpHost,
                UseDefaultCredentials = true
            };
                                    
            smtpClient.SendCompleted += new SendCompletedEventHandler(SmtpClient_SendCompleted);

            try
            {
                smtpClient.SendAsync(mailMessage, mailMessage.To.ToString());
                // smtpClient.Send(mailMessage);                
            }
            catch (Exception ex)
            {                
                logStreamer.WriteLine(ex);
                Console.WriteLine(ex);
            }

            return null;
        }
                
        private void SmtpClient_SendCompleted(object sender, AsyncCompletedEventArgs e)
        {            
            currRecipientIndex++;            

            string recipient = e.UserState.ToString();
            if (e.Cancelled == true)
            {
                Console.WriteLine(String.Format("{0}/{1} - Cancelled to sent  email to: {2} {3}", currRecipientIndex, totalRecipientCount, recipient, DateTime.Now));                       
            }
            else if (e.Error != null)
            {
                Console.WriteLine(String.Format("{0}/{1} - Failed to sent  email to: {2} {3}", currRecipientIndex, totalRecipientCount, recipient, DateTime.Now));                       
            }
            else
            {
                totalSentCount++;
                Console.WriteLine(String.Format("{0}/{1} - Successfully sent  email to: {2} {3}", currRecipientIndex, totalRecipientCount, recipient, DateTime.Now));                       
            }            
                                   
            // if (currRecipientIndex == totalRecipientCount)
            if (currRecipientIndex == recipientIndexForDoneEvent)
            {
                doneEvent.Set(); // set as all threads working has been completed                
            }            
        }        
    }
}
