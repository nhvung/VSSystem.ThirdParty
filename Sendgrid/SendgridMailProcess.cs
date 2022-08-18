using Newtonsoft.Json;
using SendGrid;
using SendGrid.Helpers.Mail;
using System;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using VSSystem.Net.Mail;

namespace VSSystem.ThirdParty.Sendgrid
{
    public class SendgridMailProcess : IMailProcess
    {
        SendGridClient _sgClient;
        public SendgridMailProcess(string host, string apiKey)
        {
            SendGrid.SendGridClientOptions sgOpts = new SendGrid.SendGridClientOptions();
            sgOpts.Host = host;
            sgOpts.ApiKey = apiKey;
            sgOpts.UrlPath = host;
            _sgClient = new SendGridClient(apiKey);
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
        }
        public EMailProcessType Type { get { return EMailProcessType.SendgridApi; } }

        Action<SendMailReponse> _OnSendComplete;
        public Action<SendMailReponse> OnSendComplete { set { _OnSendComplete = value; } }

        async public Task<SendMailReponse> SendMailAsync(MailMessageInfo mailInfo, int retryTimes = 0)
        {
            SendMailReponse result = new SendMailReponse();
            result.Mail_ID = mailInfo.Mail_ID;
            try
            {
                if (string.IsNullOrEmpty(mailInfo.FromAddress))
                {
                    throw new Exception("From Address is empty.");
                }
                if (mailInfo.ToAddresses?.Count > 0)
                {
                    DateTime bTime = DateTime.Now;
                    SendGridMessage mMess = new SendGridMessage();

                    if (!string.IsNullOrEmpty(mailInfo.Subject))
                    {
                        mMess.Subject = mailInfo.Subject;
                    }

                    if (mailInfo.ToAddresses.Count == 0)
                    {
                        result.StatusCode = (int)System.Net.Mail.SmtpStatusCode.MailboxUnavailable;
                        result.Message = "Email is empty.";
                        return result;
                    }

                    mMess.From = new EmailAddress(mailInfo.FromAddress, mailInfo.DisplayName);
                    foreach (string toAddress in mailInfo.ToAddresses)
                    {
                        if (!string.IsNullOrEmpty(toAddress))
                        {
                            try { mMess.AddTo(toAddress); }
                            catch { }
                        }
                    }

                    if (mailInfo.CcAddresses?.Count > 0)
                    {
                        foreach (string ccAddress in mailInfo.CcAddresses)
                        {
                            if (!string.IsNullOrEmpty(ccAddress))
                            {
                                try { mMess.AddCc(ccAddress); }
                                catch { }
                            }
                        }
                    }


                    if (mailInfo.BccAddresses?.Count > 0)
                    {
                        foreach (string bccAddress in mailInfo.BccAddresses)
                        {
                            if (!string.IsNullOrEmpty(bccAddress))
                            {
                                try { mMess.AddBcc(bccAddress); }
                                catch { }
                            }
                        }
                    }

                    if (mailInfo.AlternateViewItem != null)
                    {
                        mMess.AddContent(mailInfo.AlternateViewItem.MediaType, Encoding.UTF8.GetString(mailInfo.AlternateViewItem.ContentBytes));

                        if (mailInfo.LinkResourceItems?.Count > 0)
                        {
                            foreach (var lrItem in mailInfo.LinkResourceItems)
                            {
                                Attachment attachmentObj = new Attachment();
                                attachmentObj.ContentId = lrItem.ContentID;
                                attachmentObj.Type = lrItem.MediaType;
                                attachmentObj.Content = Convert.ToBase64String(lrItem.ContentBytes);
                                attachmentObj.Filename = lrItem.ContentID;
                                attachmentObj.Disposition = "inline";
                                mMess.AddAttachment(attachmentObj);
                            }
                        }
                    }
                    if (mailInfo.AttachmentItems?.Count > 0)
                    {
                        foreach (var attItem in mailInfo.AttachmentItems)
                        {
                            Attachment attachmentObj = new Attachment();
                            attachmentObj.ContentId = attItem.ContentID;
                            attachmentObj.Filename = attachmentObj.ContentId;
                            attachmentObj.Type = attItem.MediaType;
                            attachmentObj.Content = Convert.ToBase64String(attItem.ContentBytes);
                            mMess.AddAttachment(attachmentObj);
                        }
                    }

                    int retry = 0;
                RETRY:
                    try
                    {



                        var sendResult = await _sgClient.SendEmailAsync(mMess);
                        result.StatusCode = (int)sendResult.StatusCode;

                        if (sendResult.IsSuccessStatusCode)
                        {
                            if (sendResult.Headers != null)
                            {
                                foreach (var header in sendResult.Headers)
                                {
                                    if (header.Key?.Equals("X-Message-Id", StringComparison.InvariantCultureIgnoreCase) ?? false)
                                    {
                                        result.MessageID = header.Value.FirstOrDefault();
                                    }
                                }
                            }
                        }
                        else
                        {
                            result.Message = sendResult.Body.ReadAsStringAsync().Result;
                        }
                    }
                    catch (Exception ex)
                    {
                        retry++;
                        if (retry < retryTimes)
                        {
                            Thread.Sleep(5000);
                            goto RETRY;
                        }
                        result.Message = ex.Message;
                        result.StatusCode = (int)System.Net.Mail.SmtpStatusCode.MailboxUnavailable;

                    }

                    DateTime eTime = DateTime.Now;
                    result.SentTime = eTime - bTime;

                    if (_OnSendComplete != null)
                    {
                        var extInfo = new
                        {
                            mailInfo.ToAddresses,
                            mailInfo.CcAddresses,
                            mailInfo.BccAddresses,
                            mailInfo.Subject,
                        };
                        result.ExtendInfo = JsonConvert.SerializeObject(extInfo);
                        _ = Task.Run(() => _OnSendComplete.Invoke(result));
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return result;
        }
    }
}
