using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MSG = MsgReader.Outlook.Storage;

namespace MsgReader.Entities
{
    public class MessageBag
    {
        public MSG.Appointment Appointment { get; private set; }

        public IList<AttachmentBag> Attachments { get; private set; }

        public string Body { get; private set; }

        public DateTime? Created { get; private set; }

        public FlagBag Flag { get; private set; }

        public string HtmlBody { get; private set; }

        public MSG.Message.MessageImportance? Importance { get; private set; }

        public string ImportanceText { get; private set; }

        public Encoding InternetCodePage { get; private set; }

        public Encoding MessageCodePage { get; private set; }

        public string MessageType { get; private set; }

        public DateTime? Modified { get; private set; }

        public string ModifiedBy { get; private set; }

        public DateTime? Received { get; private set; }

        public List<RecipientBag> Recipients { get; private set; }

        public string RtfBody { get; private set; }

        public MSG.Sender Sender { get; private set; }

        public MSG.SenderRepresenting SenderRepresenting { get; private set; }

        public DateTime? Sent { get; private set; }

        public string SignedBy { get; private set; }

        public string Subject { get; private set; }

        public ReceivedByBag ReceivedBy { get; private set; } //receivedby

        internal static MessageBag Create(MSG.Message msg)
        {
            var result = new MessageBag();
            result.Appointment = msg.Appointment;
            result.Attachments = GetAttachments(msg);
            result.Recipients = GetRecipients(msg);
            if (msg.Flag != null)
            {
                result.Flag = new FlagBag();
                result.Flag.Request = msg.Flag.Request;
                result.Flag.Status = msg.Flag.Status;
            }

            result.Body = msg.BodyText;
            result.Created = msg.CreationTime;
            result.HtmlBody = msg.BodyHtml;
            result.Importance = msg.Importance;
            result.ImportanceText = msg.ImportanceText;
            result.InternetCodePage = msg.InternetCodePage;
            result.MessageCodePage = msg.MessageCodePage;
            result.MessageType = msg.Type.ToString();
            result.Modified = msg.LastModificationTime;
            result.ModifiedBy = msg.LastModifierName;
            result.Received = msg.ReceivedOn;
            result.RtfBody = msg.BodyRtf;
            result.Sender = msg.Sender;
            result.SenderRepresenting = msg.SenderRepresenting;
            result.Sent = msg.SentOn;
            result.SignedBy = msg.SignedBy;
            result.Subject = msg.Subject;

            if (msg.ReceivedBy != null)
            {
                var toRecipient = new ReceivedByBag();
                toRecipient.AddressType = msg.ReceivedBy.AddressType;
                toRecipient.DisplayName = msg.ReceivedBy.DisplayName;
                toRecipient.Email = msg.ReceivedBy.Email;
                result.ReceivedBy = toRecipient;
            }
            return result;
        }

        private static List<AttachmentBag> GetAttachments(MSG.Message msg)
        {
            var result = new List<AttachmentBag>();
            if (msg.Attachments != null)
            {
                foreach (object item in msg.Attachments)
                {
                    var attachment = item as MSG.Attachment;
                    if (attachment != null)
                    {
                        var att = new AttachmentBag();
                        att.AttachmentName = attachment.FileName;
                        result.Add(att);
                    }
                }
            }

            return result;
        }

        private static List<RecipientBag> GetRecipients(MSG.Message msg)
        {
            var result = new List<RecipientBag>();
            if (msg.Attachments != null)
            {
                foreach (var item in msg.Recipients)
                {
                    var recipient = new RecipientBag();
                    recipient.DisplayName = item.DisplayName;
                    recipient.Email = item.Email;
                    recipient.Type = item.Type;
                    result.Add(recipient);
                }
            }

            return result;
        }
    }
}