using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MsgReader.Entities;
using MsgReader.Outlook;

namespace MsgReader
{
    public class BagReader : Reader
    {
        public MessageBag CreateMessageBag(string inputFile, DirectoryInfo target)
        {
            MessageBag result;
            using (var stream = File.Open(inputFile, FileMode.Open, FileAccess.Read))
            using (var message = new Storage.Message(stream))
            {
                switch (message.Type)
                {
                    case Storage.Message.MessageType.Email:
                    case Storage.Message.MessageType.EmailSms:
                    case Storage.Message.MessageType.EmailNonDeliveryReport:
                    case Storage.Message.MessageType.EmailDeliveryReport:
                    case Storage.Message.MessageType.EmailDelayedDeliveryReport:
                    case Storage.Message.MessageType.EmailReadReceipt:
                    case Storage.Message.MessageType.EmailNonReadReceipt:
                    case Storage.Message.MessageType.EmailEncryptedAndMaybeSigned:
                    case Storage.Message.MessageType.EmailEncryptedAndMaybeSignedNonDelivery:
                    case Storage.Message.MessageType.EmailEncryptedAndMaybeSignedDelivery:
                    case Storage.Message.MessageType.EmailClearSignedReadReceipt:
                    case Storage.Message.MessageType.EmailClearSignedNonDelivery:
                    case Storage.Message.MessageType.EmailClearSignedDelivery:
                    case Storage.Message.MessageType.EmailBmaStub:
                    case Storage.Message.MessageType.CiscoUnityVoiceMessage:
                    case Storage.Message.MessageType.EmailClearSigned:
                        var fileName = "email";
                        bool htmlBody;
                        string body;
                        string dummy;
                        List<string> attachmentList;
                        List<string> files;

                        PreProcessMsgFile(message,
                            false,
                            target.FullName,
                            ref fileName,
                            out htmlBody,
                            out body,
                            out dummy,
                            out attachmentList,
                            out files);
                        result = MessageBag.Create(message);
                        result.PreviewBody = body;
                        break;

                    default:
                        result = null;
                        break;
                }
            }

            return result;
        }
    }
}