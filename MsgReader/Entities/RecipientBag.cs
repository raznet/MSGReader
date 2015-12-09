using System;
using System.Collections.Generic;
using System.Linq;

namespace MsgReader.Entities
{
    public class RecipientBag
    {
        public string DisplayName { get; internal set; }

        public string Email { get; internal set; }

        public Outlook.Storage.Recipient.RecipientType? Type { get; internal set; }
    }
}