using System;
using System.Collections.Generic;
using System.Linq;

namespace MsgReader.Entities
{
    public class FlagBag
    {
        public string Request { get; internal set; }

        public Outlook.Storage.Flag.FlagStatus? Status { get; internal set; }
    }
}