using System;
using System.Collections.Generic;
using System.Linq;

namespace MsgReader.Entities
{
    public class ReceivedByBag
    {
        public string AddressType { get; internal set; }

        public string DisplayName { get; internal set; }

        public string Email { get; internal set; }
    }
}