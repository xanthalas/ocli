using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ocli
{
    public class MailIdentifier
    {
        public int Id;
        public string ConversationId;
        public string ConversationIndex;

        public MailIdentifier(int id, string conversationId, string conversationIndex)
        {
            Id = id;
            ConversationId = conversationId;
            ConversationIndex = conversationIndex;
        }
    }
}
