using System;

namespace ocli
{
    public enum MailType { Mail, Meeting };

    public class InboxItem
    {
        public MailType ItemType { get; private set; }
        public string ConversationId { get; private set; }
        public string ConversationIndex { get; private set; }
        public string Sender { get; private set; }
        public string Subject { get; private set; }
        public DateTime Received { get; private set; }

        public InboxItem(MailType itemType, string conversationId, string conversationIndex, string sender, string subject, DateTime received)
        {
            ItemType = ItemType;
            ConversationId = conversationId;
            ConversationIndex = conversationIndex;
            Sender = sender;
            Subject = subject;
            Received = received;
        }
    }
}
