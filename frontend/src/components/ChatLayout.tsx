import { Button, Card, Spinner, Text } from '@fluentui/react-components';
import { DismissRegular } from '@fluentui/react-icons';
import { useCopilotChat } from '../hooks/useCopilotChat';
import { MessageList } from './MessageList';
import { MessageInput } from './MessageInput';

export const ChatLayout = () => {
  const { messages, isSending, error, sendMessage, resetConversation, conversationId } =
    useCopilotChat();

  return (
    <Card appearance="outline" className="chat-card">
      <div className="chat-header">
        <Text as="h2" weight="semibold">
          Microsoft 365 Copilot
        </Text>
        <Button
          size="small"
          appearance="transparent"
          icon={<DismissRegular />}
          onClick={resetConversation}
        >
          Reset
        </Button>
      </div>

      {conversationId ? (
        <Text size={200} className="chat-conversation-id">
          Conversation ID: {conversationId}
        </Text>
      ) : null}

      <MessageList messages={messages} isLoading={isSending} />

      {error ? (
        <Text role="alert" className="chat-error">
          {error}
        </Text>
      ) : null}

      <div className="chat-input">
        <MessageInput disabled={isSending} onSend={sendMessage} />
        {isSending ? <Spinner size="extra-tiny" /> : null}
      </div>
    </Card>
  );
};
