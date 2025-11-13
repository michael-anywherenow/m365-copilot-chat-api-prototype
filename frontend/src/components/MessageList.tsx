import { Avatar, Caption1, Card, Text } from '@fluentui/react-components';
import { ChatMessage } from '../hooks/useCopilotChat';

interface MessageListProps {
  messages: ChatMessage[];
  isLoading: boolean;
}

const roleToDisplay = (role: ChatMessage['role']) => {
  switch (role) {
    case 'assistant':
      return 'Copilot';
    case 'system':
      return 'System';
    default:
      return 'You';
  }
};

export const MessageList = ({ messages, isLoading }: MessageListProps) => (
  <div className="chat-messages">
    {messages.map((message) => (
      <Card
        key={message.id}
        appearance={message.role === 'user' ? 'filled' : 'outline'}
        className={`chat-message chat-message-${message.role}`}
      >
        <div className="chat-message-header">
          <Avatar
            aria-hidden
            size={28}
            name={roleToDisplay(message.role)}
            color={message.role === 'user' ? 'brand' : 'colorful'}
          />
          <div>
            <Text weight="semibold">{roleToDisplay(message.role)}</Text>
            <Caption1>{new Date(message.createdAt).toLocaleTimeString()}</Caption1>
          </div>
        </div>
        <Text>{message.content}</Text>
      </Card>
    ))}
    {isLoading ? (
      <Card appearance="subtle" className="chat-message chat-message-assistant">
        <Text>Thinking…</Text>
      </Card>
    ) : null}
  </div>
);
