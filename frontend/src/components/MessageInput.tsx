import { ChangeEvent, KeyboardEvent, useState } from 'react';
import { Button, Textarea } from '@fluentui/react-components';
import { Send24Regular } from '@fluentui/react-icons';

interface MessageInputProps {
  onSend: (content: string) => Promise<void> | void;
  disabled?: boolean;
}

export const MessageInput = ({ onSend, disabled = false }: MessageInputProps) => {
  const [value, setValue] = useState('');

  const handleSubmit = async () => {
    if (!value.trim()) {
      return;
    }
    await onSend(value);
    setValue('');
  };

  const handleChange = (event: ChangeEvent<HTMLTextAreaElement>) => {
    setValue(event.target.value);
  };

  const handleKeyDown = (event: KeyboardEvent<HTMLTextAreaElement>) => {
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault();
      void handleSubmit();
    }
  };

  return (
    <div className="message-input-container">
      <Textarea
        value={value}
        onChange={handleChange}
        onKeyDown={handleKeyDown}
        placeholder="Ask Microsoft 365 Copilot…"
        rows={3}
        resize="vertical"
        disabled={disabled}
      />
      <Button
        appearance="primary"
        icon={<Send24Regular />}
        disabled={disabled || !value.trim()}
        onClick={() => void handleSubmit()}
      />
    </div>
  );
};
