import { ChangeEvent, KeyboardEvent, useState } from 'react';
import { Button, Input } from '@fluentui/react-components';
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

  const handleChange = (event: ChangeEvent<HTMLInputElement>) => {
    setValue(event.target.value);
  };

  const handleKeyDown = (event: KeyboardEvent<HTMLInputElement>) => {
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault();
      void handleSubmit();
    }
  };

  return (
    <Input
      value={value}
      onChange={handleChange}
      onKeyDown={handleKeyDown}
      placeholder="Ask Microsoft 365 Copilot…"
      contentAfter={
        <Button
          appearance="primary"
          icon={<Send24Regular />}
          disabled={disabled || !value.trim()}
          onClick={() => void handleSubmit()}
        />
      }
      disabled={disabled}
    />
  );
};
