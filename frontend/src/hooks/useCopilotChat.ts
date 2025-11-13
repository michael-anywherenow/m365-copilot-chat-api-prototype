import { useCallback, useMemo, useState } from 'react';
import { useAuth } from '../context/AuthProvider';

type ChatRole = 'user' | 'assistant' | 'system';

export interface ChatMessage {
  id: string;
  role: ChatRole;
  content: string;
  createdAt: string;
}

interface CopilotResponseMessage {
  role: ChatRole;
  content: string;
}

interface CopilotResponse {
  id?: string;
  messages: CopilotResponseMessage[];
  choices?: { message?: CopilotResponseMessage }[];
  conversationId?: string;
}

export const useCopilotChat = () => {
  const { acquireToken } = useAuth();
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [isSending, setIsSending] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [conversationId, setConversationId] = useState<string | null>(null);

  const sendMessage = useCallback(
    async (content: string) => {
      if (!content.trim()) {
        return;
      }

      const timestamp = new Date().toISOString();
      const userMessage: ChatMessage = {
        id: `user-${timestamp}`,
        role: 'user',
        content,
        createdAt: timestamp,
      };

      setMessages((current) => [...current, userMessage]);
      setIsSending(true);
      setError(null);

      try {
        const copilotEndpoint =
          (import.meta.env.VITE_COPILOT_ENDPOINT as string | undefined)?.trim() ||
          'https://graph.microsoft.com/v1.0/ai/copilot/chatCompletions';
        const subscriptionKey = (import.meta.env.VITE_COPILOT_SUBSCRIPTION_KEY as string | undefined)?.trim();

        const accessToken = await acquireToken();
        if (!accessToken) {
          throw new Error('Failed to acquire access token.');
        }

        const response = await fetch(copilotEndpoint, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            Authorization: `Bearer ${accessToken}`,
            ...(subscriptionKey ? { 'Ocp-Apim-Subscription-Key': subscriptionKey } : {}),
          },
          body: JSON.stringify({
            conversationId: conversationId ?? undefined,
            messages: [...messages, userMessage].map((message) => ({
              role: message.role,
              content: message.content,
            })),
          }),
        });

        if (!response.ok) {
          let errorDetail: string;
          try {
            const errorJson = await response.json();
            errorDetail = JSON.stringify(errorJson);
          } catch (parseError) {
            errorDetail = await response.text();
          }
          throw new Error(
            `Copilot chat request failed (${response.status} ${response.statusText}): ${errorDetail}`
          );
        }

        const data: CopilotResponse = await response.json();
        if (data.id) {
          setConversationId(data.id);
        }
        if ((data as { conversationId?: string }).conversationId) {
          setConversationId((data as { conversationId?: string }).conversationId ?? null);
        }

        const copilotMessages: CopilotResponseMessage[] = Array.isArray(data.messages)
          ? data.messages
          : Array.isArray((data as { choices?: { message?: CopilotResponseMessage }[] }).choices)
          ? ((data as { choices?: { message?: CopilotResponseMessage }[] }).choices ?? [])
              .map((choice) => choice.message)
              .filter((message): message is CopilotResponseMessage => Boolean(message))
          : [];

        const assistantMessages = copilotMessages.map((message, index) => ({
          id: `${data.id ?? 'assistant'}-${index}-${Date.now()}`,
          role: message.role,
          content: message.content,
          createdAt: new Date().toISOString(),
        }));

        setMessages((current) => [...current, ...assistantMessages]);
      } catch (err) {
        const message =
          err instanceof Error ? err.message : 'Unexpected error contacting Copilot.';
        setError(message);
      } finally {
        setIsSending(false);
      }
    },
    [acquireToken, messages, conversationId]
  );

  const resetConversation = useCallback(() => {
    setMessages([]);
    setError(null);
    setConversationId(null);
  }, []);

  return useMemo(
    () => ({
      messages,
      isSending,
      error,
      sendMessage,
      resetConversation,
      conversationId,
    }),
    [messages, isSending, error, sendMessage, resetConversation, conversationId]
  );
};
