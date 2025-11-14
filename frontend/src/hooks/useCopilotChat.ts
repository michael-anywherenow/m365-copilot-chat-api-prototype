import { useCallback, useMemo, useState } from 'react';
import { useAuth } from '../context/AuthProvider';

type ChatRole = 'user' | 'assistant' | 'system';

interface CopilotContentItem {
  type: string;
  text?: string;
  value?: string;
  content?: string;
}

export interface ChatMessage {
  id: string;
  role: ChatRole;
  content: string;
  createdAt: string;
}

interface CopilotResponseMessage {
  role: ChatRole;
  content: string | CopilotContentItem[];
}

interface CopilotResponseChoice {
  message?: CopilotResponseMessage;
}

interface CopilotResponse {
  id?: string;
  messages?: CopilotResponseMessage[];
  choices?: CopilotResponseChoice[];
  conversationId?: string;
}

interface CopilotRequestMessage {
  role: ChatRole;
  content: CopilotContentItem[];
}

interface CopilotRequestBody {
  messages: CopilotRequestMessage[];
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
        const rawEndpoint =
          (import.meta.env.VITE_COPILOT_ENDPOINT as string | undefined)?.trim() ||
          'https://graph.microsoft.com/v1.0/copilot';
        const copilotEndpoint = rawEndpoint.replace(/\/+$/, '');
        const subscriptionKey = (import.meta.env.VITE_COPILOT_SUBSCRIPTION_KEY as string | undefined)?.trim();

        const accessToken = await acquireToken();
        if (!accessToken) {
          throw new Error('Failed to acquire access token.');
        }

        const requestMessage: CopilotRequestMessage = {
          role: 'user',
          content: [
            {
              type: 'text',
              text: content,
            },
          ],
        };

        const targetUrl = !conversationId
          ? `${copilotEndpoint}/conversations`
          : `${copilotEndpoint}/conversations/${conversationId}/chat`;

        const requestBody: CopilotRequestBody = {
          messages: [requestMessage],
        };

        const response = await fetch(targetUrl, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            Authorization: `Bearer ${accessToken}`,
            ...(subscriptionKey ? { 'Ocp-Apim-Subscription-Key': subscriptionKey } : {}),
          },
          body: JSON.stringify(requestBody),
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
        const nextConversationId = data.id ?? data.conversationId ?? conversationId;
        if (nextConversationId) {
          setConversationId(nextConversationId);
        }

        const responseMessages: CopilotResponseMessage[] | undefined = Array.isArray(data.messages)
          ? data.messages
          : Array.isArray(data.choices)
          ? (data.choices ?? [])
              .map((choice) => choice?.message)
              .filter((message): message is CopilotResponseMessage => Boolean(message))
          : undefined;

        if (!responseMessages || responseMessages.length === 0) {
          return;
        }

        const assistantMessages = responseMessages.map((message, index) => {
          const normalizedContent = Array.isArray(message.content)
            ? message.content
                .map((item) => item.text ?? item.value ?? item.content ?? '')
                .filter(Boolean)
                .join('\n')
            : message.content;

          return {
            id: `${nextConversationId ?? 'assistant'}-${index}-${Date.now()}`,
            role: message.role,
            content: normalizedContent,
            createdAt: new Date().toISOString(),
          };
        });

        setMessages((current) => [...current, ...assistantMessages]);
      } catch (err) {
        const message =
          err instanceof Error ? err.message : 'Unexpected error contacting Copilot.';
        setError(message);
      } finally {
        setIsSending(false);
      }
    },
    [acquireToken, conversationId]
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
