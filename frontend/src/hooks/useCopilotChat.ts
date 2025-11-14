import { useCallback, useMemo, useRef, useState } from 'react';
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
  copilotMessageId?: string;
}

interface CopilotConversationAttribution {
  providerDisplayName?: string;
  attributionSource?: string;
  seeMoreWebUrl?: string;
}

interface CopilotConversationResponseMessage {
  id?: string;
  text?: string;
  createdDateTime?: string;
  content?: string | CopilotContentItem[];
  attributions?: CopilotConversationAttribution[];
}

interface CopilotConversationResponse {
  id?: string;
  conversationId?: string;
  messages?: CopilotConversationResponseMessage[];
}

interface CopilotCreateConversationResponse {
  id?: string;
  conversationId?: string;
}

interface CopilotChatRequestBody {
  message: {
    text: string;
  };
  locationHint: {
    timeZone: string;
  };
}

const COPILOT_CITATION_REGEX = /\u{E200}cite\u{E202}([^\u{E201}]*)\u{E201}/gu;
const COPILOT_INLINE_TAG_REGEX = /<\/??[A-Z][A-Za-z0-9]*>/g;

const sanitizeCopilotText = (input: string) => {
  if (!input) {
    return '';
  }

  let output = input.replace(/\r\n/g, '\n');
  output = output.replace(/<br\s*\/?\s*>/gi, '\n');
  output = output.replace(COPILOT_INLINE_TAG_REGEX, '');
  output = output.replace(/[ \t]+\n/g, '\n');
  output = output.replace(/\n{3,}/g, '\n\n');

  return output.trim();
};

const buildSourcesAppendix = (
  attributions: CopilotConversationAttribution[] | undefined,
  requiredCount: number
) => {
  const uniqueSources: CopilotConversationAttribution[] = [];
  const seen = new Set<string>();

  attributions?.forEach((item) => {
    const title = (item?.providerDisplayName ?? '').trim();
    const link = (item?.seeMoreWebUrl ?? '').trim();
    const fingerprint = `${title.toLowerCase()}|${link.toLowerCase()}`;

    if (!title && !link) {
      return;
    }

    if (seen.has(fingerprint)) {
      return;
    }

    seen.add(fingerprint);
    uniqueSources.push({ providerDisplayName: title, seeMoreWebUrl: link });
  });

  const sourceCount = Math.max(requiredCount, uniqueSources.length);

  if (sourceCount === 0) {
    return '';
  }

  const listItems = Array.from({ length: sourceCount }, (_, index) => {
    const source = uniqueSources[index];
    const label = source?.providerDisplayName?.trim() || `Source ${index + 1}`;
    const href = source?.seeMoreWebUrl?.trim();
    const anchor = href
      ? `<a href="${href}" target="_blank" rel="noreferrer">${label}</a>`
      : label;

    return `<li id="copilot-source-${index + 1}">${anchor}</li>`;
  }).join('\n');

  return [
    '<details class="chat-message-sources">',
    '<summary>Sources</summary>',
    '<ol>',
    listItems,
    '</ol>',
    '</details>',
  ].join('\n');
};

const formatCopilotResponseText = (
  rawText: string,
  attributions?: CopilotConversationAttribution[]
) => {
  let citationCount = 0;
  const citationTokenToIndex = new Map<string, number>();

  const withCitations = rawText.replace(COPILOT_CITATION_REGEX, (_match, token) => {
    const normalizedToken = (token ?? '').trim();
    if (!citationTokenToIndex.has(normalizedToken)) {
      citationCount += 1;
      citationTokenToIndex.set(normalizedToken, citationCount);
    }

    const citationIndex = citationTokenToIndex.get(normalizedToken)!;
    return `<sup><a class="chat-citation" href="#copilot-source-${citationIndex}">[${citationIndex}]</a></sup>`;
  });

  const cleaned = sanitizeCopilotText(withCitations);
  if (!cleaned) {
    return '';
  }

  const sourcesAppendix = buildSourcesAppendix(attributions, citationCount);
  return sourcesAppendix ? `${cleaned}\n\n${sourcesAppendix}` : cleaned;
};

export const useCopilotChat = () => {
  const { acquireToken } = useAuth();
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [isSending, setIsSending] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [conversationId, setConversationId] = useState<string | null>(null);
  // Track server-side message IDs and locally queued user prompts so we can reconcile responses.
  const processedResponseIds = useRef<Set<string>>(new Set());
  const pendingUserMessages = useRef<Array<{ localId: string; content: string }>>([]);

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
      pendingUserMessages.current.push({ localId: userMessage.id, content });
      setIsSending(true);
      setError(null);

      try {
        const rawEndpoint =
          (import.meta.env.VITE_COPILOT_ENDPOINT as string | undefined)?.trim() ||
          'https://graph.microsoft.com/beta/copilot';
        const copilotEndpoint = rawEndpoint.replace(/\/+$/, '');
        const subscriptionKey = (import.meta.env.VITE_COPILOT_SUBSCRIPTION_KEY as string | undefined)?.trim();

        const accessToken = await acquireToken();
        if (!accessToken) {
          throw new Error('Failed to acquire access token.');
        }

        let activeConversationId = conversationId;
        if (!activeConversationId) {
          const createResponse = await fetch(`${copilotEndpoint}/conversations`, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              Authorization: `Bearer ${accessToken}`,
              ...(subscriptionKey ? { 'Ocp-Apim-Subscription-Key': subscriptionKey } : {}),
            },
            body: JSON.stringify({}),
          });

          if (!createResponse.ok) {
            let errorDetail: string;
            try {
              const errorJson = await createResponse.json();
              errorDetail = JSON.stringify(errorJson);
            } catch (parseError) {
              errorDetail = await createResponse.text();
            }
            throw new Error(
              `Copilot conversation creation failed (${createResponse.status} ${createResponse.statusText}): ${errorDetail}`
            );
          }

          const createData: CopilotCreateConversationResponse = await createResponse.json();
          activeConversationId = createData.id ?? createData.conversationId ?? null;
          if (!activeConversationId) {
            throw new Error('Copilot conversation creation succeeded but no conversation ID was returned.');
          }
          setConversationId(activeConversationId);
        }

        const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
        const chatBody: CopilotChatRequestBody = {
          message: {
            text: content,
          },
          locationHint: {
            timeZone,
          },
        };

        const chatResponse = await fetch(`${copilotEndpoint}/conversations/${activeConversationId}/chat`, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            Authorization: `Bearer ${accessToken}`,
            ...(subscriptionKey ? { 'Ocp-Apim-Subscription-Key': subscriptionKey } : {}),
          },
          body: JSON.stringify(chatBody),
        });

        if (!chatResponse.ok) {
          let errorDetail: string;
          try {
            const errorJson = await chatResponse.json();
            errorDetail = JSON.stringify(errorJson);
          } catch (parseError) {
            errorDetail = await chatResponse.text();
          }
          throw new Error(
            `Copilot chat request failed (${chatResponse.status} ${chatResponse.statusText}): ${errorDetail}`
          );
        }

        const data: CopilotConversationResponse = await chatResponse.json();
        const nextConversationId = data.id ?? data.conversationId ?? activeConversationId;
        if (nextConversationId && nextConversationId !== conversationId) {
          setConversationId(nextConversationId);
        }

        const responseMessages = Array.isArray(data.messages) ? data.messages : [];
        if (responseMessages.length === 0) {
          return;
        }

        const newAssistantMessages: ChatMessage[] = [];
        const userMessageMatches: Array<{ localId: string; copilotId: string }> = [];

        for (const responseMessage of responseMessages) {
          const responseId = responseMessage?.id;
          if (!responseId || processedResponseIds.current.has(responseId)) {
            continue;
          }

          const rawText = (() => {
            if (typeof responseMessage?.text === 'string' && responseMessage.text.trim()) {
              return responseMessage.text;
            }
            if (typeof responseMessage?.content === 'string') {
              return responseMessage.content;
            }
            if (Array.isArray(responseMessage?.content)) {
              return responseMessage.content
                .map((item) => item?.text ?? item?.value ?? item?.content ?? '')
                .filter(Boolean)
                .join('\n');
            }
            return '';
          })();

          const normalizedContent = formatCopilotResponseText(rawText, responseMessage?.attributions);

          processedResponseIds.current.add(responseId);

          if (!normalizedContent) {
            continue;
          }

          const pendingIndex = pendingUserMessages.current.findIndex(
            (pending) => pending.content.trim() === normalizedContent
          );

          if (pendingIndex !== -1) {
            const [pending] = pendingUserMessages.current.splice(pendingIndex, 1);
            userMessageMatches.push({ localId: pending.localId, copilotId: responseId });
            continue;
          }

          newAssistantMessages.push({
            id: `assistant-${responseId}`,
            role: 'assistant',
            content: normalizedContent,
            createdAt: responseMessage?.createdDateTime ?? new Date().toISOString(),
            copilotMessageId: responseId,
          });
        }

        if (userMessageMatches.length > 0 || newAssistantMessages.length > 0) {
          const userMatchMap = new Map(userMessageMatches.map((item) => [item.localId, item.copilotId]));
          setMessages((current) => {
            let updated = current;
            if (userMessageMatches.length > 0) {
              updated = updated.map((message) =>
                userMatchMap.has(message.id)
                  ? { ...message, copilotMessageId: userMatchMap.get(message.id) }
                  : message
              );
            }
            if (newAssistantMessages.length > 0) {
              updated = [...updated, ...newAssistantMessages];
            }
            return updated;
          });
        }
      } catch (err) {
        pendingUserMessages.current = pendingUserMessages.current.filter(
          (pending) => pending.localId !== userMessage.id
        );
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
    processedResponseIds.current.clear();
    pendingUserMessages.current = [];
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
