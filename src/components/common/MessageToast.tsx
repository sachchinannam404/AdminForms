import * as React from 'react';
import { MessageBar, MessageBarType } from '@fluentui/react';

export interface IToastMessage {
  id: string;
  text: string;
  type: MessageBarType;
}

interface IMessageToastContext {
  messages: IToastMessage[];
  show: (text: string, type?: MessageBarType) => void;
  dismiss: (id: string) => void;
  clear: () => void;
}

const MessageToastContext = React.createContext<IMessageToastContext | null>(null);

let idCounter = 0;

export const MessageToastProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [messages, setMessages] = React.useState<IToastMessage[]>([]);

  const dismiss = React.useCallback((id: string) => {
    setMessages((prev) => prev.filter((m) => m.id !== id));
  }, []);

  const show = React.useCallback(
    (text: string, type: MessageBarType = MessageBarType.info) => {
      const id = `toast-${++idCounter}`;
      setMessages((prev) => [...prev, { id, text, type }]);
      window.setTimeout(() => dismiss(id), 6000);
    },
    [dismiss]
  );

  const clear = React.useCallback(() => setMessages([]), []);

  return (
    <MessageToastContext.Provider value={{ messages, show, dismiss, clear }}>
      <div style={{ position: 'sticky', top: 0, zIndex: 100, padding: messages.length ? 8 : 0 }}>
        {messages.map((m) => (
          <MessageBar
            key={m.id}
            messageBarType={m.type}
            onDismiss={() => dismiss(m.id)}
            styles={{ root: { marginBottom: 4 } }}
          >
            {m.text}
          </MessageBar>
        ))}
      </div>
      {children}
    </MessageToastContext.Provider>
  );
};

export function useMessageToast(): IMessageToastContext {
  const ctx = React.useContext(MessageToastContext);
  if (!ctx) {
    return {
      messages: [],
      show: (text, type) => console.info('[toast]', type, text),
      dismiss: () => undefined,
      clear: () => undefined
    };
  }
  return ctx;
}
