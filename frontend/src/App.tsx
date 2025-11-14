import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { AppHeader } from './components/AppHeader';
import { ChatLayout } from './components/ChatLayout';

export const App = () => (
  <FluentProvider theme={webLightTheme} className="app-root">
    <AppHeader />
    <main className="app-main">
      <ChatLayout />
    </main>
  </FluentProvider>
);
