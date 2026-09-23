import {AppShell} from '@astryxdesign/core/AppShell';
import {Heading} from '@astryxdesign/core/Heading';
import {TopNav} from '@astryxdesign/core/TopNav';

export function App() {
  return (
    <AppShell
      topNav={<TopNav heading={<Heading level={1}>pptx-tool</Heading>} />}
      contentPadding={4}
    >
      <p>Coming soon.</p>
    </AppShell>
  );
}
