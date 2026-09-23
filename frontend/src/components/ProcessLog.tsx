import {CodeBlock} from '@astryxdesign/core/CodeBlock';
import {Collapsible} from '@astryxdesign/core/Collapsible';

export interface ProcessLogProps {
  log: string;
}

export function ProcessLog({log}: ProcessLogProps) {
  return (
    <Collapsible trigger="Processing log">
      <CodeBlock code={log.trimEnd()} language="plaintext" width="100%" size="sm" />
    </Collapsible>
  );
}
