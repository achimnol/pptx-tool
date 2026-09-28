import {BottomSheet} from '@astryxdesign/core/BottomSheet';
import {Button} from '@astryxdesign/core/Button';
import {CodeBlock} from '@astryxdesign/core/CodeBlock';
import {Heading} from '@astryxdesign/core/Heading';
import {HStack} from '@astryxdesign/core/HStack';
import {Icon} from '@astryxdesign/core/Icon';
import {IconButton} from '@astryxdesign/core/IconButton';
import {VStack} from '@astryxdesign/core/VStack';
import {useState} from 'react';

export interface ProcessLogProps {
  log: string;
}

export function ProcessLog({log}: ProcessLogProps) {
  const [isOpen, setIsOpen] = useState(false);
  return (
    <>
      <Button
        label="Processing log"
        variant="secondary"
        aria-expanded={isOpen}
        onClick={() => setIsOpen((open) => !open)}
      />
      {/* A non-modal panel docked at the bottom keeps the page usable while reading the log. */}
      <BottomSheet
        className="process-log-sheet"
        label="Processing log"
        isOpen={isOpen}
        onOpenChange={setIsOpen}
        hasScrim={false}
        height="capped"
      >
        <VStack gap={3} padding={4}>
          <HStack gap={2} align="center" justify="between">
            <Heading level={3}>Processing log</Heading>
            <IconButton
              label="Close the processing log"
              icon={<Icon icon="close" />}
              variant="ghost"
              size="sm"
              onClick={() => setIsOpen(false)}
            />
          </HStack>
          <CodeBlock code={log.trimEnd()} language="plaintext" width="100%" size="sm" />
        </VStack>
      </BottomSheet>
    </>
  );
}
