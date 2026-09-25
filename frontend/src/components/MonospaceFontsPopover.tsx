import {Button} from '@astryxdesign/core/Button';
import {Popover} from '@astryxdesign/core/Popover';
import {Text} from '@astryxdesign/core/Text';
import {VStack} from '@astryxdesign/core/VStack';

import {HelpIcon} from './HelpIcon';

export interface MonospaceFontsPopoverProps {
  fonts: readonly string[];
}

export function MonospaceFontsPopover({fonts}: MonospaceFontsPopoverProps) {
  return (
    <Popover
      label="Known monospace fonts"
      width={320}
      content={
        <VStack gap={2}>
          <Text type="supporting">
            Font names are matched exactly and case-insensitively, so names with a weight suffix
            such as &ldquo;JetBrains Mono ExtraBold&rdquo; are not matched.
          </Text>
          <Text type="code">{fonts.join(', ')}</Text>
        </VStack>
      }
    >
      <Button
        label="Which fonts count as monospace?"
        icon={<HelpIcon />}
        variant="ghost"
        size="sm"
      />
    </Popover>
  );
}
