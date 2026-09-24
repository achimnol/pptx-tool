import {HStack} from '@astryxdesign/core/HStack';
import {Icon} from '@astryxdesign/core/Icon';
import {IconButton} from '@astryxdesign/core/IconButton';
import {Text} from '@astryxdesign/core/Text';
import {useClipboard} from '@astryxdesign/core/hooks';

export interface CopyablePathProps {
  /** The name of the path, used in the accessible label of the copy button. */
  name: string;
  path: string;
}

export function CopyablePath({name, path}: CopyablePathProps) {
  const {copy, isCopied} = useClipboard({announce: 'Copied'});
  return (
    <HStack gap={1} align="center">
      <Text type="code">{path}</Text>
      <IconButton
        label={isCopied ? `Copied the ${name} path` : `Copy the ${name} path`}
        icon={<Icon icon={isCopied ? 'check' : 'copy'} />}
        variant="ghost"
        size="sm"
        onClick={() => void copy(path)}
      />
    </HStack>
  );
}
