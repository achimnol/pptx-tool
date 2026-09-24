import {HStack} from '@astryxdesign/core/HStack';
import {Icon} from '@astryxdesign/core/Icon';
import {IconButton} from '@astryxdesign/core/IconButton';
import {Text} from '@astryxdesign/core/Text';
import {useAnnounce} from '@astryxdesign/core/hooks';
import {useEffect, useRef, useState} from 'react';

import {copyText} from '../utils/clipboard';

/** How long the checkmark replaces the copy icon after a successful copy. */
const COPIED_DURATION_MS = 1500;

export interface CopyablePathProps {
  /** The name of the path, used in the accessible label of the copy button. */
  name: string;
  path: string;
}

export function CopyablePath({name, path}: CopyablePathProps) {
  const announce = useAnnounce();
  const [isCopied, setIsCopied] = useState(false);
  // Counts the successful copies so that each one replays the checkmark animation.
  const [copyCount, setCopyCount] = useState(0);
  const resetTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  useEffect(
    () => () => {
      if (resetTimerRef.current !== null) clearTimeout(resetTimerRef.current);
    },
    [],
  );

  const handleCopy = async () => {
    try {
      await copyText(path);
    } catch {
      return;
    }
    setIsCopied(true);
    setCopyCount((count) => count + 1);
    announce('Copied');
    if (resetTimerRef.current !== null) clearTimeout(resetTimerRef.current);
    resetTimerRef.current = setTimeout(() => {
      resetTimerRef.current = null;
      setIsCopied(false);
    }, COPIED_DURATION_MS);
  };

  return (
    <HStack gap={1} align="center">
      <Text type="code" size="sm">
        {path}
      </Text>
      <IconButton
        className="copyable-path-button"
        label={isCopied ? `Copied the ${name} path` : `Copy the ${name} path`}
        icon={
          // The key remounts the icon on every change, so each swap plays its entry animation.
          <span
            key={isCopied ? `check-${copyCount}` : `copy-${copyCount}`}
            className={isCopied ? 'copyable-path-icon is-copied' : 'copyable-path-icon'}
            data-animated={copyCount > 0 ? '' : undefined}
          >
            <Icon icon={isCopied ? 'check' : 'copy'} color="inherit" />
          </span>
        }
        variant="ghost"
        size="sm"
        onClick={() => void handleCopy()}
      />
    </HStack>
  );
}
