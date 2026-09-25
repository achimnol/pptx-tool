import {Button} from '@astryxdesign/core/Button';
import {Link} from '@astryxdesign/core/Link';
import {Popover} from '@astryxdesign/core/Popover';
import {Text} from '@astryxdesign/core/Text';
import {VStack} from '@astryxdesign/core/VStack';

import {FONT_DOWNLOADS} from '../theme-editor/fontDownloads';
import {HelpIcon} from './HelpIcon';

export function FontDownloadsPopover() {
  return (
    <Popover
      label="Font downloads"
      width={320}
      content={
        <VStack gap={2}>
          <Text type="supporting">
            The official download pages of the fonts used by the bundled themes.
          </Text>
          <VStack gap={1}>
            {FONT_DOWNLOADS.map(({name, href}) => (
              <Link key={href} href={href} isExternalLink isStandalone>
                {name}
              </Link>
            ))}
          </VStack>
        </VStack>
      }
    >
      <Button label="Where to download the fonts?" icon={<HelpIcon />} variant="ghost" size="sm" />
    </Popover>
  );
}
