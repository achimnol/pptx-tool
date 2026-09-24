import {Button} from '@astryxdesign/core/Button';
import {Link} from '@astryxdesign/core/Link';
import {Popover} from '@astryxdesign/core/Popover';
import {Text} from '@astryxdesign/core/Text';
import {VStack} from '@astryxdesign/core/VStack';

import {FONT_DOWNLOADS} from '../theme-editor/fontDownloads';

// A circled question mark in the style of the Astryx default icons, which have no help icon.
const helpIcon = (
  <svg
    xmlns="http://www.w3.org/2000/svg"
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth={1.5}
    strokeLinecap="round"
    strokeLinejoin="round"
    width="1em"
    height="1em"
    aria-hidden="true"
  >
    <circle cx="12" cy="12" r="9" />
    <path d="M9.5 9.5a2.5 2.5 0 1 1 3.5 2.3c-.7.3-1 .9-1 1.7" />
    <circle cx="12" cy="17" r="0.6" fill="currentColor" />
  </svg>
);

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
      <Button label="Where to download the fonts?" icon={helpIcon} variant="ghost" size="sm" />
    </Popover>
  );
}
