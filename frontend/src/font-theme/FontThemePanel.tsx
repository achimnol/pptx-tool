import {AlertDialog} from '@astryxdesign/core/AlertDialog';
import {Banner} from '@astryxdesign/core/Banner';
import {Button} from '@astryxdesign/core/Button';
import {Heading} from '@astryxdesign/core/Heading';
import {HStack} from '@astryxdesign/core/HStack';
import {Text} from '@astryxdesign/core/Text';
import {TextInput} from '@astryxdesign/core/TextInput';
import {useToast} from '@astryxdesign/core/Toast';
import {VStack} from '@astryxdesign/core/VStack';
import {useState} from 'react';

import {downloadFontTheme, installFontTheme} from '../api/actions';
import {ApiError, type FieldError} from '../api/errors';
import type {ThemeData} from '../api/types';
import {downloadBlob} from '../utils/download';
import {validateFontThemeName} from './validateFontThemeName';

export interface FontThemePanelProps {
  theme: ThemeData;
  isThemeValid: boolean;
  isLocal: boolean;
  onFieldErrors: (errors: FieldError[]) => void;
}

type Action = 'download' | 'install';

const MACOS_THEME_FONTS_DIR =
  '~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Themes.localized/Theme Fonts';
const WINDOWS_THEME_FONTS_DIR = '%APPDATA%\\Microsoft\\Templates\\Document Themes\\Theme Fonts';

export function FontThemePanel({theme, isThemeValid, isLocal, onFieldErrors}: FontThemePanelProps) {
  const toast = useToast();
  const [name, setName] = useState('');
  const [isNameTouched, setIsNameTouched] = useState(false);
  const [runningAction, setRunningAction] = useState<Action | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [isConfirmingOverwrite, setIsConfirmingOverwrite] = useState(false);

  const nameError = validateFontThemeName(name);
  const canRun = nameError === null && isThemeValid && runningAction === null;

  const handleError = (e: unknown) => {
    setError(e instanceof Error ? e.message : String(e));
    if (e instanceof ApiError) onFieldErrors(e.fieldErrors);
  };

  const handleDownload = async () => {
    setRunningAction('download');
    setError(null);
    try {
      const blob = await downloadFontTheme(name.trim(), theme);
      downloadBlob(blob, `${name.trim()}.xml`);
    } catch (e) {
      handleError(e);
    } finally {
      setRunningAction(null);
    }
  };

  const handleInstall = async (overwrite: boolean) => {
    setRunningAction('install');
    setError(null);
    try {
      const path = await installFontTheme(name.trim(), theme, overwrite);
      setIsConfirmingOverwrite(false);
      toast({body: `Installed the font theme at ${path}. Restart Office to use it.`});
    } catch (e) {
      if (e instanceof ApiError && e.status === 409 && !overwrite) {
        setIsConfirmingOverwrite(true);
      } else {
        setIsConfirmingOverwrite(false);
        handleError(e);
      }
    } finally {
      setRunningAction(null);
    }
  };

  return (
    <VStack gap={4}>
      <Heading level={2}>Office font theme</Heading>
      <Banner
        status="info"
        title="An Office font theme only defines the heading and body fonts."
        description="Bold titles, the first-level bullet weight and monospace preservation do not apply. Restart the Office apps after installing a font theme to see it in the Design ribbon."
      />
      <TextInput
        label="Font theme name"
        value={name}
        onChange={(value) => {
          setName(value);
          setIsNameTouched(true);
        }}
        description="The name shown in Office. It is also used as the file name."
        status={isNameTouched && nameError ? {type: 'error', message: nameError} : undefined}
      />
      <HStack gap={2}>
        <Button
          label="Download XML"
          variant={isLocal ? 'secondary' : 'primary'}
          onClick={() => void handleDownload()}
          isLoading={runningAction === 'download'}
          isDisabled={!canRun}
        />
        {isLocal && (
          <Button
            label="Install to Office"
            variant="primary"
            onClick={() => void handleInstall(false)}
            isLoading={runningAction === 'install'}
            isDisabled={!canRun}
          />
        )}
      </HStack>
      <VStack gap={1}>
        <Text type="supporting">
          Store the downloaded XML file in Office&rsquo;s theme fonts directory, then restart
          Office:
        </Text>
        <Text type="supporting">macOS</Text>
        <Text type="code">{MACOS_THEME_FONTS_DIR}</Text>
        <Text type="supporting">Windows</Text>
        <Text type="code">{WINDOWS_THEME_FONTS_DIR}</Text>
      </VStack>
      {!isThemeValid && (
        <Banner status="warning" title="Fix the invalid theme fields before continuing." />
      )}
      {error !== null && (
        <Banner status="error" title="Failed to generate the font theme" description={error} />
      )}
      <AlertDialog
        isOpen={isConfirmingOverwrite}
        onOpenChange={setIsConfirmingOverwrite}
        title="Overwrite the font theme?"
        description={`A font theme named “${name.trim()}” already exists. Overwriting it cannot be undone.`}
        actionLabel="Overwrite"
        onAction={() => void handleInstall(true)}
        isActionLoading={runningAction === 'install'}
      />
    </VStack>
  );
}
