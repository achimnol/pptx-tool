import {Banner} from '@astryxdesign/core/Banner';
import {Button} from '@astryxdesign/core/Button';
import {FileInput} from '@astryxdesign/core/FileInput';
import {TextInput} from '@astryxdesign/core/TextInput';
import {useToast} from '@astryxdesign/core/Toast';
import {VStack} from '@astryxdesign/core/VStack';
import {useState} from 'react';

import {fixFont} from '../api/actions';
import {ApiError, type FieldError} from '../api/errors';
import type {ThemeData} from '../api/types';
import {ProcessLog} from '../components/ProcessLog';
import {downloadBlob} from '../utils/download';
import {defaultOutputName, ensurePptxSuffix, formatBytes} from '../utils/filenames';

export interface FixFontsPanelProps {
  theme: ThemeData;
  isThemeValid: boolean;
  maxUploadSize: number | undefined;
  onFieldErrors: (errors: FieldError[]) => void;
}

export function FixFontsPanel({
  theme,
  isThemeValid,
  maxUploadSize,
  onFieldErrors,
}: FixFontsPanelProps) {
  const toast = useToast();
  const [file, setFile] = useState<File | null>(null);
  const [outputName, setOutputName] = useState('');
  const [isRunning, setIsRunning] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [log, setLog] = useState<string | null>(null);

  const handleFileChange = (value: File | File[] | null) => {
    const selected = Array.isArray(value) ? (value[0] ?? null) : value;
    setFile(selected);
    setOutputName(selected ? defaultOutputName(selected.name) : '');
    setError(null);
    setLog(null);
  };

  const handleRun = async () => {
    if (file === null) return;
    setIsRunning(true);
    setError(null);
    setLog(null);
    onFieldErrors([]);
    try {
      const result = await fixFont(file, theme);
      const filename = ensurePptxSuffix(outputName.trim() || result.filename);
      downloadBlob(result.file, filename);
      setLog(result.log);
      toast({body: `Fixed the fonts and saved ${filename}.`});
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
      if (e instanceof ApiError) onFieldErrors(e.fieldErrors);
    } finally {
      setIsRunning(false);
    }
  };

  const canRun = file !== null && isThemeValid && !isRunning;
  return (
    <VStack gap={4}>
      <FileInput
        label="Source file"
        mode="dropzone"
        accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
        maxSize={maxUploadSize}
        value={file}
        onChange={handleFileChange}
        description={
          file
            ? `${file.name} (${formatBytes(file.size)})`
            : 'Drop a .pptx file here or choose one.'
        }
      />
      <TextInput
        label="Output file name"
        value={outputName}
        onChange={setOutputName}
        isDisabled={file === null}
        disabledMessage="Choose a presentation first."
        description="The fixed presentation is downloaded with this name."
      />
      <Button
        label="Fix fonts"
        variant="primary"
        onClick={() => void handleRun()}
        isLoading={isRunning}
        isDisabled={!canRun}
      />
      {!isThemeValid && (
        <Banner status="warning" title="Fix the invalid theme fields before running." />
      )}
      {error !== null && (
        <Banner status="error" title="Failed to fix the fonts" description={error} />
      )}
      {log !== null && <ProcessLog log={log} />}
    </VStack>
  );
}
