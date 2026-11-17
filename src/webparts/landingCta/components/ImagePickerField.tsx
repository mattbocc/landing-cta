import * as React from 'react';
import {
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  IconButton,
  Label,
  MessageBar,
  MessageBarType,
  Pivot,
  PivotItem,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  TextField
} from '@fluentui/react';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { sanitizeUrl, validateUrlInput } from '../utils/urlUtils';

export interface IImagePickerFieldProps {
  label: string;
  value: string;
  context: WebPartContext;
  onChange: (newUrl: string) => void;
}

interface ISiteImageFile {
  Name: string;
  ServerRelativeUrl: string;
}

const IMAGE_EXTENSIONS: ReadonlySet<string> = new Set([
  'png',
  'jpg',
  'jpeg',
  'webp',
  'gif',
  'svg',
  'avif'
]);

const MAX_UPLOAD_BYTES = 15 * 1024 * 1024; // 15 MB

function isImageFilename(filename: string): boolean {
  const parts = filename.toLowerCase().split('.');
  if (parts.length < 2) {
    return false;
  }
  return IMAGE_EXTENSIONS.has(parts[parts.length - 1]);
}

function sanitizeFilename(filename: string): string {
  const base = filename.split(/[/\\]/).pop() || 'hero-image.png';
  return base.replace(/[~#%&*{}\\:<>?/+|"]/g, '-').trim() || 'hero-image.png';
}

export const ImagePickerField: React.FC<IImagePickerFieldProps> = ({
  label,
  value,
  context,
  onChange
}) => {
  const [isDialogOpen, setIsDialogOpen] = React.useState(false);
  const [activeTab, setActiveTab] = React.useState<string>('siteAssets');
  const [libraryFolder, setLibraryFolder] = React.useState<'SiteAssets' | 'Shared Documents'>('SiteAssets');
  const [files, setFiles] = React.useState<ISiteImageFile[]>([]);
  const [isLoading, setIsLoading] = React.useState(false);
  const [isUploading, setIsUploading] = React.useState(false);
  const [errorMsg, setErrorMsg] = React.useState<string | undefined>(undefined);
  const [selectedUrl, setSelectedUrl] = React.useState<string>(value || '');
  const [linkInput, setLinkInput] = React.useState<string>(value || '');
  const fileInputRef = React.useRef<HTMLInputElement | null>(null);

  const safeCurrentUrl = sanitizeUrl(value);

  const loadLibraryImages = React.useCallback(
    async (folder: 'SiteAssets' | 'Shared Documents'): Promise<void> => {
      setIsLoading(true);
      setErrorMsg(undefined);

      try {
        const webAbsoluteUrl = context.pageContext.web.absoluteUrl;
        const webServerRelative = context.pageContext.web.serverRelativeUrl.replace(/\/$/, '');
        const folderServerRelative = `${webServerRelative}/${folder}`;
        const endpoint =
          `${webAbsoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderServerRelative)}')/Files` +
          `?$select=Name,ServerRelativeUrl,TimeLastModified&$orderby=TimeLastModified desc&$top=60`;

        const response = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
        if (!response.ok) {
          throw new Error(`Could not load ${folder} (${response.status}).`);
        }

        const data = (await response.json()) as { value?: ISiteImageFile[] };
        const imageFiles = (data.value || []).filter((f) => isImageFilename(f.Name));
        setFiles(imageFiles);
      } catch (err) {
        setFiles([]);
        setErrorMsg(
          err instanceof Error ? err.message : `Unable to list images in ${folder}.`
        );
      } finally {
        setIsLoading(false);
      }
    },
    [context]
  );

  React.useEffect(() => {
    if (isDialogOpen && activeTab === 'siteAssets') {
      loadLibraryImages(libraryFolder).catch(() => undefined);
    }
  }, [isDialogOpen, activeTab, libraryFolder, loadLibraryImages]);

  const openPicker = (): void => {
    setSelectedUrl(value || '');
    setLinkInput(value || '');
    setErrorMsg(undefined);
    setActiveTab('siteAssets');
    setIsDialogOpen(true);
  };

  const handleConfirm = (): void => {
    if (activeTab === 'link') {
      const validationError = validateUrlInput(linkInput);
      if (validationError) {
        setErrorMsg(validationError);
        return;
      }
      onChange(linkInput.trim());
    } else if (selectedUrl) {
      onChange(selectedUrl);
    }
    setIsDialogOpen(false);
  };

  const handleUploadChange = async (
    event: React.ChangeEvent<HTMLInputElement>
  ): Promise<void> => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    if (!isImageFilename(file.name)) {
      setErrorMsg('Please select a valid image file (.png, .jpg, .webp, .gif, .svg).');
      return;
    }

    if (file.size > MAX_UPLOAD_BYTES) {
      setErrorMsg('Image exceeds the 15 MB upload limit.');
      return;
    }

    setIsUploading(true);
    setErrorMsg(undefined);

    try {
      const buffer = await file.arrayBuffer();
      const safeName = sanitizeFilename(file.name);
      const webAbsoluteUrl = context.pageContext.web.absoluteUrl;
      const webServerRelative = context.pageContext.web.serverRelativeUrl.replace(/\/$/, '');
      const folderServerRelative = `${webServerRelative}/SiteAssets`;
      const uploadUrl =
        `${webAbsoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderServerRelative)}')/Files/add(url='${encodeURIComponent(safeName)}',overwrite=true)`;

      const response = await context.spHttpClient.post(
        uploadUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata'
          },
          body: buffer
        }
      );

      if (!response.ok) {
        throw new Error(`Upload failed (${response.status}). Check your permissions for Site Assets.`);
      }

      const result = (await response.json()) as { ServerRelativeUrl?: string };
      const uploadedUrl = result.ServerRelativeUrl || `${folderServerRelative}/${safeName}`;
      onChange(uploadedUrl);
      setIsDialogOpen(false);
    } catch (err) {
      setErrorMsg(err instanceof Error ? err.message : 'Failed to upload image.');
    } finally {
      setIsUploading(false);
      if (fileInputRef.current) {
        fileInputRef.current.value = '';
      }
    }
  };

  return (
    <div style={{ marginBottom: 12 }}>
      <Label>{label}</Label>

      {safeCurrentUrl && (
        <div
          style={{
            marginBottom: 8,
            borderRadius: 4,
            overflow: 'hidden',
            border: '1px solid #edebe9',
            backgroundColor: '#f3f2f1'
          }}
        >
          <img
            src={safeCurrentUrl}
            alt="Selected background preview"
            style={{
              display: 'block',
              width: '100%',
              height: 112,
              objectFit: 'cover'
            }}
          />
        </div>
      )}

      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
        <DefaultButton text="Change image" onClick={openPicker} />
        <IconButton
          iconProps={{ iconName: 'Delete' }}
          title="Remove image"
          ariaLabel="Remove image"
          disabled={!value}
          onClick={() => onChange('')}
        />
      </div>

      <Dialog
        hidden={!isDialogOpen}
        onDismiss={() => setIsDialogOpen(false)}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: 'Pick a background image'
        }}
        modalProps={{ isBlocking: false }}
        minWidth={560}
        maxWidth={720}
      >
        {errorMsg && (
          <MessageBar
            messageBarType={MessageBarType.error}
            onDismiss={() => setErrorMsg(undefined)}
            styles={{ root: { marginBottom: 12 } }}
          >
            {errorMsg}
          </MessageBar>
        )}

        <Pivot
          selectedKey={activeTab}
          onLinkClick={(item) => {
            if (item?.props.itemKey) {
              setActiveTab(item.props.itemKey);
              setErrorMsg(undefined);
            }
          }}
        >
          <PivotItem headerText="Site library" itemKey="siteAssets">
            <div style={{ marginTop: 12 }}>
              <div style={{ display: 'flex', gap: 8, marginBottom: 12 }}>
                <DefaultButton
                  text="Site Assets"
                  primary={libraryFolder === 'SiteAssets'}
                  onClick={() => setLibraryFolder('SiteAssets')}
                />
                <DefaultButton
                  text="Documents"
                  primary={libraryFolder === 'Shared Documents'}
                  onClick={() => setLibraryFolder('Shared Documents')}
                />
              </div>

              {isLoading ? (
                <div style={{ padding: 24, textAlign: 'center' }}>
                  <Spinner size={SpinnerSize.medium} label="Loading images..." />
                </div>
              ) : files.length === 0 ? (
                <p style={{ color: '#605e5c', margin: '16px 0' }}>
                  No images found in {libraryFolder}. Use the Upload tab to add one.
                </p>
              ) : (
                <div
                  style={{
                    display: 'grid',
                    gridTemplateColumns: 'repeat(auto-fill, minmax(120px, 1fr))',
                    gap: 10,
                    maxHeight: 280,
                    overflowY: 'auto',
                    padding: 2
                  }}
                >
                  {files.map((file) => {
                    const isSelected = selectedUrl === file.ServerRelativeUrl;
                    return (
                      <button
                        key={file.ServerRelativeUrl}
                        type="button"
                        onClick={() => setSelectedUrl(file.ServerRelativeUrl)}
                        onDoubleClick={() => {
                          onChange(file.ServerRelativeUrl);
                          setIsDialogOpen(false);
                        }}
                        style={{
                          padding: 4,
                          borderRadius: 4,
                          border: isSelected ? '2px solid #0078d4' : '1px solid #edebe9',
                          backgroundColor: isSelected ? '#eff6fc' : '#ffffff',
                          cursor: 'pointer',
                          textAlign: 'left'
                        }}
                      >
                        <img
                          src={file.ServerRelativeUrl}
                          alt={file.Name}
                          style={{
                            width: '100%',
                            height: 80,
                            objectFit: 'cover',
                            borderRadius: 2,
                            display: 'block'
                          }}
                        />
                        <div
                          style={{
                            marginTop: 4,
                            fontSize: 11,
                            color: '#323130',
                            overflow: 'hidden',
                            textOverflow: 'ellipsis',
                            whiteSpace: 'nowrap'
                          }}
                          title={file.Name}
                        >
                          {file.Name}
                        </div>
                      </button>
                    );
                  })}
                </div>
              )}
            </div>
          </PivotItem>

          <PivotItem headerText="Upload" itemKey="upload">
            <div style={{ marginTop: 16 }}>
              <p style={{ margin: '0 0 12px', color: '#323130' }}>
                Upload an image from your device into this site&apos;s <strong>Site Assets</strong> library.
              </p>
              <input
                ref={fileInputRef}
                type="file"
                accept="image/*"
                style={{ display: 'none' }}
                onChange={(e) => {
                  handleUploadChange(e).catch(() => undefined);
                }}
              />
              <PrimaryButton
                text={isUploading ? 'Uploading...' : 'Choose file to upload'}
                disabled={isUploading}
                onClick={() => fileInputRef.current?.click()}
              />
            </div>
          </PivotItem>

          <PivotItem headerText="From a link" itemKey="link">
            <div style={{ marginTop: 12 }}>
              <TextField
                label="Image URL (https:// or /site-path)"
                placeholder="/sites/your-site/SiteAssets/hero.jpg"
                value={linkInput}
                onChange={(_, newValue) => {
                  setLinkInput(newValue || '');
                  setErrorMsg(undefined);
                }}
              />
            </div>
          </PivotItem>
        </Pivot>

        <DialogFooter>
          {activeTab !== 'upload' && (
            <PrimaryButton
              text="Select"
              onClick={handleConfirm}
              disabled={activeTab === 'siteAssets' && !selectedUrl}
            />
          )}
          <DefaultButton text="Cancel" onClick={() => setIsDialogOpen(false)} />
        </DialogFooter>
      </Dialog>
    </div>
  );
};
