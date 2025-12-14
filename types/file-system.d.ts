export { };

declare global {
    interface Window {
        showOpenFilePicker(options?: OpenFilePickerOptions): Promise<FileSystemFileHandle[]>;
        showSaveFilePicker(options?: SaveFilePickerOptions): Promise<FileSystemFileHandle>;
    }

    interface FileSystemHandle {
        kind: 'file' | 'directory';
        name: string;
        isSameEntry(other: FileSystemHandle): Promise<boolean>;
    }

    interface FileSystemFileHandle extends FileSystemHandle {
        kind: 'file';
        getFile(): Promise<File>;
        createWritable(options?: FileSystemCreateWritableOptions): Promise<FileSystemWritableFileStream>;
    }

    interface FileSystemDirectoryHandle extends FileSystemHandle {
        kind: 'directory';
        // Add methods if needed, but for now we only need file picking
    }

    interface FileSystemWritableFileStream extends WritableStream {
        write(data: BufferSource | Blob | string | WriteParams): Promise<void>;
        seek(position: number): Promise<void>;
        truncate(size: number): Promise<void>;
    }

    interface OpenFilePickerOptions {
        multiple?: boolean;
        excludeAcceptAllOption?: boolean;
        types?: FilePickerAcceptType[];
    }

    interface SaveFilePickerOptions {
        excludeAcceptAllOption?: boolean;
        suggestedName?: string;
        types?: FilePickerAcceptType[];
    }

    interface FilePickerAcceptType {
        description?: string;
        accept: Record<string, string[]>;
    }

    interface FileSystemCreateWritableOptions {
        keepExistingData?: boolean;
    }

    type WriteParams =
        | { type: 'write'; position?: number; data: BufferSource | Blob | string }
        | { type: 'seek'; position: number }
        | { type: 'truncate'; size: number };
}
