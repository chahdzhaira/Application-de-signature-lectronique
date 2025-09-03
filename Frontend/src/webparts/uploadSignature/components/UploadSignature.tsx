import * as React from 'react';
import { useRef, useState } from 'react';
import { MessageBar, MessageBarType, PrimaryButton } from '@fluentui/react';
import { ArrowUploadFilled } from '@fluentui/react-icons';
import styles from './UploadSignature.module.scss';
import { sp } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IUploadSignatureProps } from './IUploadSignatureProps';


const UploadSignature: React.FC<IUploadSignatureProps> = (props) => {
  const [image, setImage] = useState<string | null>(null);
  const [message, setMessage] = useState<{ text: string; type: MessageBarType } | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const userName = props.context.pageContext.user.displayName;
  const userEmail = props.context.pageContext.user.email;

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file && file.type.startsWith('image/')) {
      const reader = new FileReader();
      reader.onload = () => setImage(reader.result as string);
      reader.readAsDataURL(file);
    }
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file && file.type.startsWith('image/')) {
      const reader = new FileReader();
      reader.onload = () => setImage(reader.result as string);
      reader.readAsDataURL(file);
    }
  };

  const handleClick = () => {
    fileInputRef.current?.click();
  };

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
  };

  const handleOkClick = async () => {
    if (!image) return;

    try {
      const base64 = image.split(',')[1];
      const byteCharacters = atob(base64);
      const byteNumbers = Array.from(byteCharacters).map(char => char.charCodeAt(0));
      const byteArray = new Uint8Array(byteNumbers);
      const blob = new Blob([byteArray], { type: 'image/png' });

      const fileExtension = image.split(';')[0].split('/')[1];
      const fileName = `signature_${userName}.${fileExtension}`;

      const existingFiles = await sp.web.getFolderByServerRelativeUrl("/sites/App_signature_electronique/SignaturesLibrary")
        .files.filter(`Title eq 'Signature for ${userName}'`).get();

      if (existingFiles.length > 0) {
        const fileToDelete = existingFiles[0];
        await sp.web.getFileByServerRelativeUrl(fileToDelete.ServerRelativeUrl).delete();
        console.log("Old signature removed");
      }
      const uploadResult = await sp.web.getFolderByServerRelativeUrl("/sites/App_signature_electronique/SignaturesLibrary").files.add(fileName, blob, true);
      const item = await uploadResult.file.getItem();
      await item.update({ Title: `Signature for ${userName}`, UserEmail: userEmail });

      console.log("File + metadata added :", item);
      setMessage({ text: "Your signature has been added successfully!", type: MessageBarType.success });
      setTimeout(() => props.onClose?.(), 1500);
    } catch (err) {
      console.error("Upload error :", err);
      setMessage({ text: "Error while uploading your signature.", type: MessageBarType.error });
    }
  };

  return (
    <div className={styles.wrapper}>
      <h2 className={styles.title}>Add your signature</h2>
      {message && (
        <MessageBar messageBarType={message.type} isMultiline={false} onDismiss={() => setMessage(null)} dismissButtonAriaLabel="Fermer"> {message.text} </MessageBar>
      )}
      <div className={`${styles.dropZone} ${image ? styles.dropZoneWithImage : ''}`} onDrop={handleDrop} onDragOver={handleDragOver} onClick={handleClick} >
        <input type="file" accept="image/*" ref={fileInputRef} onChange={handleFileChange} style={{ display: 'none' }} />
        {image ? (
          <img src={image} alt="Signature" className={styles.previewImage} />
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 4 }}>
            <ArrowUploadFilled style={{ fontSize: 24, color: '#666' }} />
            <p style={{ margin: 0 }}>Upload a signature</p>
          </div>
        )}

      </div>

      {image && (<PrimaryButton text="Save" onClick={handleOkClick} className={styles.btn} />)}
    </div>
  );
};

export default UploadSignature;