import * as React from 'react';
import { useState, useRef } from 'react';
import SignatureCanvas from 'react-signature-canvas';
import { MessageBar, MessageBarType, PrimaryButton } from '@fluentui/react';
import styles from './PenSignature.module.scss';
import { sp } from '@pnp/sp/presets/all';
import { IPenSignatureProps } from './IPenSignatureProps';


const PenSignature: React.FC<IPenSignatureProps> = (props) => {
  const [result, setResult] = useState<string | null>(null);
  const [message, setMessage] = useState<{ text: string; type: MessageBarType } | null>(null);

  const signatureRef = useRef<any>(null);
  const userName = props.context.pageContext.user.displayName;
  const userEmail = props.context.pageContext.user.email;

  const handleClear = () => {
    signatureRef.current.clear();
    setResult(null);
  };

  const handleSave = async () => {
    if (!result || !userName || !userEmail) {
      console.log("Missing required data: base64Image, userName, or userEmail");
      setMessage({ text: "Missing data for registration.", type: MessageBarType.error });
      return;
    }

    try {
      const base64 = result.split(',')[1];
      const byteCharacters = atob(base64);
      const byteNumbers = Array.from(byteCharacters).map(char => char.charCodeAt(0));
      const byteArray = new Uint8Array(byteNumbers);
      const blob = new Blob([byteArray], { type: 'image/png' });
      const fileName = `signature_${userName}.png`;
      const existingFiles = await sp.web.getFolderByServerRelativeUrl("/sites/App_signature_electronique/SignaturesLibrary").files.filter(`Title eq 'Signature for ${userName}'`).get();

      if (existingFiles.length > 0) {
        const fileToDelete = existingFiles[0];
        await sp.web.getFileByServerRelativeUrl(fileToDelete.ServerRelativeUrl).delete();
      }

      const uploadResult = await sp.web.getFolderByServerRelativeUrl("/sites/App_signature_electronique/SignaturesLibrary").files.add(fileName, blob, true);
      const item = await uploadResult.file.getItem();
      await item.update({
        Title: `Signature for ${userName}`,
        UserEmail: userEmail
      });
      setMessage({ text: "Your handwritten signature has been saved!", type: MessageBarType.success });
      setTimeout(() => props.onClose?.(), 1500);
    } catch (err) {
      console.error("Erreur lors de l'enregistrement :", err);
      setMessage({ text: "Error while saving signature.", type: MessageBarType.error });
    }
  };

  const handleEnd = () => {
    const dataUrl = signatureRef.current.toDataURL();
    setResult(dataUrl);
  };

  return (
    <div className={styles.wrapper}>
      <h2 className={styles.title}>Provide your handwritten signature</h2>
      {message && (
        <MessageBar messageBarType={message.type} isMultiline={false} onDismiss={() => setMessage(null)} dismissButtonAriaLabel="Fermer" >{message.text}</MessageBar>
      )}
      <div className={styles.canvasWrapper}>
        <SignatureCanvas penColor="black" ref={signatureRef} velocityFilterWeight={0.2} minWidth={0.5} maxWidth={2.5} canvasProps={{ width: 500, height: 200, className: 'sigCanvas', }} onEnd={handleEnd} />
      </div>

      <div style={{ marginTop: '12px', display: 'flex', gap: '10px' }}>
        <PrimaryButton text="Clear" onClick={handleClear} className={styles.btn} />
        <PrimaryButton text="Save" onClick={handleSave} className={styles.btn} />
      </div>
    </div>
  );
};

export default PenSignature;
