import * as React from 'react';
import styles from './Home.module.scss';
import { useState, useEffect } from 'react';
import { PrimaryButton, Spinner, SpinnerSize } from '@fluentui/react';
import { Dialog, DialogSurface, DialogBody, DialogTitle, DialogContent } from '@fluentui/react-components';
import { AddCircleFilled } from "@fluentui/react-icons";

import { sp } from '@pnp/sp/presets/all';
import { ChooseSignature } from '../../chooseSignature/components/ChooseSignature'
import { IHomeProps } from './IHomeProps';



const Home: React.FC<IHomeProps> = (props) => {
  const [signatureUrl, setSignatureUrl] = useState<string | null>(null);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);
  const [showDialog, setShowDialog] = useState<boolean>(false);

  const username = props.context.pageContext.user.displayName;


  const fetchUserSignature = async (email: string) => {
    try {
      const items = await sp.web.lists.getByTitle("SignaturesLibrary")
        .items
        .select("File/ServerRelativeUrl", "File/Name", "File/Title", "UserEmail")
        .expand("File")
        .filter(`UserEmail eq '${email}'`)
        .top(1)
        .get();

      if (items.length > 0) {
        const signature = items[0];
        const timestamp = new Date().getTime();
        const fullUrl = `${window.location.protocol}//${window.location.hostname}${signature.File.ServerRelativeUrl}?v=${timestamp}`;
        setSignatureUrl(fullUrl);
      } else {
        setSignatureUrl(null);
      }
    } catch (error) {
      console.error("Error retrieving signature:", error);
      setError("Error loading signature");
      setSignatureUrl(null);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    const fetchData = async () => {
      const email = props.context.pageContext.user.email;
      if (email) {
        await fetchUserSignature(email);
      } else {
        setLoading(false);
      }
    };

    fetchData();
  }, []);

  const handleImageClick = () => {
    setShowDialog(true);
  };

  const handleCloseDialog = async () => {
    setShowDialog(false);
    const email = props.context.pageContext.user.email;
    if (email) {
      setLoading(true);
      await fetchUserSignature(email);
    }
  };

  return (
    <>
      <div className={styles.horizontalBox}>
        <div className={styles.welcomeMessage}>
          Welcome {username} ! <br /> Ready to send or sign documents with one click
        </div>
        <div>
          {loading ? (
            <div className={styles.signatureContainer}>
              <Spinner size={SpinnerSize.large} />
            </div>
          ) : error ? (
            <p style={{ color: 'red' }}>{error}</p>

          ) : signatureUrl ? (
            <div className={styles.signatureContainer}>
              <img
                src={signatureUrl}
                alt="Your signature"
                className={styles.signatureImage}
                onClick={handleImageClick}
              />
              <button className={styles.changeButton} onClick={handleImageClick}>Change</button>
            </div>
          ) : (
            <div className={styles.cardTitle}>
              <h3>Add your signature</h3>
              <AddCircleFilled role="button" onClick={handleImageClick} className={styles.addIcon} />
            </div>
          )}
        </div>
      </div>
      <div>
        <h1 className={styles.title}>Get started</h1>
      </div>
      <div className={styles.boxesContainer}>
        <div className={styles.box}>
          <h3 className={styles.cardTitle}>Start a new signature request</h3>
          <PrimaryButton className={styles.btn} text="Start" onClick={() => { window.location.href = "https://48c2lx.sharepoint.com/sites/App_signature_electronique/SitePages/Demande-Signature.aspx" }} />
        </div>
      </div>
      <Dialog open={showDialog} onOpenChange={(_, data) => setShowDialog(data.open)}>
        <DialogSurface >
          <div className={styles.dialogSurface} >
            <DialogBody>
              <DialogTitle className={styles.titleDialog}>
                Change your signature
              </DialogTitle>
              <DialogContent className={styles.dialogContent} >
                <ChooseSignature context={props.context} onClose={handleCloseDialog} />
              </DialogContent>
            </DialogBody>
          </div>
        </DialogSurface>
      </Dialog>
    </>
  );
};

export default Home;