import * as React from 'react';
import { useState, useEffect } from 'react';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { Label } from '@fluentui/react/lib/Label';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { TagPicker, TagPickerControl, TagPickerOption, TagPickerGroup, Tag, Avatar, Field, TagPickerInput, TagPickerList } from "@fluentui/react-components";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { RadioGroup, Radio } from '@fluentui/react-components';
import { DragDropContext, Droppable, Draggable } from 'react-beautiful-dnd';
import { ArrowUploadFilled, DismissCircleFilled, PersonAddFilled, PersonDeleteFilled, GridDotsFilled } from '@fluentui/react-icons';
import styles from './SignatureRequest.module.scss';
import { sp } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { v4 as uuidv4 } from 'uuid';

interface ISignatureRequestProps {
  context: WebPartContext;
}

interface Signer {
  id: string;
  email: string;
  searchTerm: string;
  filteredOptions: { email: string; displayName: string; jobTitle: string }[];
}

const SignatureRequest: React.FC<ISignatureRequestProps> = (props) => {
  const [file, setFile] = useState<File | null>(null);
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const [message, setMessage] = useState('');
  const [emailOptions, setEmailOptions] = useState<{ email: string, displayName: string, jobTitle: string }[]>([]);
  const [filteredEmailOptions, setFilteredEmailOptions] = useState<string[]>([]);
  const [signers, setSigners] = useState<Signer[]>([{ id: '', email: '', searchTerm: '', filteredOptions: [] }]);
  const [signingMode, setSigningMode] = useState<'Parallel' | 'Sequential'>();
  const [globalError, setGlobalError] = useState<string | null>(null);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);
  const [isSubmitting, setIsSubmitting] = React.useState(false);
  const errorRef = React.useRef<HTMLDivElement>(null);
  const requestorEmail = props.context.pageContext.user.email;
  const [groupRequests, setGroupRequests] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [isEmailOptionsLoaded, setIsEmailOptionsLoaded] = useState(false);
  const [isEditMode, setIsEditMode] = useState(false);
  const [itemId, setItemId] = useState<string | null>(null);

  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const id = params.get("itemId");
    if (id) {
      setItemId(id);
      setIsEditMode(true);
    }
  }, []);

  useEffect(() => {
    if (signers.length > 1 && signingMode === undefined) {
      setSigningMode('Parallel');
    }
  }, [signers.length, signingMode]);


  useEffect(() => {
    const fetchUsers = async () => {
      try {
        const client = await props.context.msGraphClientFactory.getClient("3");
        const response = await client.api('/users?$select=mail,displayName,jobTitle').top(50).get();
        const users = response.value.filter((user: any) => user.mail).map((user: any) => ({
          email: user.mail,
          displayName: user.displayName,
          jobTitle: user.jobTitle || 'Not specified'
        }));
        setEmailOptions(users);
        setSigners([{ id: '', email: '', searchTerm: '', filteredOptions: users }]);
        setFilteredEmailOptions(users);
        console.log(filteredEmailOptions)
        setIsEmailOptionsLoaded(true);
      } catch (err) {
        console.error('Error retrieving users :', err);
      }
    };
    fetchUsers();
  }, []);

  useEffect(() => {
    if ((globalError || successMessage) && errorRef.current) {
      errorRef.current.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
  }, [globalError, successMessage]);

  useEffect(() => {
    if (!isEmailOptionsLoaded) return;

    const fetchGroupData = async () => {
      const requestGroupId = getRequestGroupIdFromUrl();

      if (!requestGroupId) {
        console.error("Aucun itemId dans l’URL !");
        return;
      }

      try {
        let items = [];
        if (requestGroupId.startsWith("single-")) {
          const idOnly = requestGroupId.replace("single-", "");
          const item = await sp.web.lists.getByTitle("SignatureRequest").items.getById(parseInt(idOnly)).get();
          items = [item];
        } else {
          items = await sp.web.lists.getByTitle("SignatureRequest")
            .items
            .filter(`RequestGroupId eq '${requestGroupId}'`)
            .select("ID", "SignerEmail", "SenderEmail", "Message", "ApprovalID", "ApprovalLink", "TypeSignature", "DocumentID/ID", "ApprovalDecision")
            .orderBy("OrderNumber", true)
            .expand("DocumentID")
            .get();
        }

        if (items.length > 0) {
          setGroupRequests(items);
          console.log(groupRequests)

          const signersFromItems = items
            .sort((a, b) => (a.OrderNumber ?? 9999) - (b.OrderNumber ?? 9999))
            .map((item) => ({
              id: item.ID.toString(),
              email: item.SignerEmail,
              searchTerm: '',
              filteredOptions: emailOptions
            }));

          setSigners(signersFromItems);
          setMessage(items[0].Message || '');
          setSigningMode(items[0].TypeSignature);
        }

        const fileId = items[0].DocumentID?.ID;
        if (fileId) {
          try {
            const file = await sp.web.lists
              .getByTitle("DocumentsLibrary")
              .items.getById(fileId)
              .file.get();

            const fileRef = file.ServerRelativeUrl;

            if (!fileRef) {
              console.error("Aucun chemin de fichier trouvé.");
              return;
            }

            const fileResponse = await fetch(fileRef, { credentials: 'include' });

            if (!fileResponse.ok) {
              console.error("Erreur lors du fetch du fichier :", fileResponse.statusText);
              return;
            }

            const blob = await fileResponse.blob();
            const fileName = decodeURIComponent(fileRef.split('/').pop() || 'document.pdf');
            const loadedFile = new File([blob], fileName, { type: blob.type });

            console.log("Fichier PDF chargé avec succès :", loadedFile);
            setFile(loadedFile);
          } catch (err) {
            console.error("Erreur lors du chargement du fichier PDF depuis SharePoint :", err);
          }
        }



      } catch (err) {
        console.error("Erreur lors du chargement du groupe :", err);
      } finally {
        setLoading(false);
        console.log(loading)

      }
    };

    fetchGroupData();
  }, [isEmailOptionsLoaded]);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (selectedFile && selectedFile.type === "application/pdf") {
      setFile(selectedFile);
      setGlobalError(null);
    } else {
      setFile(null);
      setGlobalError("Only PDF files are allowed.");
    }
  };

  const handleAddSigner = () => {
    setSigners([...signers, { id: Date.now().toString(), email: '', searchTerm: '', filteredOptions: emailOptions }]);
  };

  const handleRemoveSigner = (index: number) => {
    const updatedSigners = signers.filter((_, i) => i !== index);
    setSigners(updatedSigners);
    if (updatedSigners.length < 2) {
      setSigningMode(undefined);
    }
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    const droppedFile = e.dataTransfer.files[0];
    if (droppedFile && droppedFile.type === "application/pdf") {
      setFile(droppedFile);
      setGlobalError(null);
    } else {
      setFile(null);
      setGlobalError("Only PDF files are allowed.");
    }
  };

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
  };

  const handleClick = () => {
    fileInputRef.current?.click();
  };

  const handleDragEnd = (result: any) => {
    if (!result.destination) return;
    const reordered = Array.from(signers);
    const [removed] = reordered.splice(result.source.index, 1);
    reordered.splice(result.destination.index, 0, removed);
    setSigners(reordered);
  };

  const uploadFileToSharePoint = async (file: File) => {
    try {
      const folder = sp.web.getFolderByServerRelativeUrl("DocumentsLibrary/DocToBeSigned");
      const uploadedFile = await folder.files.add(file.name, file, true);
      const listItem = await uploadedFile.file.getItem();
      const itemId = (listItem as any).Id;

      setSuccessMessage(`File "${file.name}" imported successfully !`);
      return {
        data: {
          UniqueId: itemId
        },
        fileName: file.name
      };
    } catch (error) {
      console.error("Erreur lors de l'upload du fichier :", error);
      setGlobalError("Erreur lors de l'import du fichier dans SharePoint.");
    }
  };

  const createSignatureRequestItems = async (fileId: string, fileName: string) => {
    const createdItemIds: number[] = [];
    const requestGroupId = uuidv4();

    try {
      for (let index = 0; index < signers.length; index++) {
        const signer = signers[index];
        if (signer.email) {
          const item = await sp.web.lists.getByTitle("SignatureRequest").items.add({
            Title: fileName,
            SignerEmail: signer.email,
            Message: message,
            DocumentIDId: Number(fileId),
            SenderEmail: requestorEmail,
            OrderNumber: signingMode === 'Sequential' ? index + 1 : 1,
            TypeSignature: signingMode || 'OneSigner',
            RequestGroupId: requestGroupId
          });
          createdItemIds.push(item.data.ID);
        }
      };
      setSuccessMessage("Signature requests have been successfully created.");
      return createdItemIds;

    } catch (error) {
      console.error("Error creating elements in SignatureRequest:", error);
      setGlobalError("An error occurred while creating the signature requests.");
    }
  };

  const handleNext = async () => {

    if (!file) {
      setGlobalError("Please upload a document.");
      return;
    }
    // Vérifier qu’aucun signer n’a d’email vide
    const emptySigner = signers.find(s => !s.email?.trim());
    if (emptySigner) {
      setGlobalError("Please provide the email addresses of all signatories.");
      return;
    }

    const emails = signers.map(s => s.email.trim());
    const hasDuplicate = emails.some((email, i) => emails.indexOf(email) !== i);
    if (hasDuplicate) {
      setGlobalError("You cannot send this same file to a duplicate user !");
      return;
    }
    setIsSubmitting(true);
    setGlobalError(null);
    try {
      const uploadedFile = await uploadFileToSharePoint(file);
      if (uploadedFile?.data?.UniqueId) {
        const fileId = uploadedFile.data.UniqueId;
        const fileName = file.name;
        const createdItemIds = await createSignatureRequestItems(fileId, fileName);
        const formattedSigners = signers.map((signer, index) => ({
          signerEmail: signer.email,
          orderNumber: signingMode === 'Sequential' ? index + 1 : 1,
        }));

        const requestPayload = {
          fileName: file.name,
          requestorEmail: requestorEmail,
          signers: formattedSigners,
          message,
          fileId: fileId,
          signingMode: signers.length === 1 ? 'OneSigner' : signingMode,
          signatureRequestItemIds: createdItemIds,
        };
        const response = await fetch("https://prod-55.northeurope.logic.azure.com:443/workflows/263d83f77f6b4d649750730902eef23e/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=U6Wf4aAOb8yUvCLo2bH6G6qzz7IPWFo8O2BHWjtNdbc", {
          method: "POST",
          headers: {
            "Content-Type": "application/json"
          },
          body: JSON.stringify(requestPayload)
        });

        if (response.ok) {
          setSuccessMessage("Signature request sent successfully !");
          setGlobalError(null);
          setIsSubmitting(false);
          window.location.href = "https://48c2lx.sharepoint.com/sites/App_signature_electronique/SitePages/SignatureRequestOverview.aspx";
        } else {
          setSuccessMessage(null);
          setGlobalError("Error sending request to Power Automate.");
          setIsSubmitting(false);
        }

      } else {
        setGlobalError("The file was uploaded but the ID was not found.");
      }
    } catch (err) {
      console.error("Erreur lors de l'exécution :", err);
      setGlobalError("An error has occurred! Please try again.");
      setIsSubmitting(false);
    }

    setIsSubmitting(false);
  };

  const truncateFileName = (name: string, maxLength = 20) => {
    const nameWithoutExt = name.replace(/\.[^/.]+$/, "");
    if (nameWithoutExt.length > maxLength) {
      return nameWithoutExt.substring(0, maxLength) + "...";
    }
    return nameWithoutExt;
  };

  const handleOptionSelect = (index: number) => (
    e: React.SyntheticEvent,
    data: { value: string }
  ) => {
    const updatedSigners = [...signers];
    updatedSigners[index].email = updatedSigners[index].email === data.value ? '' : data.value;
    setSigners(updatedSigners);
  };

  useEffect(() => {
    const fetchFileFromUrl = async () => {
      const urlParams = new URLSearchParams(window.location.search);
      const fileUrl = urlParams.get("fileUrl");

      if (fileUrl) {
        try {
          const response = await fetch(fileUrl, {
            credentials: 'include'
          });

          const blob = await response.blob();
          const fileName = decodeURIComponent(fileUrl.split('/').pop() || 'document.pdf');
          const file = new File([blob], fileName, { type: blob.type });
          setFile(file);
        } catch (error) {
          console.error("Erreur lors du chargement du fichier depuis l'URL :", error);
          setGlobalError("Impossible de charger le fichier.");
        }
      }
    };

    fetchFileFromUrl();
  }, []);

  const getRequestGroupIdFromUrl = (): string | null => {
    const params = new URLSearchParams(window.location.search);
    return params.get("itemId");
  };

  const handleEdit = async () => {
    if (!itemId) {
      setGlobalError("No id found for the modification.");
      return;
    }

    if (!file) {
      setGlobalError("Please select a file.");
      return;
    }

    try {
      setIsSubmitting(true);
      setGlobalError(null);

      const items = await sp.web.lists
        .getByTitle("SignatureRequest")
        .items
        .filter(`RequestGroupId eq '${itemId}'`)
        .get();

      if (items.length === 0) {
        setGlobalError("No items found to modify.");
        setIsSubmitting(false);
        return;
      }

      const requestGroupId = items[0].RequestGroupId;
      const allIds = items.map(item => item.ID);

      for (const item of items) {
        await fetch("https://prod-24.northeurope.logic.azure.com:443/workflows/04b7fe4e6051451e8d460b4f61c515e3/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=UCKgzZ7baadk4-rp7MB_VoWa7kTfcVnImthT6SoZpzk", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            userEmail: item.SignerEmail,
            signatureRequestId: allIds,
            approvalId: item.ApprovalID
          })
        });
      }

      const uploadedFile = await uploadFileToSharePoint(file);
      if (!uploadedFile?.data?.UniqueId) {
        setGlobalError("Error importing the file into SharePoint.");
        setIsSubmitting(false);
        return;
      }

      const newFileId = uploadedFile.data.UniqueId;
      const createdOrUpdatedIds: number[] = [];
      for (let i = 0; i < signers.length; i++) {
        const signer = signers[i];
        const added = await sp.web.lists.getByTitle("SignatureRequest").items.add({
          Title: file?.name || '',
          DocumentIDId: Number(newFileId),
          SenderEmail: requestorEmail,
          SignerEmail: signer.email,
          OrderNumber: signingMode === 'Sequential' ? i + 1 : 1,
          Message: message,
          TypeSignature: signers.length === 1 ? "OneSigner" : signingMode,
          RequestGroupId: requestGroupId
        });
        createdOrUpdatedIds.push(added.data.ID);
      }
      const formattedSigners = signers.map((signer, index) => ({
        signerEmail: signer.email,
        orderNumber: signingMode === 'Sequential' ? index + 1 : 1
      }));

      await fetch("https://prod-55.northeurope.logic.azure.com:443/workflows/263d83f77f6b4d649750730902eef23e/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=U6Wf4aAOb8yUvCLo2bH6G6qzz7IPWFo8O2BHWjtNdbc", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          fileName: file.name,
          requestorEmail,
          signers: formattedSigners,
          message,
          fileId: newFileId,
          signingMode: signers.length === 1 ? 'OneSigner' : signingMode,
          signatureRequestItemIds: createdOrUpdatedIds
        })
      });

      setSuccessMessage("Requests successfully modified and re-launched");
      setIsSubmitting(false);
      setTimeout(() => {
        window.location.href = "https://48c2lx.sharepoint.com/sites/App_signature_electronique/SitePages/SignatureRequestOverview.aspx";
      }, 1500);

    } catch (error) {
      console.error("Erreur lors de la modification", error);
      setGlobalError("Error during modification");
      setIsSubmitting(false);
    }
  };

  return (
    <div>

      <div ref={errorRef}>
        {globalError && (<MessageBar className={styles.messageBar} messageBarType={MessageBarType.error} isMultiline={false} onDismiss={() => setGlobalError(null)} dismissButtonAriaLabel="Close"> {globalError} </MessageBar>)}
        {successMessage && (<MessageBar className={styles.messageBar} messageBarType={MessageBarType.success} onDismiss={() => setSuccessMessage(null)} isMultiline={false} > {successMessage} </MessageBar>)}
      </div>

      <div style={{ maxWidth: 600, padding: 20 }}>
        <Label required>Import file</Label>
        <div className={`${styles.dropZone} ${file ? styles.dropZoneWithFile : ''}`} onDrop={handleDrop} onDragOver={handleDragOver} onClick={handleClick} style={{ position: 'relative' }} >
          <input type="file" accept=".pdf" ref={fileInputRef} onChange={handleFileChange} style={{ display: 'none' }} />
          {file ? (
            <div>
              <img src={require("../assets/documentIcon.png")} alt="Document icon" className={styles.previewFile} />
              <DismissCircleFilled
                className={styles.removeFileIcon}
                onClick={(e) => {
                  e.stopPropagation();
                  setFile(null);
                }}
                title="Remove file"
                aria-label="Remove file"
              />
              <p className={styles.text}>{truncateFileName(file.name)}</p>
            </div>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 4 }}>
              <ArrowUploadFilled style={{ fontSize: 24, color: '#666' }} />
              <p className={styles.subtitleMd}>Upload a PDF file</p>
              <p className={styles.subtitleSm}>File supported: PDF</p>
              <p className={styles.subtitleSm}>Only one file is allowed !</p>
            </div>
          )}
        </div>
        <br />

        {signers.length > 1 && (
          <RadioGroup layout="horizontal" value={signingMode} onChange={(_, data) => setSigningMode(data.value as 'Parallel' | 'Sequential')} className={styles.RadioGroup}>
            <Radio value="Parallel" label="Parallel" />
            <Radio value="Sequential" label="Sequential" />
          </RadioGroup>
        )}

        <DragDropContext onDragEnd={handleDragEnd}>
          <Droppable droppableId="signers">
            {(provided) => (
              <div {...provided.droppableProps} ref={provided.innerRef}>
                {signers.map((signer, index) => (
                  <Draggable key={signer.id} draggableId={signer.id} index={index}>
                    {(provided) => (
                      <div ref={provided.innerRef} {...provided.draggableProps} {...provided.dragHandleProps} style={{ display: 'flex', alignItems: 'center', marginBottom: 10, ...provided.draggableProps.style }} >
                        {signers.length > 1 && (
                          <div {...provided.dragHandleProps} style={{ marginRight: 4, cursor: 'grab' }}>
                            <GridDotsFilled className={styles.iconDrag} />
                          </div>
                        )}
                        <Field label={signingMode === 'Sequential' ? `Signer ${index + 1}` : 'Signer Email'} style={{ flex: 1 }} required >
                          <TagPicker onOptionSelect={handleOptionSelect(index)} selectedOptions={signer.email ? [signer.email] : []}>
                            <TagPickerControl>
                              {signer.email && (
                                <TagPickerGroup aria-label="Selected Email">
                                  <Tag key={signer.email} shape="rounded" media={<Avatar aria-hidden name={signer.email} color="colorful" />} value={signer.email} >{signer.email} </Tag>
                                </TagPickerGroup>
                              )}
                              <TagPickerInput aria-label="Select a signer" />
                            </TagPickerControl>
                            <TagPickerList>
                              {emailOptions
                                .filter((user) =>
                                  // N'affiche pas les emails déjà sélectionnés ailleurs
                                  !signers.some((s, i) => s.email === user.email && i !== index)
                                )
                                .map((user) => (
                                  <TagPickerOption key={user.email} value={user.email} media={<Avatar name={user.email} color="colorful" />} secondaryContent={`${user.email} • ${user.jobTitle || 'Not specified'} `} >{user.displayName}</TagPickerOption>
                                ))}
                            </TagPickerList>
                          </TagPicker>
                        </Field>
                        {signers.length > 1 && (
                          <PersonDeleteFilled className={styles.iconDelete} onClick={() => handleRemoveSigner(index)} title="Delete signer" />
                        )}
                      </div>
                    )}
                  </Draggable>
                ))}
                {provided.placeholder}
              </div>
            )}
          </Droppable>
        </DragDropContext>

        <PersonAddFilled className={styles.iconAdd} title="Add another signer" onClick={handleAddSigner} />

        <Label>Message</Label>
        <textarea rows={3} style={{ width: '100%' }} value={message} onChange={(e) => setMessage(e.target.value)} />
        <br /><br />
        <PrimaryButton text={isEditMode ? "Edit" : "Send"} onClick={isEditMode ? handleEdit : handleNext} className={styles.btn} disabled={isSubmitting} />
      </div>
    </div>
  );

};

export default SignatureRequest;
