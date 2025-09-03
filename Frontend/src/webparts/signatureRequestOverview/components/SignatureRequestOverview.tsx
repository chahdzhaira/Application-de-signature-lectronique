import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './SignatureRequestOverview.module.scss';
import { sp } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/files";
import { Table, TableHeader, TableHeaderCell, TableBody, TableRow, TableCell, TabValue, Button, useToastController, Toast, ToastTitle, ToastBody, ToastFooter, Link, useId, TableCellLayout, useArrowNavigationGroup, TableSelectionCell } from "@fluentui/react-components";
import { ArrowDownloadFilled, CheckmarkCircleRegular, DismissCircleRegular, EyeRegular, MailInboxArrowDownFilled, SendFilled } from "@fluentui/react-icons"
import { IconButton, Label, MessageBar, MessageBarType, Modal, PrimaryButton, Spinner, Stack, Dropdown, TextField } from '@fluentui/react';
import { EditRegular, AppsListDetailRegular, DeleteRegular } from '@fluentui/react-icons';
import { ISignatureRequestOverviewProps } from './ISignatureRequestOverviewProps';


const SignatureRequestOverview: React.FC<ISignatureRequestOverviewProps> = (props) => {
  const [requests, setRequests] = useState<any[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [userEmail, setUserEmail] = useState<string | null>(null);
  const [selectedTab, setSelectedTab] = useState<TabValue>('sent');
  const [signatureUrl, setSignatureUrl] = useState<string | null>(null);
  const [loading, setLoading] = useState<boolean>(true);
  const [lloading, setLLoading] = useState<boolean>(true);
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 6;
  const [submittingIds, setSubmittingIds] = useState<Set<number>>(new Set());
  const [globalError, setGlobalError] = useState<string | null>(null);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);
  const errorRef = React.useRef<HTMLDivElement>(null);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [isModalDetailsOpen, setIsModalDetailsOpen] = useState(false);
  const [selectedDocumentUrl, setSelectedDocumentUrl] = useState<string | null>(null);
  const [hasSignature, setHasSignature] = useState<boolean | null>(null);
  const [statusFilter, setStatusFilter] = useState<string>("all");
  const [searchQuery, setSearchQuery] = useState<string>("");
  const [startDate, setStartDate] = useState<string>("");
  const [endDate, setEndDate] = useState<string>("");
  const [isRefreshing, setIsRefreshing] = useState(false);
  const [selectedRequest, setSelectedRequest] = useState<any | null>(null);
  const [groupedRequestsToShow, setGroupedRequestsToShow] = useState<any[]>([]);
  const [rawRequests, setRawRequests] = useState<any[]>([]);
  const keyboardNavAttr = useArrowNavigationGroup({ axis: "grid" });
  const [selectedRows, setSelectedRows] = React.useState<Set<number>>(new Set());
  const [isProcessing, setIsProcessing] = React.useState(false);

  const toasterId = useId("toaster");
  const { dispatchToast } = useToastController(toasterId);

  const notify = (message: string, type: "success" | "error" = "success") => {
    const toastClass = type === "success" ? styles.toastSuccess : styles.toastError;
    dispatchToast(
      <Toast className={toastClass}>
        <ToastTitle>{type === "success" ? "Succès" : "Erreur"}</ToastTitle>
        <ToastBody subtitle="Notification">{message}</ToastBody>
        <ToastFooter>
          <Link onClick={() => console.log("Toast dismissed")}>Fermer</Link>
        </ToastFooter>
      </Toast>,
      { intent: type }
    );
  };

  useEffect(() => {
    const fetchUserEmail = async () => {
      const email = await props.context.pageContext.user.email;
      if (email) {
        setUserEmail(email);
      } else {
        notify("Impossible de récupérer l'utilisateur", "error");
      }
    };

    fetchUserEmail();
  }, []);

  useEffect(() => {
    fetchSignatureRequests(true);
    const interval = setInterval(() => fetchSignatureRequests(false), 60000);
    return () => clearInterval(interval);
  }, [userEmail]);

  useEffect(() => {
    const fetchData = async () => {
      const email = await props.context.pageContext.user.email;

      if (email) {
        await fetchUserSignature(email);
      } else {
        setLoading(false);
        console.log(loading);
      }
    };

    fetchData();
  }, []);

  useEffect(() => {
    if ((globalError || successMessage) && errorRef.current) {
      errorRef.current.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
  }, [globalError, successMessage]);

  useEffect(() => {
    setCurrentPage(1);
  }, [statusFilter, searchQuery, startDate, endDate]);

  useEffect(() => {
    setCurrentPage(1);
  }, [selectedTab]);

  const fetchSignatureRequests = async (firstLoad = false, manualRefresh = false) => {

    if (!userEmail) return;
    if (manualRefresh) setIsRefreshing(true);
    try {
      const data = await sp.web.lists.getByTitle('SignatureRequest')
        .items
        .select('ID', 'Title', 'Created', 'SignerEmail', 'SenderEmail', 'ApprovalLink', 'ApprovalID', 'ApprovalDecision', 'ApprovalStatus', 'DocumentID/ID', 'TypeSignature', 'RequestGroupId', 'OrderNumber')
        .expand('DocumentID')
        .filter(`SenderEmail eq '${userEmail}' or SignerEmail eq '${userEmail}'`)
        .get();

      const enrichedData = await Promise.all(data.map(async (request) => {
        const signerName = await getUserDisplayName(request.SignerEmail);
        const senderName = await getUserDisplayName(request.SenderEmail);

        let status = "Pend";
        let fileName = "Document";
        let fileUrl = "";

        if (request.DocumentID?.ID) {
          try {
            const docItem = await sp.web.lists.getByTitle("DocumentsLibrary")
              .items.getById(request.DocumentID.ID)
              .select("FileLeafRef", "FileRef", "Status")
              .get();

            fileName = docItem.FileLeafRef || "Document";
            status = docItem.Status || "Pend";
            fileUrl = docItem.FileRef || "";
          } catch (docErr) {
            console.warn(`Unable to retrieve document details ${request.DocumentID.ID}`, docErr);
          }
        }

        return {
          id: request.ID,
          created: request.Created,
          title: request.Title,
          signerEmail: request.SignerEmail,
          senderEmail: request.SenderEmail,
          approvalLink: request.ApprovalLink,
          approvalId: request.ApprovalID,
          signerName,
          senderName,
          status,
          fileName,
          fileUrl,
          approvalDecision: request.ApprovalDecision,
          approvalStatus: request.ApprovalStatus,
          typeSignature: request.TypeSignature,
          requestGroupId: request.RequestGroupId,
          orderNumber: request.OrderNumber !== undefined && request.OrderNumber !== null ? Number(request.OrderNumber) : null
        };
      }));

      setRawRequests(enrichedData);

      const groupedRequests: { [key: string]: any } = {};

      for (const item of enrichedData) {
        const groupId = item.requestGroupId;
        if (groupId) {
          if (!groupedRequests[groupId]) {
            groupedRequests[groupId] = {
              ...item,
              ids: [item.id],
              signerEmails: [item.signerEmail],
            };
          } else {
            groupedRequests[groupId].signerEmails.push(item.signerEmail);
            groupedRequests[groupId].ids.push(item.id);
          }
        } else {
          groupedRequests[`single-${item.id}`] = {
            ...item,
            ids: [item.id],
            signerEmails: [item.signerEmail],
          };
        }
      }

      const groupedArray = Object.keys(groupedRequests).map(key => groupedRequests[key]);
      setRequests(groupedArray.sort((a, b) => new Date(b.created).getTime() - new Date(a.created).getTime()));


      if (firstLoad) setLLoading(false);
    } catch (err: any) {
      console.error('Request retrieval error:', err);
      setError(err.message);
      console.log(error)
      notify("Request retrieval error", "error");
      if (firstLoad) setLLoading(false);
    } finally {
      if (manualRefresh) setIsRefreshing(false);
    }
  };

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
        setHasSignature(true);
      } else {
        setSignatureUrl(null);
        setHasSignature(false);
      }
    } catch (error) {
      console.error("Error retrieving signature :", error);
      setError("Error loading signature");
      setSignatureUrl(null);
      setHasSignature(false);
    } finally {
      setLoading(false);
    }
  };

  const handleResponse = async (id: number, responseValue: "Approve" | "Reject") => {

    if (!userEmail || submittingIds.has(id)) return;

    if (responseValue === "Approve" && hasSignature === false) {
      setGlobalError("You must first add your signature before approving this document.");
      return;
    }

    const updatedIds = new Set(submittingIds);
    updatedIds.add(id);
    setSubmittingIds(updatedIds);

    const reqData = requests.find(r => r.id === id);
    if (!reqData) {
      setGlobalError("Request not found");
      return;
    }


    try {
      const flowUrl = "https://prod-49.northeurope.logic.azure.com:443/workflows/2f3f6575bc0249ed85d04b9eb7516c97/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=oAY7mjDOoIbZ5sQcFvzK3plh4fJzH3d5kx3KdRCBDZE";

      const response = await fetch(flowUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ userEmail: userEmail, signatureRequestId: id, approvalLink: reqData.approvalLink, approvalId: reqData.approvalId, Response: responseValue })
      });

      if (response.ok) {
        setSuccessMessage(`Document ${responseValue === "Approve" ? "approuved" : "rejected"} successfully !`);

      } else {
        setSuccessMessage(null);
        const err = await response.text();
        console.error("Erreur lors de l'exécution :", err);
        setGlobalError("An error has occurred! Please try again.");
      }
    } catch (error) {
      console.error("Erreur Power Automate:", error);
      setSuccessMessage(null);
      setGlobalError("Error sending request to Power Automate.");
    }
  };

  const getUserDisplayName = async (email: string): Promise<string> => {
    try {
      const user = await sp.web.siteUsers.getByEmail(email).get();
      return user.Title || user.LoginName || email;
    } catch (err) {
      console.warn(err);
      return email;
    }
  };

  const sentRequests = requests.filter(req => req.senderEmail?.toLowerCase() === userEmail?.toLowerCase());
  const receivedRequests = rawRequests.filter(
    req => {
      const isReceived = req.signerEmail?.toLowerCase() === userEmail?.toLowerCase();
      return isReceived;
    }
  );

  const currentRequests = selectedTab === 'sent'
    ? sentRequests
    : receivedRequests

  const totalPages = Math.ceil(currentRequests.length / itemsPerPage);
  const filteredRequests = currentRequests.filter((req) => {
    const matchesStatus =
      statusFilter === "all" || req.approvalDecision?.toLowerCase() === statusFilter;

    const matchesSearch =
      req.fileName?.toLowerCase().includes(searchQuery.toLowerCase()) ||
      req.signerName?.toLowerCase().includes(searchQuery.toLowerCase()) ||
      req.senderEmail?.toLowerCase().includes(searchQuery.toLowerCase());

    const requestDate = new Date(req.created);

    const matchesStartDate = startDate ? requestDate >= new Date(startDate) : true;
    const matchesEndDate = endDate ? requestDate <= new Date(endDate) : true;

    return matchesStatus && matchesSearch && matchesStartDate && matchesEndDate;
  });

  const paginatedRequests = filteredRequests.slice(
    (currentPage - 1) * itemsPerPage,
    currentPage * itemsPerPage
  );

  const goToPreviousPage = () => {
    setCurrentPage(prev => Math.max(prev - 1, 1));
  };

  const goToNextPage = () => {
    setCurrentPage(prev => Math.min(prev + 1, totalPages));
  };

  const handleView = (url: string) => {
    setSelectedDocumentUrl(url);
    setIsModalOpen(true);
  };

  const closeModal = () => {
    setIsModalOpen(false);
    setSelectedDocumentUrl(null);
  };

  const handleDeleteRequest = async (ids: number[]) => {
    if (!window.confirm("Do you really want to delete these requests?")) return;
    setSubmittingIds(prev => {
      const newSet = new Set(prev);
      ids.forEach(id => newSet.add(id));
      return newSet;
    });

    try {
      const reqData = requests.find(r => r.id === ids[0]);
      if (!reqData) throw new Error("Request not found");
      const flowUrl = "https://prod-24.northeurope.logic.azure.com:443/workflows/04b7fe4e6051451e8d460b4f61c515e3/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=UCKgzZ7baadk4-rp7MB_VoWa7kTfcVnImthT6SoZpzk";
      const response = await fetch(flowUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          userEmail: userEmail,
          signatureRequestId: ids,
          approvalId: reqData.approvalId,
          approvalLink: reqData.approvalLink
        })
      });

      if (!response.ok) throw new Error("Error when calling the flow");

      for (const id of ids) {
        const item = await sp.web.lists
          .getByTitle("SignatureRequest")
          .items.getById(id)
          .select("DocumentID/ID")
          .expand("DocumentID")
          .get();

        const documentId = item.DocumentID?.ID;

        await sp.web.lists.getByTitle("SignatureRequest").items.getById(id).delete();
        if (documentId) {
          const linkedRequests = await sp.web.lists
            .getByTitle("SignatureRequest")
            .items.select("ID")
            .filter(`DocumentID eq ${documentId}`)
            .top(1)
            .get();

          if (linkedRequests.length === 0) {
            await sp.web.lists.getByTitle("DocumentsLibrary").items.getById(documentId).delete();
            console.log(`📄 Document ${documentId} deleted`);
          } else {
            console.log(`⏸ Document ${documentId} conservé `);
          }
        }
      }

      setRequests(prev => prev.filter(req =>
        ids.some((id: number) => ids.indexOf(id) !== -1)
      ));
      notify("Requests successfully deleted", "success");
    } catch (err) {
      console.error("Error during deletion :", err);
      setGlobalError("Error during deletion.");
    } finally {
      setSubmittingIds(prev => {
        const copy = new Set(prev);
        ids.forEach(id => copy.delete(id));
        return copy;
      });
    }
  };

  const handleViewDetails = (request: any) => {
    setSelectedRequest(request);
    console.log(selectedRequest)

    if (request.requestGroupId) {
      const group = rawRequests
        .filter(r => r.requestGroupId === request.requestGroupId)
        .sort((a, b) => {
          const aNum = a.orderNumber != null ? Number(a.orderNumber) : Number(new Date(a.created).getTime());
          const bNum = b.orderNumber != null ? Number(b.orderNumber) : Number(new Date(b.created).getTime());
          return aNum - bNum;
        });
      setGroupedRequestsToShow(group);
      console.log(rawRequests);
    } else {
      setGroupedRequestsToShow([request]);
    }

    setIsModalDetailsOpen(true);
  };


  const closeViewDetailsModal = () => {
    setIsModalDetailsOpen(false);
    setSelectedDocumentUrl(null);
  };

  const handleEditClick = (requestGroupId: string) => {
    window.location.href = `https://48c2lx.sharepoint.com/sites/App_signature_electronique/SitePages/Demande-Signature.aspx?itemId=${requestGroupId}`;
  };

  const toggleRowSelection = (id: number) => {
    setSelectedRows(prev => {
      const newSet = new Set(prev);
      if (newSet.has(id)) {
        newSet.delete(id);
      } else {
        newSet.add(id);
      }
      return newSet;
    });
  };


  const isAllVisibleRowsSelected = paginatedRequests.length > 0 &&
    paginatedRequests.every(req => selectedRows.has(req.id));

  const isSomeVisibleRowsSelected = paginatedRequests.some(req => selectedRows.has(req.id));

  const bulkRespond = async (responseValue: "Approve" | "Reject") => {
    if (selectedRows.size === 0) {
      setGlobalError("Please select at least one item.");
      return;
    }
    if (isProcessing) {
      setGlobalError("Please wait until the documents are signed.");
      return;
    }

    setIsProcessing(true);
    setSuccessMessage(`Processing your ${responseValue.toLowerCase()} request...`);
    const selectedItems = requests.filter(req => selectedRows.has(req.id));

    const payload = {
      userEmail: selectedItems[0].signerEmail,
      signatureRequestId: selectedItems.map(req => req.id.toString()),
      Response: responseValue
    };
    try {
      const res = await fetch("https://prod-31.northeurope.logic.azure.com:443/workflows/bcd78a3af64c40659c2e9e42bfafab87/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=fHu2W9xtZa-A7fVenYvbCfVihuKa4TtYNTFLLFfrpMw", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });

      if (res.ok) {
        setSuccessMessage("Successfully processed ");
        fetchSignatureRequests(false, true);
        setSelectedRows(new Set());
      } else {
        const errorText = await res.text();
        console.error("Erreur :", errorText);
        setGlobalError("Error during bulk sending.");
      }
    } catch (err) {
      console.error("Exception :", err);
      setGlobalError("Connection issue");
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <>
      <div className={styles.wrapper}>
        <div className={styles.sidebar}>
          <div className={styles.userCard}>
            {signatureUrl ? (
              <>
                <div>
                  <h4>Your signature</h4>
                </div>
                <div>
                  <img src={signatureUrl} alt="Signature" className={styles.signatureImg} />
                </div>
              </>
            ) : (
              <div>
                <h4>You haven't added a signature yet !</h4>
              </div>
            )}
          </div>
          <ul className={styles.sidebarMenu}>
            <li onClick={() => setSelectedTab('inbox')} className={selectedTab === 'inbox' ? styles.active : ''}>
              <h3><MailInboxArrowDownFilled className={styles.icons} />Inbox</h3>
            </li>
            <li onClick={() => setSelectedTab('sent')} className={selectedTab === 'sent' ? styles.active : ''}>
              <h3><SendFilled className={styles.icons} />Sent</h3>
            </li>
          </ul>
        </div>
        <div className={styles.mainContent}>
          <div ref={errorRef}>
            {globalError && (<MessageBar className={styles.messageBar} messageBarType={MessageBarType.error} isMultiline={false} onDismiss={() => setGlobalError(null)} dismissButtonAriaLabel="Close"> {globalError} </MessageBar>)}
            {successMessage && (<MessageBar className={styles.messageBar} messageBarType={MessageBarType.success} onDismiss={() => setSuccessMessage(null)} isMultiline={false} > {successMessage} </MessageBar>)}
          </div>
          <Stack horizontal wrap tokens={{ childrenGap: 20 }} styles={{ root: { marginBottom: 20 } }} verticalAlign="end">
            <Stack>
              <Label>Status</Label>
              <Dropdown
                selectedKey={statusFilter}
                onChange={(_, option) => setStatusFilter(option?.key as string)}
                options={[
                  { key: 'all', text: 'All' },
                  { key: 'pending', text: 'Pending' },
                  { key: 'approve', text: 'Approved' },
                  { key: 'reject', text: 'Rejected' },
                ]}
                styles={{ dropdown: { width: 150 } }}
              />
            </Stack>
            <Stack>
              <Label>Search</Label>
              <TextField placeholder="Search by file name..." value={searchQuery} onChange={(_, newValue) => setSearchQuery(newValue || '')} styles={{ root: { width: 300 } }} />
            </Stack>
            <Stack>
              <Label>From</Label>
              <input type="date" id="startDate" value={startDate} onChange={(e) => setStartDate(e.target.value)} style={{ marginLeft: "0.5rem" }} />
            </Stack>
            <Stack>
              <Label>To</Label>
              <input type="date" id="endDate" value={endDate} onChange={(e) => setEndDate(e.target.value)} style={{ marginLeft: "0.5rem" }} />
            </Stack>
          </Stack>
          {selectedTab === 'sent' && (
            <Table>
              <div style={{ display: 'flex', alignItems: 'center', gap: '0.3rem', marginBottom: '0.5rem' }}>
                <Button appearance="transparent" onClick={() => fetchSignatureRequests(false, true)} disabled={isRefreshing} icon={<span style={{ fontSize: '18px', color: isRefreshing ? '#A6A6A6' : 'black', opacity: isRefreshing ? 0.5 : 1, cursor: isRefreshing ? 'not-allowed' : 'pointer', }}>🔁</span>} style={{ minWidth: '32px', height: '32px', padding: 0 }} aria-label="Refresh" />
                <h2 style={{ margin: 0 }}>Sent</h2>
              </div>
              {lloading ? (
                <div style={{ textAlign: 'center', padding: '2rem' }}>
                  <Spinner className={styles.spinner} />
                </div>
              ) : (
                <>
                  <TableHeader>
                    <TableRow>
                      <TableHeaderCell>Filename</TableHeaderCell>
                      <TableHeaderCell>Date</TableHeaderCell>
                      <TableHeaderCell>Status</TableHeaderCell>
                      <TableHeaderCell>Mode signature</TableHeaderCell>
                      <TableHeaderCell>Action</TableHeaderCell>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {paginatedRequests.map(req => (
                      <TableRow key={req.id}>
                        <TableCell title={req.fileName}>{req.title.length > 19 ? req.title.substring(0, 19) + "..." : req.title}</TableCell>
                        <TableCell>{new Date(req.created).toLocaleString()}</TableCell>
                        <TableCell>{req.approvalStatus}</TableCell>
                        <TableCell>{req.typeSignature}</TableCell>
                        <TableCell>
                          <div className={styles.buttonGroup}>
                            <IconButton style={{ marginLeft: 0, marginRight: 0 }} className={styles.btn} onClick={() => handleViewDetails(req)} aria-label="View details" title="View details">
                              <AppsListDetailRegular style={{ fontSize: 16, color: 'white' }} />
                            </IconButton>
                            <IconButton onClick={() => handleEditClick(req.requestGroupId)} disabled={req.approvalDecision !== 'Pending'} className={styles.btn} aria-label="Edit" title="Edit">
                              <EditRegular style={{ fontSize: 16, color: 'white' }} />
                            </IconButton>
                            <div>
                              <IconButton className={styles.btn} onClick={() => handleDeleteRequest(req.ids)} disabled={req.approvalStatus === 'Approved' || req.approvalDecision !== 'Pending'} aria-label="Delete" title="Delete">
                                <DeleteRegular style={{ fontSize: 16, color: 'white' }} />
                              </IconButton>
                            </div>
                          </div>
                        </TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </>
              )}
            </Table>
          )}

          {selectedTab === 'inbox' && (
            <>
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '0.5rem' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '0.3rem', marginBottom: '0.5rem' }}>
                  <Button appearance="transparent" onClick={() => fetchSignatureRequests(false, true)} disabled={isRefreshing} icon={<span style={{ fontSize: '18px', color: isRefreshing ? '#A6A6A6' : 'black', opacity: isRefreshing ? 0.5 : 1, cursor: isRefreshing ? 'not-allowed' : 'pointer', }}>🔁</span>} style={{ minWidth: '32px', height: '32px', padding: 0 }} aria-label="Refresh" />
                  <h2 style={{ margin: 0 }}>Inbox</h2>
                </div>
                <Stack horizontal tokens={{ childrenGap: 6 }}>
                  <PrimaryButton onClick={() => bulkRespond("Approve")} disabled={selectedRows.size === 0 || isProcessing} style={{ backgroundColor: 'transparent', color: selectedRows.size === 0 || isProcessing ? '#a6a6a6' : '#183060', border: 'none', fontWeight: '600', }} iconProps={{ iconName: 'CheckMark' }} text="Approve" />
                  <PrimaryButton onClick={() => bulkRespond("Reject")} disabled={selectedRows.size === 0 || isProcessing} style={{ backgroundColor: 'transparent', color: selectedRows.size === 0 || isProcessing ? '#a6a6a6' : '#d13438', border: 'none', fontWeight: '600', }} iconProps={{ iconName: 'Cancel' }} text="Reject" />
                </Stack>
              </div>
              <Table>
                {lloading ? (
                  <div style={{ textAlign: 'center', padding: '2rem' }}>
                    <Spinner />
                  </div>
                ) : (
                  <>
                    <TableHeader>
                      <TableRow>
                        <TableSelectionCell
                          checked={isAllVisibleRowsSelected ? true : isSomeVisibleRowsSelected ? 'mixed' : false}
                          onClick={() => {
                            const onlyPendingIds = paginatedRequests
                              .filter(r => r.approvalDecision?.toLowerCase() === "pending")
                              .map(r => r.id);

                            const allSelected = onlyPendingIds.every(id => selectedRows.has(id));

                            setSelectedRows(prev => {
                              const newSet = new Set(prev);
                              if (allSelected) {
                                onlyPendingIds.forEach(id => newSet.delete(id));
                              } else {
                                onlyPendingIds.forEach(id => newSet.add(id));
                              }
                              return newSet;
                            });
                          }}
                          checkboxIndicator={{ 'aria-label': 'Select all rows' }}
                        />
                        <TableHeaderCell>Filename</TableHeaderCell>
                        <TableHeaderCell>Date</TableHeaderCell>
                        <TableHeaderCell>From</TableHeaderCell>
                        <TableHeaderCell>Status</TableHeaderCell>
                        <TableHeaderCell>Action</TableHeaderCell>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {paginatedRequests
                        .filter(req => req.approvalDecision?.toLowerCase() !== "new")
                        .map((req) => (
                          <TableRow key={req.id} aria-selected={selectedRows.has(req.id)}>
                            {req.approvalDecision?.toLowerCase() === "pending" ? (
                              <TableSelectionCell
                                checked={selectedRows.has(req.id)}
                                onClick={() => {
                                  if (req.approvalDecision?.toLowerCase() === "pending") {
                                    toggleRowSelection(req.id);
                                  }
                                }}
                                checkboxIndicator={{ 'aria-label': 'Select row' }}
                                aria-disabled={req.approvalDecision?.toLowerCase() !== "pending" || isProcessing}
                                style={{
                                  pointerEvents: req.approvalDecision?.toLowerCase() !== "pending" || isProcessing ? 'none' : 'auto',
                                  opacity: req.approvalDecision?.toLowerCase() !== "pending" || isProcessing ? 0.4 : 1,
                                }}
                              />
                            ) : (
                              <TableCell />
                            )}
                            <TableCell>
                              <span className={styles.fileNameLink} title={req.fileName} onClick={() => handleView(req.fileUrl)}>
                                {req.fileName.length > 19 ? req.fileName.substring(0, 19) + "..." : req.fileName}
                              </span>
                            </TableCell>
                            <TableCell>{new Date(req.created).toLocaleString()}</TableCell>
                            <TableCell>{req.senderName}</TableCell>
                            <TableCell>{req.approvalDecision}</TableCell>
                            <TableCell>
                              {req.approvalDecision?.toLowerCase() === "pending" ? (
                                <div className={styles.buttonGroup}>
                                  <PrimaryButton text="Approve" className={styles.btn} onClick={() => handleResponse(req.id, "Approve")} disabled={submittingIds.has(req.id)} >
                                    <CheckmarkCircleRegular style={{ fontSize: 16, color: 'white' }} />
                                  </PrimaryButton>
                                  <PrimaryButton text="Reject" className={styles.btn} onClick={() => handleResponse(req.id, "Reject")} disabled={submittingIds.has(req.id)} >
                                    <DismissCircleRegular style={{ fontSize: 16, color: 'white' }} />
                                  </PrimaryButton>
                                </div>
                              ) : (
                                <div className={styles.buttonGroup}>
                                  <IconButton style={{ marginLeft: 0, marginRight: 0 }} className={styles.btn} onClick={() => handleView(req.fileUrl)} aria-label="View" title="View">
                                    <EyeRegular style={{ fontSize: 16, color: 'white' }} />
                                  </IconButton>
                                  <a href={req.fileUrl} download target="_blank" rel="noopener noreferrer">
                                    <IconButton className={styles.btn}>
                                      <ArrowDownloadFilled className={styles.downloadIcon} />
                                    </IconButton>
                                  </a>
                                </div>
                              )}
                            </TableCell>
                          </TableRow>
                        ))}
                    </TableBody>
                  </>
                )}
              </Table>
            </>
          )}

          {isModalOpen && selectedDocumentUrl && (
            <Modal isOpen={isModalOpen} onDismiss={closeModal} isBlocking={false} containerClassName={styles.customModal} >
              <div className={styles.modalHeader}>
                <span>Document Viewer</span>
                <IconButton iconProps={{ iconName: 'Cancel' }} onClick={closeModal} />
              </div>
              <div className={styles.modalContent}>
                <iframe src={`${selectedDocumentUrl}#toolbar=1&navpanes=0&scrollbar=1`} title="Document" />
              </div>
            </Modal>
          )}

          {isModalDetailsOpen && groupedRequestsToShow.length > 0 && (
            <Modal isOpen={isModalDetailsOpen} onDismiss={closeViewDetailsModal} isBlocking={false} containerClassName={styles.customModal} >
              <div className={styles.modalHeader}>
                <span>Request Details</span>
                <IconButton iconProps={{ iconName: 'Cancel' }} onClick={closeViewDetailsModal} />
              </div>
              <div className={styles.modalContent}>
                <Table {...keyboardNavAttr} role="grid" aria-label="Grouped signature requests" style={{ minWidth: "620px" }} >
                  <TableHeader style={{ borderBottom: '2px solid #ccc' }}>
                    <TableRow>
                      <TableHeaderCell>File</TableHeaderCell>
                      <TableHeaderCell>Signer</TableHeaderCell>
                      <TableHeaderCell>Date</TableHeaderCell>
                      <TableHeaderCell>Status</TableHeaderCell>
                      <TableHeaderCell>Signature Mode</TableHeaderCell>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {groupedRequestsToShow.map((item, index) => (
                      <TableRow key={index} className={styles.tableRow} onMouseEnter={(e) => e.currentTarget.style.background = "#f9f9f9"} onMouseLeave={(e) => e.currentTarget.style.background = "white"} >
                        <TableCell>
                          <span className={styles.fileNameLink} title={item.fileName} onClick={() => handleView(item.fileUrl)}>
                            {item.fileName.length > 19 ? item.fileName.substring(0, 19) + "..." : item.fileName}
                          </span>
                        </TableCell>
                        <TableCell>
                          {item.signerName}
                        </TableCell>
                        <TableCell>
                          {new Date(item.created).toLocaleString()}
                        </TableCell>
                        <TableCell>
                          <TableCellLayout
                            media={
                              item.approvalDecision?.toLowerCase() === "approve" ? (
                                <span style={{ marginRight: "3px" }}>✔️</span>
                              ) : item.approvalDecision?.toLowerCase() === "reject" ? (
                                <span style={{ marginRight: "3px" }}>❌</span>
                              ) : (
                                <span style={{ marginRight: "3px" }}>⏳</span>
                              )
                            }
                          >
                            {item.approvalDecision}
                          </TableCellLayout>
                        </TableCell>
                        <TableCell>
                          {item.typeSignature}
                        </TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </div>
            </Modal>
          )}
          <div style={{ marginTop: "1rem", display: "flex", justifyContent: "center", gap: "1rem" }}>
            <Button onClick={goToPreviousPage} disabled={currentPage === 1}> Previous</Button>
            <span>Page {currentPage} / {totalPages}</span>
            <Button onClick={goToNextPage} disabled={currentPage === totalPages}> Next </Button>
          </div>
        </div>
      </div>
    </>
  );
};

export default SignatureRequestOverview;
