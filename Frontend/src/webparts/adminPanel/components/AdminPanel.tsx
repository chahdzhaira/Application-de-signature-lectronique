import * as React from 'react';
import { sp } from '@pnp/sp/presets/all';
type IPrincipalInfo = any;
import { DetailsList, PrimaryButton, Toggle, Label, MessageBar, MessageBarType, Dropdown } from '@fluentui/react';
import { TagPicker, TagPickerControl, TagPickerGroup, TagPickerOption, TagPickerInput, TagPickerList, Tag, Avatar, Field } from "@fluentui/react-components";
import { IAdminPanelProps } from './IAdminPanelProps';
import styles from './AdminPanel.module.scss';


const AdminPanel: React.FC<IAdminPanelProps> = (props) => {
  const [errorMessage, setErrorMessage] = React.useState<string | null>(null);
  const [successMessage, setSuccessMessage] = React.useState<string | null>(null);
  const [users, setUsers] = React.useState<any[]>([]);
  const [selectedUser, setSelectedUser] = React.useState<IPrincipalInfo | null>(null);
  const [role, setRole] = React.useState<string>('user');
  const [isActive, setIsActive] = React.useState<boolean>(true);
  const [emailOptions, setEmailOptions] = React.useState<any[]>([]);


  React.useEffect(() => {
    const init = async () => {
      try {
        await fetchUsers();
        await loadUsers();
      } catch (err) {
        setErrorMessage("Error loading data.");
      }
    };
    init();
  }, []);

  const fetchUsers = async () => {
    const client = await props.context.msGraphClientFactory.getClient("3");
    const response = await client.api('/users?$select=mail,displayName,jobTitle').top(50).get();
    const users = response.value
      .filter((user: any) => user.mail)
      .map((user: any) => ({
        email: user.mail,
        displayName: user.displayName,
        jobTitle: user.jobTitle || 'Not specified'
      }));
    setEmailOptions(users);
  };

  const loadUsers = async () => {
    const items = await sp.web.lists.getByTitle("UsersList")
      .items.expand("User")
      .select("Id", "Role", "Actif", "User/Title", "User/EMail")
      .get();
    setUsers(items);
  };

  const addUser = async () => {
    if (!selectedUser) {
      setErrorMessage("Please select a user.");
      return;
    }

    try {
      await sp.web.ensureUser(selectedUser.Email);
      const userInfo = await sp.web.siteUsers.getByEmail(selectedUser.Email).get();
      await sp.web.lists.getByTitle("UsersList").items.add({
        Role: role,
        Actif: isActive,
        UserId: userInfo.Id
      });
      setSelectedUser(null);
      setRole("user");
      setIsActive(true);
      loadUsers();
      setSuccessMessage("User added successfully !");
    } catch (error) {
      console.error("Error adding user :", error);
      setErrorMessage("Error adding user.");
    }
  };

  return (
    <div style={{padding: 20}}>
      {errorMessage && (<MessageBar className={styles.messageBar} messageBarType={MessageBarType.error} onDismiss={() => setErrorMessage(null)} > {errorMessage} </MessageBar>)}
      {successMessage && (<MessageBar className={styles.messageBar} messageBarType={MessageBarType.success} onDismiss={() => setSuccessMessage(null)}>{successMessage} </MessageBar>)}

      <h1 className={styles.sectionTitle}>User management</h1>
      <div className={styles.panelContent}>

        <div className={styles.leftCol}>
          <Field label="User" required>
            <TagPicker
              onOptionSelect={(_, data) => {
                const selected = emailOptions.find(u => u.email === data.value);
                if (selected) {
                  setSelectedUser({ Email: selected.email, Title: selected.displayName });
                }
              }}
              selectedOptions={selectedUser ? [selectedUser.Email] : []}
            >
              <TagPickerControl>
                {selectedUser && (
                  <TagPickerGroup aria-label="Selected Email">
                    <Tag
                      key={selectedUser.Email}
                      shape="rounded"
                      media={<Avatar name={selectedUser.Email} color="colorful" />}
                      value={selectedUser.Email}
                    >
                      {selectedUser.Email}
                    </Tag>
                  </TagPickerGroup>
                )}
                <TagPickerInput aria-label="Select a user" />
              </TagPickerControl>

              <TagPickerList className={styles.tagPickerListFix}>
                {emailOptions.map((user) => (
                  <TagPickerOption
                    key={user.email}
                    value={user.email}
                    media={<Avatar name={user.email} color="colorful" />}
                    secondaryContent={`${user.email} • ${user.jobTitle}`}
                  >
                    {user.displayName}
                  </TagPickerOption>
                ))}
              </TagPickerList>
            </TagPicker>
          </Field>

          <Dropdown label="Role" options={[{ key: "admin", text: "Admin" }, { key: "user", text: "User" }]} selectedKey={role} onChange={(e, option) => setRole(option!.key as string)} styles={{ dropdown: { minWidth: 150, borderRadius: 4, border: "1px solid #ccc" } }} />
          <Toggle label="Actif" checked={isActive} onChange={(e, checked) => setIsActive(!!checked)} className={styles.fieldMargin}/>
          <PrimaryButton text="Add user" onClick={addUser} className={styles.btn} />
        </div>

        <div className={styles.divider}/>
        <div className={styles.rightCol}>
          <Label style={{ marginBottom: 10 }}>List of users</Label>
          <DetailsList
            items={users.map(user => ({
              key: user.Id,
              Name: user.User?.Title,
              Email: user.User?.EMail,
              Role: user.Role,
              Actif: user.Actif ? 'Oui' : 'Non'
            }))}
            columns={[
              { key: 'Name', name: 'Name', fieldName: 'Name', minWidth: 100 },
              { key: 'Email', name: 'Email', fieldName: 'Email', minWidth: 150 },
              { key: 'Role', name: 'Role', fieldName: 'Role', minWidth: 80 },
              { key: 'Actif', name: 'Actif', fieldName: 'Actif', minWidth: 60 },
            ]}
          />
        </div>
      </div>
    </div>
  );

};

export default AdminPanel;