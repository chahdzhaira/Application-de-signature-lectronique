import * as React from 'react';
import { useState, useEffect } from 'react';
import AdminPanel from './AdminPanel'; 
import { PrimaryButton, Spinner } from '@fluentui/react';
import styles from './AdminPanel.module.scss';
import { checkUserRole } from '../../../services/permissions';


interface IAdminPanelWrapperProps {
  context: any;
}

const AdminPanelWrapper: React.FC<IAdminPanelWrapperProps> = ({ context }) => {
  const [role, setRole] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const fetchRole = async () => {
      try {
        const { isAllowed, role } = await checkUserRole();
        if (isAllowed) {
          setRole(role);
        } else {
          setRole(null);
        }
      } catch (error) {
        console.error("Loading permissions...Error while verifying role :", error);
        setRole(null);
      }
      setLoading(false);
    };

    fetchRole();
  }, []);

  if (loading) {
    return <Spinner label="Loading permissions..." />;
  }

  if (role !== 'admin') {
    return (
      <div className={styles.accessDenied}>
        <h1>404</h1>
        <h2>Access Denied</h2>
        <p>Sorry, but you don't have permission to access this page.</p>
        <PrimaryButton onClick={() => window.history.back()} className={styles.btn}>
          ⬅ Go Back
        </PrimaryButton>
        <div>
          <img src="https://cdn-icons-png.flaticon.com/512/2748/2748558.png" alt="Access denied" width="200" />
          </div>
      </div>
    );
  }

  return <AdminPanel context={context} />;
};

export default AdminPanelWrapper;
