import * as React from 'react';
import { useState, useEffect } from 'react';
import { sp } from '@pnp/sp/presets/all';
import { PrimaryButton, Spinner } from '@fluentui/react';
import Insights from './Insights';
import styles from './Insights.module.scss';


interface IInsightsWrapperProps {
    context: any;
}

const InsightsWrapper: React.FC<IInsightsWrapperProps> = ({ context }) => {
    const [role, setRole] = useState<string | null>(null);
    const [loading, setLoading] = useState(true);

    useEffect(() => {
        const checkUserRole = async () => {
            try {
                const currentUser = await sp.web.currentUser.get();
                const userEmail = currentUser.Email;

                const userItems = await sp.web.lists.getByTitle("UsersList")
                    .items
                    .filter(`User/EMail eq '${userEmail}' and Actif eq 1`)
                    .expand("User")
                    .select("Role", "User/Title")
                    .top(1)
                    .get();

                if (userItems.length === 0) {
                    setRole(null);
                } else {
                    setRole(userItems[0].Role);
                }
            } catch (error) {
                console.error("Error while verifying role:", error);
                setRole(null);
            }
            setLoading(false);
        };

        checkUserRole();
    }, []);

    if (loading) {
        return <Spinner label="Checking permissions..." />;
    }

    if (role !== 'admin') {
        return (
            <div style={{ textAlign: "center", marginTop: "5px", color: '#183060' }}>
                <h1 style={{ fontSize: "72px", margin: "0" }}>404</h1>
                <h2>Access Denied</h2>
                <p>Sorry, but you don't have permission to access this page.</p>
                <PrimaryButton onClick={() => window.history.back()} className={styles.btn}>
                    ⬅ Go Back
                </PrimaryButton>
                <div style={{ marginTop: "30px" }}>
                    <img src="https://cdn-icons-png.flaticon.com/512/2748/2748558.png" alt="Access denied" width="200" />
                </div>
            </div>
        );
    }

    return <Insights />;
};

export default InsightsWrapper;
