import * as React from 'react';
import styles from './Navbar.module.scss';
import { useMsal } from "@azure/msal-react";
import { Avatar } from "@fluentui/react-components";
import { INavbarProps } from './INavbarProps';
import { checkUserRole } from '../../../services/permissions';

const Navbar: React.FC<INavbarProps> = (props) => {
  const [isMenuOpen, setIsMenuOpen] = React.useState(false);
  const { accounts } = useMsal();
  const username = props.context.pageContext.user.displayName;

  const [role, setRole] = React.useState<string | null>(null);
  const [loadingRole, setLoadingRole] = React.useState(true);

  React.useEffect(() => {
  const cachedRole = sessionStorage.getItem("userRole");

  if (cachedRole) {
    setRole(cachedRole);
    setLoadingRole(false);
    console.log(loadingRole)
  } else {
    const fetchUserRole = async () => {
      try {
        const { role } = await checkUserRole();
        setRole(role);
        if (role) {
          sessionStorage.setItem("userRole", role);
        }
      } catch {
        setRole(null);
      } finally {
        setLoadingRole(false);
      }
    };
    fetchUserRole();
  }
}, []);

  const toggleMenu = () => setIsMenuOpen(!isMenuOpen);

  const getActiveLink = () => {
    const url = window.location.href;
    if (url.includes("HomeApp.aspx")) return "Home";
    if (url.includes("Demande-Signature.aspx")) return "SignatureRequest";
    if (url.includes("SignatureRequestOverview.aspx")) return "Agreements";
    if (url.includes("Insights.aspx")) return "Insights";
    if (url.includes("AdminPanel.aspx")) return "AdminPanel";
    if (url.includes("Contact.aspx")) return "Contact";
    return "";
  };

  const activeLink = getActiveLink();

  return (
    <nav className={styles.navbar}>
      <div className={styles.container}>
        <div className={styles.brand}>
          <img src={require("../assets/ESign_Logo.png")} alt="logo" width={80} height={80} />
          <h2 className={styles.NameApp}>QuickSign</h2>
        </div>
        <button className={styles.toggleButton} onClick={toggleMenu}>
          ☰
        </button>
        <div className={`${styles.navLinks} ${isMenuOpen ? styles.show : ''}`}>
          <ul>
            <li><a href="https://48c2lx.sharepoint.com/sites/App_signature_electronique/SitePages/HomeApp.aspx" className={activeLink === "Home" ? styles.active : ""} >Home</a></li>
            <li><a href="https://48c2lx.sharepoint.com/sites/App_signature_electronique/SitePages/Demande-Signature.aspx" className={activeLink === "SignatureRequest" ? styles.active : ""} >Signature Request</a></li>
            <li><a href="https://48c2lx.sharepoint.com/sites/App_signature_electronique/SitePages/SignatureRequestOverview.aspx" className={activeLink === "Agreements" ? styles.active : ""} >Agreements</a></li>
            {role && role.toLowerCase().trim() === 'admin' && (
              <li><a href="https://48c2lx.sharepoint.com/sites/App_signature_electronique/SitePages/Insights.aspx" className={activeLink === "Insights" ? styles.active : ""} >Insights</a></li>
            )}
            {role && role.toLowerCase().trim() === 'admin' && (
              <li><a href="https://48c2lx.sharepoint.com/sites/App_signature_electronique/SitePages/AdminPanel.aspx" className={activeLink === "AdminPanel" ? styles.active : ""} >AdminPanel</a></li>
            )}


            <li><a href="https://48c2lx.sharepoint.com/sites/App_signature_electronique/SitePages/Contact.aspx" className={activeLink === "Contact" ? styles.active : ""} >Contact</a></li>
          </ul>
        </div>
        <div className={styles.profileMenu}>
          <div className={styles.profileContainer}>
            <span className={styles.username}>{username}</span>
            <button className={styles.profileButton} aria-label="Open Profile Menu">
              {((accounts[0]?.idTokenClaims as any)?.picture) ? (
                <img
                  src={(accounts[0]?.idTokenClaims as any).picture as string}
                  alt="User"
                  className={styles.profileImage}
                />
              ) : (
                <Avatar
                  name={username}
                  size={32}
                  className={styles.profileAvatar}
                  color="colorful"
                />
              )}
            </button>
          </div>
        </div>
      </div>
    </nav>
  );
};

export default Navbar;
