import * as React from "react";
import styles from "./Footer.module.scss";
import { Mail20Filled, Globe20Filled, Location20Filled, Call20Filled, } from "@fluentui/react-icons";

interface IFooterProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}

const Footer: React.FC<IFooterProps> = () => {
  return (
    <footer className={styles.footer}>
      <div className={styles.logosContainer}>
        <img src={require("../assets/logo_alight.png")} alt="Alight Logo" className={styles.logoCompany} />
        <img src={require("../assets/QuickSign_Logo.png")} alt="QuickSign Logo" className={styles.logoApp} />
      </div>
      <div className={styles.socialIcons}>
        <a href="https://alightmea.com" target="_blank" rel="noopener noreferrer">
          <Globe20Filled className={styles.icon} />
        </a>
        <a href="mailto:info@Alight.eu" target="_blank" rel="noopener noreferrer">
          <Mail20Filled className={styles.icon} />
        </a>
        <a href="tel:+21671948549" target="_blank" rel="noopener noreferrer">
          <Call20Filled className={styles.icon} />
        </a>
        <a href="https://maps.app.goo.gl/rQrLAFvWkYYBuQPk9" target="_blank" rel="noopener noreferrer">
          <Location20Filled className={styles.icon} />
        </a>
      </div>

      <nav className={styles.navbar}>
        <a href="#">Home</a>
        <a href="#">About</a>
        <a href="#">Services</a>
        <a href="#">Contact</a>
      </nav>

      <p className={styles.copyright}>
        &copy; {new Date().getFullYear()} QuickSign. All rights reserved.
      </p>
    </footer>
  );
};

export default Footer;
