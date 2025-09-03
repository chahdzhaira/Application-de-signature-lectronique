import * as React from 'react';
import styles from './Contact.module.scss';
import { IContactProps } from './IContactProps';
import { MessageBar, MessageBarType, PrimaryButton } from '@fluentui/react';

const Contact: React.FC<IContactProps> = (props) => {
  const [formData, setFormData] = React.useState({
    name: '',
    email: '',
    phone: '',
    message: ''
  });
  const [globalError, setGlobalError] = React.useState<string | null>(null);
  const [successMessage, setSuccessMessage] = React.useState<string | null>(null);
  const errorRef = React.useRef<HTMLDivElement | null>(null);


  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    const { name, value } = e.target;
    setFormData(prev => ({
      ...prev,
      [name]: value
    }));
  };

  const handleSubmit = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();

    setGlobalError(null);
    setSuccessMessage(null);

    try {
      const response = await fetch("https://prod-54.northeurope.logic.azure.com:443/workflows/464cc963e2254ea3984cbd6c3da447be/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=huJdpU7o-vTFjKYVwEARvguUPGpAhvkBhqA8X9OcVQo", {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(formData)
      });

      if (response.ok) {
        setSuccessMessage("Message sent successfully!");
        setFormData({ name: '', email: '', phone: '', message: '' });
        errorRef.current?.scrollIntoView({ behavior: 'smooth' });
      } else {
        setGlobalError("Failed to send message.");
        errorRef.current?.scrollIntoView({ behavior: 'smooth' });
      }
    } catch (error) {
      console.error("Error sending form data:", error);
      setGlobalError("An error occurred while sending the message.");
      errorRef.current?.scrollIntoView({ behavior: 'smooth' });
    }
  };

  return (
    <div className={styles.contactForm}>
      <div ref={errorRef}>
        {globalError && (<MessageBar className={styles.messageBar} messageBarType={MessageBarType.error} isMultiline={false} onDismiss={() => setGlobalError(null)} dismissButtonAriaLabel="Close" > {globalError} </MessageBar>)}
        {successMessage && (<MessageBar className={styles.messageBar} messageBarType={MessageBarType.success} isMultiline={false} onDismiss={() => setSuccessMessage(null)}> {successMessage} </MessageBar>)}
      </div>

      <div className={styles.container}>
        <div className={styles.formSection}>
          <div className={styles.header}>
            <span className={styles.contactTag}>CONTACT US</span>
            <h2 className={styles.title}> <br /> Contact us to find out more. </h2>
          </div>

          <form className={styles.form} onSubmit={handleSubmit}>
            <div className={styles.inputGroup}>
              <input type="text" name="name" placeholder="Your Name*" className={`${styles.input} ${styles.fullWidth}`} value={formData.name} onChange={handleInputChange} required />
            </div>
            <div className={styles.inputRow}>
              <input type="email" name="email" placeholder="Email*" className={`${styles.input} ${styles.halfWidth}`} value={formData.email} onChange={handleInputChange} required />
              <input type="tel" name="phone" placeholder="Phone" className={`${styles.input} ${styles.halfWidth}`} value={formData.phone} onChange={handleInputChange} />
            </div>
            <div className={styles.inputGroup}>
              <textarea name="message" placeholder="Message*" className={styles.textarea} rows={5} value={formData.message} onChange={handleInputChange} required />
            </div>
            <PrimaryButton text="Send" className={styles.btn} type="submit" />
          </form>
        </div>

        <div className={styles.contactInfo}>
          <div className={styles.contactItem}>
            <div className={styles.iconWrapper}>
              <svg className={styles.icon} viewBox="0 0 24 24" fill="none" stroke="currentColor">
                <path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z" />
              </svg>
            </div>
            <div className={styles.contactDetails}>
              <h3 className={styles.contactTitle}>Phone</h3>
              <p className={styles.contactText}>
                (+216) 54 704 630
              </p>
            </div>
          </div>

          <div className={styles.contactItem}>
            <div className={styles.iconWrapper}>
              <svg className={styles.icon} viewBox="0 0 24 24" fill="none" stroke="currentColor">
                <path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z" />
                <polyline points="22,6 12,13 2,6" />
              </svg>
            </div>
            <div className={styles.contactDetails}>
              <h3 className={styles.contactTitle}>E-Mail</h3>
              <p className={styles.contactText}>
                Chahd@48c2lx.onmicrosoft.com
              </p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};


export default Contact;