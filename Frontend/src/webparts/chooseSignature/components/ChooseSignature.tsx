import * as React from "react";
import { Tab, TabList } from "@fluentui/react-components";
import { SelectTabData, SelectTabEvent, TabValue, makeStyles, shorthands } from "@fluentui/react-components";
import { ImageRegular, DrawImageRegular } from "@fluentui/react-icons";

import PenSignature from "../../penSignature/components/PenSignature";
import UploadSignature from "../../uploadSignature/components/UploadSignature";
import { IChooseSignatureProps } from "./IChooseSignatureProps";

const useStyles = makeStyles({
  tab: {
    color: "#545454",
    borderBottom: "2px solid transparent",
    ...shorthands.padding("8px", "12px"),
    cursor: "pointer",
    ":hover": {
      color: "#ff7811"
    }
  },
  tabActive: {
    color: "#ff7811",
    fontWeight: "bold",
    ":global(.fui-Tab__content)": {
      color: "#ff7811  !important",
    }
  },
  tabList: {
    display: "flex",
    gap: "12px",
    marginBottom: "16px"
  }
  
});

export const ChooseSignature : React.FC<IChooseSignatureProps> = ({context, onClose }) => {
  const [selectedValue, setSelectedValue] = React.useState<TabValue>("upload");

  const styles = useStyles();

  const onTabSelect = (event: SelectTabEvent, data: SelectTabData) => {
    setSelectedValue(data.value);
  };
  const getIconColor = (tab: TabValue) =>
    selectedValue === tab ? "#ff7811" : "#545454";
  return (
    <div >
      <TabList selectedValue={selectedValue} onTabSelect={onTabSelect} className={styles.tabList}>
        <Tab
          id="UploadSignature"
          value="upload"
          icon={<ImageRegular style={{ color: getIconColor("upload") }} />}
          className={selectedValue === "upload" ? `${styles.tab} ${styles.tabActive}` : styles.tab}
        >
          <span style={{ color: selectedValue === "upload" ? "#ff7811" : "#545454", fontWeight: selectedValue === "upload" ? "bold" : "normal" }}
          >
            Upload Signature
          </span>
        </Tab>

        <Tab
          id="PenSignature"
          value="pen"
          icon={<DrawImageRegular style={{ color: getIconColor("pen") }} />}
          className={selectedValue === "pen" ? `${styles.tab} ${styles.tabActive}` : styles.tab}
        >
          <span style={{ color: selectedValue === "pen" ? "#ff7811" : "#545454", fontWeight: selectedValue === "pen" ? "bold" : "normal" }}>
            Pen Signature
          </span>
        </Tab>

      </TabList>

      <div >
        {selectedValue === "upload" &&  <UploadSignature context={context}  onClose={onClose}  />}
        {selectedValue === "pen" && <PenSignature context={context}  onClose={onClose} />}
      </div>
    </div>
  );
};


export default ChooseSignature;
