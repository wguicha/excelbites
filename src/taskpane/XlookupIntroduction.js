import React from "react";
import { useTranslation } from "react-i18next";
import logo from "../../assets/logo_excel_bites.png";

const XlookupIntroduction = () => {
  const { t } = useTranslation();

  const handlePrepareData = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // Clear existing data
        sheet.getRange("A:C").clear();

        // Insert headers
        const headers = [["ID Producto", "Producto", "Precio"]];
        sheet.getRange("A1:C1").values = headers;
        sheet.getRange("A1:C1").format.font.bold = true;

        // Insert example data
        const data = [
          [101, "Laptop", 1200],
          [102, "Mouse", 25],
          [103, "Keyboard", 75],
          [104, "Monitor", 300],
          [105, "Webcam", 50],
          [106, "Microphone", 80],
          [107, "Headphones", 150],
          [108, "Printer", 200],
          [109, "Scanner", 100],
          [110, "External Hard Drive", 90],
        ];
        sheet.getRange("A2:C" + (data.length + 1)).values = data;

        // Set up search ID and result cells
        sheet.getRange("E1").values = [["Buscar ID:"]];
        sheet.getRange("E1").format.font.bold = true;
        sheet.getRange("F1").values = [[103]]; // Default search ID
        sheet.getRange("E2").values = [["Resultado:"]];
        sheet.getRange("E2").format.font.bold = true;

        sheet.getUsedRange().format.autofitColumns();
        await context.sync();
      });
    } catch (error) {
      console.error("Error preparing data:", error);
    }
  };

  return (
    <div className="ms-welcome">
      <img src={logo} alt="ExcelBites Logo" className="logo" />
      <h1 className="ms-font-xl ms-fontWeight-semilight ms-fontColor-neutralPrimary">{t("introduction_title")}</h1>
      <p className="ms-font-m ms-fontColor-neutralSecondary" dangerouslySetInnerHTML={{ __html: t("introduction_text") }}></p>
      <button className="ms-Button ms-Button--primary" onClick={handlePrepareData}>
        <span className="ms-Button-label">{t("prepare_data_button")}</span>
      </button>
    </div>
  );
};

export default XlookupIntroduction;
