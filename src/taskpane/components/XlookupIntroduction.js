/* global Excel */

import React from "react";
import { useTranslation } from "react-i18next";
import styled from "styled-components";
import logo from "../../../assets/logo_excel_bites.png";
import { setRangeBold, clearRange, autofitColumns } from "../excelFormatters";

const StyledContainer = styled.div`
  text-align: center;
  padding: 20px;
  background-color: white; /* Fondo blanco */
  font-family: Arial, sans-serif; /* Fuente más adaptada a un tutorial */
`;

const StyledLogo = styled.img`
  max-width: 150px;
  margin-bottom: 20px;
`;

const StyledTitle = styled.h1`
  color: #217346; /* Excel green */
  font-size: 28px;
  margin-bottom: 15px;
`;

const StyledParagraph = styled.p`
  font-size: 16px;
  line-height: 1.5;
  margin-bottom: 20px;
`;

const StyledButton = styled.button`
  background-color: #217346; /* Excel green */
  color: white;
  border: none;
  padding: 10px 20px;
  font-size: 18px;
  cursor: pointer;
  border-radius: 5px;
  margin: 5px; /* Add some margin for spacing between buttons */

  &:hover {
    background-color: #1a5c38;
  }
`;

const XlookupIntroduction = () => {
  const { t } = useTranslation();

  const handlePrepareData = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // Clear existing data in a larger range to ensure old data is removed
        clearRange(context, "A:G");

        // Insert headers starting from row 5
        const headers = [["ID Producto", "Producto", "Precio"]];
        sheet.getRange("A5:C5").values = headers;
        setRangeBold(context, "A5:C5");

        // Insert example data starting from row 6
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
        sheet.getRange("A6:C" + (data.length + 5)).values = data;

        // Set up search ID and result cells below the data
        const searchRow = data.length + 7; // 2 rows below the end of data
        sheet.getRange("E" + searchRow).values = [["Buscar ID:"]];
        setRangeBold(context, "E" + searchRow);
        sheet.getRange("F" + searchRow).values = [[103]]; // Default search ID

        const resultRow = data.length + 8; // 1 row below search ID
        sheet.getRange("E" + resultRow).values = [["Resultado:"]];
        setRangeBold(context, "E" + resultRow);

        await autofitColumns(context, sheet.getUsedRange());
        await context.sync();
      });
    } catch (error) {
      console.error("Error preparing data:", error);
    }
  };

  return (
    <StyledContainer>
      <StyledLogo src={logo} alt="ExcelBites Logo" />
      <StyledTitle>{t("introduction_title")}</StyledTitle>
      <StyledParagraph dangerouslySetInnerHTML={{ __html: t("introduction_text") }} />
      <StyledButton onClick={handlePrepareData}>{t("prepare_data_button")}</StyledButton>
    </StyledContainer>
  );
};

export default XlookupIntroduction;
