/* global Excel */

import React from "react";
import { useTranslation } from "react-i18next";
import styled from "styled-components";
import logo from "../../../assets/logo_excel_bites.png";
import { setRangeBold, clearRange, autofitColumns, setColumnWidth, setRangeCenter, setRangeRight, setFontSize, setRangeItalic } from "../excelFormatters";

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

const StyledAdvantagesContainer = styled.div`
  margin: 20px 0;
  padding: 15px;
  border: 1px solid #e0e0e0;
  border-radius: 8px;
  background-color: #f9f9f9;
  text-align: left;
`;

const StyledAdvantagesTitle = styled.h2`
  color: #217346;
  font-size: 20px;
  margin-bottom: 10px;
  text-align: center;
`;

const StyledAdvantagesList = styled.ul`
  list-style: none;
  padding: 0;
  margin: 0;
`;

const StyledAdvantageItem = styled.li`
  font-size: 16px;
  margin-bottom: 8px;
  display: flex;
  align-items: center;
`;

const CheckMark = styled.span`
  color: #217346;
  font-size: 20px;
  margin-right: 10px;
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

const StyledNavButton = styled(StyledButton)`
  background-color: #a9a9a9;

  &:hover {
    background-color: #808080;
  }
`;

const ButtonContainer = styled.div`
  margin-top: 10px;
`;

const XlookupIntroduction = ({ goToNextStep }) => {
  const { t } = useTranslation();

  const handlePrepareData = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // Clear existing data in a larger range to ensure old data is removed
        clearRange(context, "A:G");

        // Insert title
        sheet.getRange("A1").values = [["ExcelBites: La poderosa BUSCARX"]];
        setRangeBold(context, "A1");
        setFontSize(context, "A1", 18);

        // Insert formula structure
        sheet.getRange("A2").values = [["'=BUSCARX(valor_buscado, matriz_buscada, matriz_devuelta, [si_no_se_encuentra], [modo_de_coincidencia], [modo_de_búsqueda])"]];
        setFontSize(context, "A2", 15);
        setRangeItalic(context, "A2");

        // Insert headers starting from row 5
        const headers = [["ID Producto", "Producto", "Precio"]];
        sheet.getRange("A5:C5").values = headers;
        setRangeBold(context, "A5:C5");
        setRangeCenter(context, "A5:C5");

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
        setRangeCenter(context, "A6:A" + (data.length + 5));

        // Set up search ID and result cells
        sheet.getRange("E5").values = [["Buscar ID:"]];
        setRangeBold(context, "E5");
        sheet.getRange("F5").values = [[104]]; // Default search ID
        setRangeCenter(context, "F5");

        sheet.getRange("E7").values = [["Formula simple:"]];
        setRangeBold(context, "E7");
        sheet.getRange("F7").values = [[""]]; // Empty cell for result

        sheet.getRange("E9").values = [["Formula Múltiple:"]];
        setRangeBold(context, "E9");

        setRangeRight(context, "E5:E15");

        setColumnWidth(context, ["A", "C", "D", "E", "F"], 75);
        setColumnWidth(context, ["B"], 100);
        //await autofitColumns(context, sheet.getUsedRange());

        // Set cursor to F7
        sheet.getRange("F7").select();

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

      <StyledAdvantagesContainer>
        <StyledAdvantagesTitle>{t("advantages_title")}</StyledAdvantagesTitle>
        <StyledAdvantagesList>
          <StyledAdvantageItem>
            <CheckMark>✔</CheckMark> {t("advantage1")}
          </StyledAdvantageItem>
          <StyledAdvantageItem>
            <CheckMark>✔</CheckMark> {t("advantage2")}
          </StyledAdvantageItem>
          <StyledAdvantageItem>
            <CheckMark>✔</CheckMark> {t("advantage3")}
          </StyledAdvantageItem>
          <StyledAdvantageItem>
            <CheckMark>✔</CheckMark> {t("advantage4")}
          </StyledAdvantageItem>
        </StyledAdvantagesList>
      </StyledAdvantagesContainer>

      <StyledButton onClick={handlePrepareData}>{t("prepare_data_button")}</StyledButton>
      <ButtonContainer>
        <StyledNavButton onClick={() => { console.log("Next button clicked in XlookupIntroduction"); goToNextStep(); }}>&#9654;</StyledNavButton>
      </ButtonContainer>
    </StyledContainer>
  );
};

export default XlookupIntroduction;
