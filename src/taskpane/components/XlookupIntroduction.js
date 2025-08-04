/* global Excel */

import React from "react";
import { useTranslation } from "react-i18next";
import styled from "styled-components";
import logo from "../../../assets/logo_excel_bites.png";
import { setRangeBold, clearRange, autofitColumns, setColumnWidth, setRangeCenter, setRangeRight, setFontSize, setRangeItalic } from "../excelFormatters";

const StyledContainer = styled.div`
  text-align: center;
  padding: 10px; /* Further reduced padding */
  background-color: white;
  font-family: Arial, sans-serif;
`;

const StyledLogo = styled.img`
  max-width: 100px; /* Even smaller logo */
  margin-bottom: 10px; /* Reduced margin */
`;

const StyledTitle = styled.h1`
  color: #217346;
  font-size: 22px; /* Further smaller font size */
  margin-bottom: 8px; /* Reduced margin */
`;

const StyledParagraph = styled.p`
  font-size: 13px; /* Further smaller font size */
  line-height: 1.3;
  margin-bottom: 10px; /* Reduced margin */
`;

const StyledAdvantagesContainer = styled.div`
  margin: 10px 0; /* Reduced margin */
  padding: 8px; /* Reduced padding */
  border: 1px solid #e0e0e0;
  border-radius: 8px;
  background-color: #f9f9f9;
  text-align: left;
`;

const StyledAdvantagesTitle = styled.h2`
  color: #217346;
  font-size: 16px; /* Further smaller font size */
  margin-bottom: 6px; /* Reduced margin */
  text-align: center;
`;

const StyledAdvantagesList = styled.ul`
  list-style: none;
  padding: 0;
  margin: 0;
`;

const StyledAdvantageItem = styled.li`
  font-size: 13px; /* Further smaller font size */
  margin-bottom: 4px; /* Reduced margin */
  display: flex;
  align-items: center;
`;

const CheckMark = styled.span`
  color: #217346;
  font-size: 16px; /* Further smaller font size */
  margin-right: 6px; /* Reduced margin */
`;

const StyledButton = styled.button`
  background-color: #217346;
  color: white;
  border: none;
  padding: 6px 12px; /* Further reduced padding */
  font-size: 14px; /* Further smaller font size */
  cursor: pointer;
  border-radius: 5px;
  margin: 2px; /* Reduced margin */

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

const StyledResetButton = styled(StyledButton)`
  background-color: #f44336; /* Red color for reset */

  &:hover {
    background-color: #d32f2f;
  }
`;

const ButtonContainer = styled.div`
  margin-top: 6px; /* Reduced margin */
  display: flex;
  justify-content: center;
  gap: 8px; /* Reduced space between buttons */
`;

const XlookupIntroduction = ({ goToNextStep, resetLesson }) => {
  const { t } = useTranslation();

  const handlePrepareData = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // Clear existing data in a larger range to ensure old data is removed
        clearRange(context, "A:G");

        // Insert title
        sheet.getRange("A1").values = [[t("excelbites_title")]];
        setRangeBold(context, "A1");
        setFontSize(context, "A1", 18);

        // Insert formula structure
        sheet.getRange("A2").values = [[t("xlookup_formula_structure")]];
        setFontSize(context, "A2", 13);
        setRangeItalic(context, "A2");

        // Insert headers starting from row 5
        const headers = [[t("product_id_header"), t("product_header"), t("price_header")]];
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
        sheet.getRange("E5").values = [[t("search_id_label")]];
        setRangeBold(context, "E5");
        sheet.getRange("F5").values = [[104]]; // Default search ID
        setRangeCenter(context, "F5:G9");

        sheet.getRange("E7").values = [[t("simple_formula_label")]];
        setRangeBold(context, "E7");
        sheet.getRange("F7").values = [[""]]; // Empty cell for result

        setRangeRight(context, "E5:E16");

        setColumnWidth(context, ["A", "C", "D", "F", "G", "I", "J"], 75);
        setColumnWidth(context, ["D"], 30);
        setColumnWidth(context, ["E"], 130);
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
            <CheckMark>{t("checkmark")}</CheckMark> {t("advantage1")}
          </StyledAdvantageItem>
          <StyledAdvantageItem>
            <CheckMark>{t("checkmark")}</CheckMark> {t("advantage2")}
          </StyledAdvantageItem>
          <StyledAdvantageItem>
            <CheckMark>{t("checkmark")}</CheckMark> {t("advantage3")}
          </StyledAdvantageItem>
          <StyledAdvantageItem>
            <CheckMark>{t("checkmark")}</CheckMark> {t("advantage4")}
          </StyledAdvantageItem>
        </StyledAdvantagesList>
      </StyledAdvantagesContainer>

      <ButtonContainer>
        <StyledButton onClick={handlePrepareData}>{t("prepare_data_button")}</StyledButton>
        <StyledResetButton onClick={resetLesson}>{t("reset_lesson_button")}</StyledResetButton>

      </ButtonContainer>
    </StyledContainer>
  );
};

export default XlookupIntroduction;
