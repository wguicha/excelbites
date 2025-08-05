/* global Excel */

import React from 'react';
import { useTranslation } from 'react-i18next';
import styled from 'styled-components';
import { setRangeBold, clearRange } from '../excelFormatters';

const StyledContainer = styled.div`
  text-align: center;
  padding: 15px;
  background-color: white;
  font-family: Arial, sans-serif;
`;

const StyledTitle = styled.h1`
  color: #217346;
  font-size: 24px;
  margin-bottom: 10px;
`;

const StyledText = styled.p`
  font-size: 14px;
  line-height: 1.4;
  margin-bottom: 15px;
`;

const StyledButton = styled.button`
  background-color: #217346;
  color: white;
  border: none;
  padding: 8px 15px;
  font-size: 16px;
  cursor: pointer;
  border-radius: 5px;
  margin-top: 10px;

  &:hover {
    background-color: #1a5c38;
  }
`;

const StyledResetButton = styled(StyledButton)`
  background-color: #f44336; /* Red color for reset */

  &:hover {
    background-color: #d32f2f;
  }
`;

const ButtonContainer = styled.div`
  margin-top: 8px;
  display: flex;
  justify-content: center;
  gap: 10px;
`;

const XlookupHorizontalSearch = ({ resetLesson }) => {
  const { t } = useTranslation();

  const handlePrepareData = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // Clear existing data
        clearRange(context, "A20:D30");

        // Insert headers in row 20
        const headers = [[
          t("month_header"),
          t("month_january"),
          t("month_february"),
          t("month_march")
        ]];
        sheet.getRange("A20:D20").values = headers;
        setRangeBold(context, "A20:D20");

        // Insert sales data in row 21
        const salesData = [[
          t("sales_header"),
          5000,
          7500,
          6200
        ]];
        sheet.getRange("A21:D21").values = salesData;
        setRangeBold(context, "A21");

        // Set up search cell
        sheet.getRange("A23").values = [[t("search_month_label")]];
        setRangeBold(context, "A23");
        sheet.getRange("B23").values = [[t("month_february")]]; // Default search month

        // Set up result cell
        sheet.getRange("A25").values = [[t("sales_header")]];
        setRangeBold(context, "A25");
        const formula = `=XLOOKUP(B23,B20:D20,B21:D21)`;
        sheet.getRange("B25").formulas = [[formula]];

        await context.sync();
      });
    } catch (error) {
      console.error("Error preparing data:", error);
    }
  };

  return (
    <StyledContainer>
      <StyledTitle>{t('horizontal_search_title')}</StyledTitle>
      <StyledText>{t('horizontal_search_text')}</StyledText>
      <ButtonContainer>
        <StyledButton onClick={handlePrepareData}>{t('prepare_data_button')}</StyledButton>
        <StyledResetButton onClick={resetLesson}>{t("reset_lesson_button")}</StyledResetButton>
      </ButtonContainer>
    </StyledContainer>
  );
};

export default XlookupHorizontalSearch;