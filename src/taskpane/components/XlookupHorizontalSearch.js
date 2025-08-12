/* global Excel */

import React from 'react';
import { useTranslation } from 'react-i18next';
import { setRangeBold, clearRange } from '../excelFormatters';
import {
  StyledContainer,
  StyledTitle,
  StyledText,
  StyledButton,
  StyledResetButton,
  ButtonContainer,
} from './styles/XlookupHorizontalSearch.styles';

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