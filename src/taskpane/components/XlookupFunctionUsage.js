/* global Excel */

import React from "react";
import { useTranslation } from "react-i18next";
import styled from "styled-components";
import { setColumnWidth } from "../excelFormatters";

const StyledContainer = styled.div`
  text-align: center;
  padding: 20px;
  background-color: white;
  font-family: Arial, sans-serif;
`;

const StyledTitle = styled.h1`
  color: #217346;
  font-size: 28px;
  margin-bottom: 15px;
`;

const StyledParagraph = styled.p`
  font-size: 16px;
  line-height: 1.5;
  margin-bottom: 20px;
`;

const StyledButton = styled.button`
  background-color: #217346;
  color: white;
  border: none;
  padding: 10px 20px;
  font-size: 18px;
  cursor: pointer;
  border-radius: 5px;
  margin: 5px;

  &:hover {
    background-color: #1a5c38;
  }
`;

const XlookupFunctionUsage = ({ goToNextStep, goToPreviousStep }) => {
  const { t } = useTranslation();

  const handleVerifyFormula = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const dataLength = 10; // Number of data rows from XlookupIntroduction
        const searchIdRow = dataLength + 7; // F17
        const resultRow = dataLength + 8; // F18

        const resultCell = sheet.getRange("F" + resultRow);
        resultCell.load("formula");
        await context.sync();

        const expectedFormula = `=XLOOKUP(F${searchIdRow},A6:A${dataLength + 5},C6:C${dataLength + 5})`;

        if (resultCell.formula.toUpperCase() === expectedFormula.toUpperCase()) {
          alert("¡Fórmula correcta! ¡Bien hecho!");
        } else {
          alert("Fórmula incorrecta. Inténtalo de nuevo.\nEsperado: " + expectedFormula + "\nObtenido: " + resultCell.formula);
        }
      });
    }
    catch (error) {
      console.error("Error verifying formula:", error);
    }
  };

  const handleSetColumnWidth = async () => {
    try {
      await Excel.run(async (context) => {
        setColumnWidth(context, ["A", "B", "C", "D", "E", "F"], 14);
        await context.sync();
      });
    } catch (error) {
      console.error("Error setting column width:", error);
    }
  };

  return (
    <StyledContainer>
      <StyledTitle>{t("xlookup_usage_title")}</StyledTitle>
      <StyledParagraph dangerouslySetInnerHTML={{ __html: t("xlookup_usage_description") }} />
      <StyledParagraph>
        <strong>{t("xlookup_formula_example_title")}</strong>
        <br />
        <code>=BUSCARX(F17,A6:A15,C6:C15)</code>
      </StyledParagraph>
      <StyledButton onClick={handleVerifyFormula}>{t("verify_formula_button")}</StyledButton>
      <StyledParagraph>{t("set_column_width_instruction")}</StyledParagraph>
      <StyledButton onClick={handleSetColumnWidth}>Set Column Width</StyledButton>
      <StyledButton onClick={goToPreviousStep}>Previous</StyledButton>
      <StyledButton onClick={goToNextStep}>Next</StyledButton>
    </StyledContainer>
  );
};

export default XlookupFunctionUsage;
