/* global Excel */

import React, { useState } from "react";
import { useTranslation } from "react-i18next";
import styled from "styled-components";

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

const StyledText = styled.p`
  font-size: 16px;
  line-height: 1.5;
  margin-bottom: 20px;
`;

const StyledForm = styled.div`
  display: flex;
  flex-direction: column;
  align-items: flex-start;
  margin: 0 auto;
  max-width: 300px;
  padding: 10px;
  border: 1px solid #ccc;
  border-radius: 5px;
`;

const StyledLabel = styled.label`
  margin-top: 10px;
  font-weight: bold;
`;

const StyledInput = styled.input`
  width: 100%;
  padding: 8px;
  margin-top: 5px;
  border: 1px solid #ddd;
  border-radius: 4px;
`;

const StyledButton = styled.button`
  background-color: #217346;
  color: white;
  border: none;
  padding: 10px 20px;
  font-size: 18px;
  cursor: pointer;
  border-radius: 5px;
  margin-top: 20px;

  &:hover {
    background-color: #1a5c38;
  }
`;

const XlookupFormulaTest = ({ goToNextStep, goToPreviousStep }) => {
  const { t } = useTranslation();
  const [lookupValue, setLookupValue] = useState("F5");
  const [lookupArray, setLookupArray] = useState("A6:A15");
  const [returnArray, setReturnArray] = useState("C6:C15");

  const handleInsertFormula = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        // Generate formula with commas as separators
        const formula = `=XLOOKUP(${lookupValue},${lookupArray},${returnArray})`;
        
        console.log("Attempting to insert formula:", formula);
        
        const targetRange = sheet.getRange("F7");
        targetRange.formulas = [[formula]];
        
        // Ensure calculation mode is automatic and force recalculation
        context.workbook.application.calculationMode = Excel.CalculationMode.automatic;
        context.workbook.application.calculate(Excel.CalculationType.full);

        await context.sync();
        
        targetRange.load("formula, values");
        await context.sync();
        console.log("Formula read from F7 after sync:", targetRange.formula);
        console.log("Value read from F7 after sync:", targetRange.values[0][0]);

        console.log("Formula inserted successfully!");
      });
    } catch (error) {
      console.error("Error inserting formula:", error);
    }
  };

  return (
    <StyledContainer>
      <StyledTitle>{t("formula_test_title")}</StyledTitle>
      <StyledText>{t("formula_test_text")}</StyledText>
      <StyledForm>
        <StyledLabel>{t("lookup_value_label")}</StyledLabel>
        <StyledInput type="text" value={lookupValue} onChange={(e) => setLookupValue(e.target.value)} />

        <StyledLabel>{t("lookup_array_label")}</StyledLabel>
        <StyledInput type="text" value={lookupArray} onChange={(e) => setLookupArray(e.target.value)} />

        <StyledLabel>{t("return_array_label")}</StyledLabel>
        <StyledInput type="text" value={returnArray} onChange={(e) => setReturnArray(e.target.value)} />
      </StyledForm>
      <StyledButton onClick={handleInsertFormula}>{t("insert_formula_button")}</StyledButton>
      <StyledButton onClick={goToPreviousStep}>Previous</StyledButton>
      <StyledButton onClick={goToNextStep}>Next</StyledButton>
    </StyledContainer>
  );
};

export default XlookupFormulaTest;