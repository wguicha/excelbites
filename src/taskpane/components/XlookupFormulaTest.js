/* global Excel */

import React, { useState } from "react";
import { useTranslation } from "react-i18next";
import styled from "styled-components";
import { setRangeBold, setRangeFillColor } from "../excelFormatters";

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
  padding: 20px;
  border: none;
  border-radius: 0;
  background-color: white;
  box-shadow: none;
`;

const StyledLabel = styled.label`
  margin-top: 10px;
  font-weight: bold;
  text-align: left;
  width: 100%;
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

const StyledNavButton = styled(StyledButton)`
  background-color: #a9a9a9;

  &:hover {
    background-color: #808080;
  }
`;

const ButtonContainer = styled.div`
  margin-top: 10px;
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
        const formula = `=XLOOKUP(${lookupValue},${lookupArray},${returnArray})`;
        const targetRange = sheet.getRange("F7");
        targetRange.formulas = [[formula]];

        // Parameter descriptions
        const descriptions = [
          [t("param_lookup_value") + ":", lookupValue],
          [t("param_lookup_array") + ":", lookupArray],
          [t("param_return_array") + ":", returnArray],
          [t("if_not_found_label") + ":", t("not_specified")],
          [t("match_mode_label") + ":", t("exact_match_default")],
          [t("search_mode_label") + ":", t("first_to_last_default")],
        ];

        const descriptionRange = sheet.getRange("E11:F16");
        descriptionRange.values = descriptions;

        // Formatting using formatter function
        setRangeBold(context, "E11:E16");
        setRangeFillColor(context, "F5", "#DAE9F8");
        setRangeFillColor(context, "E11:F11", "#DAE9F8");
        setRangeFillColor(context, "A6:A15", "#FFCDCD");
        setRangeFillColor(context, "E12:F12", "#FFCDCD");
        setRangeFillColor(context, "B6:B15", "#E8D9F3");
        setRangeFillColor(context, "E13:F13", "#E8D9F3");
        await context.sync();
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
      <ButtonContainer>
        <StyledNavButton onClick={() => { console.log("Previous button clicked in XlookupFormulaTest"); goToPreviousStep(); }}>&#9664;</StyledNavButton>
        <StyledNavButton onClick={() => { console.log("Next button clicked in XlookupFormulaTest"); goToNextStep(); }}>&#9654;</StyledNavButton>
      </ButtonContainer>
    </StyledContainer>
  );
};

export default XlookupFormulaTest;