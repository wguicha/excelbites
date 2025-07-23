/* global Excel */

import React, { useState, useEffect, useRef } from "react";
import { useTranslation } from "react-i18next";
import styled from "styled-components";
import { setRangeBold, setRangeFillColor, clearRangeFill } from "../excelFormatters";

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

const XlookupMultipleSearch = ({ goToNextStep, goToPreviousStep }) => {
  const { t } = useTranslation();
  const [lookupValue, setLookupValue] = useState("F5");
  const [lookupArray, setLookupArray] = useState("A6:A15");
  const [returnArray, setReturnArray] = useState("B6:C15"); // Changed to two columns
  const activeInputRef = useRef(null);
  const eventContextRef = useRef(null);
  const lastHighlightedRangeRef = useRef(null);

  useEffect(() => {
    const registerSelectionHandler = async () => {
      try {
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const handler = sheet.onSelectionChanged.add(handleSelectionChange);
          eventContextRef.current = context;
          await context.sync();
        });
      } catch (error) {
        console.error("Error registering selection handler:", error);
      }
    };

    registerSelectionHandler();

    return () => {
      if (eventContextRef.current) {
        eventContextRef.current.workbook.worksheets.getActiveWorksheet().onSelectionChanged.remove(handleSelectionChange);
      }
    };
  }, []);

  const handleSelectionChange = async (event) => {
    await Excel.run(async (context) => {
      if (activeInputRef.current) {
        const rangeAddress = event.address;
        switch (activeInputRef.current) {
          case "lookupValue":
            setLookupValue(rangeAddress);
            break;
          case "lookupArray":
            setLookupArray(rangeAddress);
            break;
          case "returnArray":
            setReturnArray(rangeAddress);
            break;
          default:
            break;
        }
        activeInputRef.current = null;
      }
    });
  };

  const handleFocus = async (inputName, rangeAddress) => {
    activeInputRef.current = inputName;
    try {
      await Excel.run(async (context) => {
        if (lastHighlightedRangeRef.current) {
          clearRangeFill(context, lastHighlightedRangeRef.current);
        }
        setRangeFillColor(context, rangeAddress, "#FFFF00"); // Yellow highlight
        lastHighlightedRangeRef.current = rangeAddress;
        await context.sync();
      });
    } catch (error) {
      console.error("Error highlighting range:", error);
    }
  };

  const handleInsertFormula = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        // Generate formula with commas as separators
        const formula = `=XLOOKUP(${lookupValue},${lookupArray},${returnArray})`;

        console.log("Attempting to insert formula:", formula);

        sheet.getRange("E9").values = [[t("multiple_formula_label")]];
        setRangeBold(context, "E9");

        const targetRange = sheet.getRange("F9"); // Changed target cell to F9
        targetRange.formulas = [[formula]];

        // Ensure calculation mode is automatic and force recalculation
        context.workbook.application.calculationMode = Excel.CalculationMode.automatic;
        context.workbook.application.calculate(Excel.CalculationType.full);

        await context.sync();

        targetRange.load("formula, values");

        sheet.getRange("F13").values = "B6:C15";
        setRangeFillColor(context, "B6:C15", "#E8D9F3");

        sheet.getRange("F9:G9").select();

        await context.sync();
        console.log("Formula read from F9 after sync:", targetRange.formula);
        console.log("Value read from F9 after sync:", targetRange.values[0][0]);

        console.log("Formula inserted successfully!");
      });
    } catch (error) {
      console.error("Error inserting formula:", error);
    }
  };

  return (
    <StyledContainer>
      <StyledTitle>{t("multiple_search_title")}</StyledTitle>
      <StyledText>{t("multiple_search_text")}</StyledText>
      <StyledForm>
        <StyledLabel>{t("lookup_value_label")}</StyledLabel>
        <StyledInput type="text" value={lookupValue} onFocus={() => handleFocus("lookupValue", lookupValue)} onChange={(e) => setLookupValue(e.target.value)} />

        <StyledLabel>{t("lookup_array_label")}</StyledLabel>
        <StyledInput type="text" value={lookupArray} onFocus={() => handleFocus("lookupArray", lookupArray)} onChange={(e) => setLookupArray(e.target.value)} />

        <StyledLabel>{t("return_array_label")}</StyledLabel>
        <StyledInput type="text" value={returnArray} onFocus={() => handleFocus("returnArray", returnArray)} onChange={(e) => setReturnArray(e.target.value)} />
      </StyledForm>
      <StyledButton onClick={handleInsertFormula}>{t("insert_formula_button")}</StyledButton>
      <ButtonContainer>
        <StyledNavButton onClick={goToPreviousStep}>&#9664;</StyledNavButton>
        <StyledNavButton onClick={goToNextStep}>&#9654;</StyledNavButton>
      </ButtonContainer>
    </StyledContainer>
  );
};

export default XlookupMultipleSearch;
