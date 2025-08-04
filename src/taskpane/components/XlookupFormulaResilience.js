/* global Excel */

import React, { useState, useEffect, useRef } from "react";
import { useTranslation } from "react-i18next";
import styled from "styled-components";
import { setRangeBold, setRangeFillColor, clearRangeFill } from "../excelFormatters";

const StyledContainer = styled.div`
  text-align: center;
  padding: 15px; /* Reduced padding */
  background-color: white;
  font-family: Arial, sans-serif;
`;

const StyledTitle = styled.h1`
  color: #217346;
  font-size: 24px; /* Slightly smaller font size */
  margin-bottom: 10px; /* Reduced margin */
`;

const StyledText = styled.p`
  font-size: 14px; /* Slightly smaller font size */
  line-height: 1.4;
  margin-bottom: 15px; /* Reduced margin */
`;

const StyledForm = styled.div`
  display: flex;
  flex-direction: column;
  align-items: flex-start;
  margin: 0 auto;
  max-width: 280px; /* Slightly reduced max-width */
  padding: 15px; /* Reduced padding */
  border: none;
  border-radius: 0;
  background-color: white;
  box-shadow: none;
`;

const StyledLabel = styled.label`
  margin-top: 8px; /* Reduced margin */
  font-weight: bold;
  text-align: left;
  width: 100%;
  font-size: 14px; /* Slightly smaller font size */
`;

const StyledInput = styled.input`
  width: 100%;
  padding: 6px; /* Reduced padding */
  margin-top: 3px; /* Reduced margin */
  border: 1px solid #ddd;
  border-radius: 4px;
  font-size: 14px; /* Slightly smaller font size */
`;

const StyledButton = styled.button`
  background-color: #217346;
  color: white;
  border: none;
  padding: 6px 12px; /* Further reduced padding */
  font-size: 14px; /* Further smaller font size */
  cursor: pointer;
  border-radius: 5px;
  margin-top: 10px; /* Reduced margin */
  min-width: 150px; /* Added min-width for consistent sizing */

  &:hover {
    background-color: #1a5c38;
  }

  &:disabled {
    background-color: #cccccc;
    cursor: not-allowed;
  }
`;

const StyledNavButton = styled(StyledButton)`
  background-color: #a9a9a9;
  margin: 3px; /* Reduced margin */

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
  margin-top: 8px; /* Reduced margin */
  display: flex;
  justify-content: center;
  gap: 10px; /* Space between buttons */
`;

const StyledMessage = styled.p`
  color: #217346;
  font-weight: bold;
  margin-top: 8px; /* Reduced margin */
  background-color: #e6ffe6;
  border: 1px solid #217346;
  padding: 4px; /* Reduced padding */
  border-radius: 4px;
  font-size: 14px; /* Slightly smaller font size */
`;

const XlookupFormulaResilience = ({ goToNextStep, goToPreviousStep, resetLesson }) => {
  const { t } = useTranslation();
  const [message, setMessage] = useState(null);
  const [isMoveButtonDisabled, setIsMoveButtonDisabled] = useState(false);

  const handleMoveColumns = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        
        // Move columns B and C to column I
        const rangeToMove = sheet.getRange("B:C");
        const targetRange = sheet.getRange("I:I");
        rangeToMove.moveTo(targetRange);

        // Delete columns B and C
        sheet.getRange("B:C").delete(Excel.DeleteShiftDirection.left);

        // Apply colors to relevant ranges AFTER columns are moved and deleted
        setRangeFillColor(context, "F5", "#DAE9F8");
        setRangeFillColor(context, "A6:A15", "#FFCDCD");
        setRangeFillColor(context, "C6:C15", "#E8D9F3");

        await context.sync();

        setMessage(t("columns_moved_success"));
        setIsMoveButtonDisabled(true); // Disable the button after successful operation
        setTimeout(() => setMessage(null), 5000); // Clear message after 5 seconds
      });
    } catch (error) {
      console.error("Error moving columns:", error);
      setMessage(t("columns_moved_error"));
      setTimeout(() => setMessage(null), 5000); // Clear message after 5 seconds
    }
  };

  return (
    <StyledContainer>
      <StyledTitle>{t("formula_resilience_title")}</StyledTitle>
      <StyledText>{t("formula_resilience_text")}</StyledText>
      <ButtonContainer>
        <StyledButton onClick={handleMoveColumns} disabled={isMoveButtonDisabled}>{t("move_columns_button")}</StyledButton>
        <StyledResetButton onClick={resetLesson}>{t("reset_lesson_button")}</StyledResetButton>
      </ButtonContainer>
      {message && <StyledMessage>{message}</StyledMessage>}
    </StyledContainer>
  );
};

export default XlookupFormulaResilience;