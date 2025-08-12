/* global Excel */

import React, { useState, useEffect, useRef } from "react";
import { useTranslation } from "react-i18next";
import { setRangeBold, setRangeFillColor, clearRangeFill } from "../excelFormatters";
import {
  StyledContainer,
  StyledTitle,
  StyledText,
  StyledButton,
  StyledResetButton,
  ButtonContainer,
  StyledMessage,
} from "./styles/XlookupFormulaResilience.styles";

const XlookupFormulaResilience = ({ goToNextStep, goToPreviousStep, resetLesson }) => {
  const { t } = useTranslation();
  const [message, setMessage] = useState(null);
  const [areColumnsMoved, setAreColumnsMoved] = useState(false);

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
        setRangeFillColor(context, "D5", "#DAE9F8");
        setRangeFillColor(context, "A6:A15", "#FFCDCD");
        setRangeFillColor(context, "H6:H15", "#E8D9F3");

        sheet.getRange("D7").select();

        await context.sync();

        setMessage(t("columns_moved_success"));
        setAreColumnsMoved(true);
        setTimeout(() => setMessage(null), 5000); // Clear message after 5 seconds
      });
    } catch (error) {
      console.error("Error moving columns:", error);
      setMessage(t("columns_moved_error"));
      setTimeout(() => setMessage(null), 5000); // Clear message after 5 seconds
    }
  };

  const handleRestoreColumns = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // Insert new columns at B:C
        sheet.getRange("B:C").insert(Excel.InsertShiftDirection.right);

        // Move columns from I:J back to B:C
        const rangeToMove = sheet.getRange("I:J");
        const targetRange = sheet.getRange("B:B");
        rangeToMove.moveTo(targetRange);

        // Clear the now-empty columns I:J
        sheet.getRange("I:J").clear(Excel.ClearApplyTo.all);

        // Re-apply colors to original positions
        setRangeFillColor(context, "F5", "#DAE9F8");
        setRangeFillColor(context, "A6:A15", "#FFCDCD");
        setRangeFillColor(context, "C6:C15", "#E8D9F3");

        sheet.getRange("F7").select();

        await context.sync();

        setMessage(t("columns_restored_success"));
        setAreColumnsMoved(false);
        setTimeout(() => setMessage(null), 5000);
      });
    } catch (error) {
      console.error("Error restoring columns:", error);
      setMessage(t("columns_restored_error"));
      setTimeout(() => setMessage(null), 5000);
    }
  };

  return (
    <StyledContainer>
      <StyledTitle>{t("formula_resilience_title")}</StyledTitle>
      <StyledText>{t("formula_resilience_text")}</StyledText>
      <ButtonContainer>
        {!areColumnsMoved ? (
          <StyledButton onClick={handleMoveColumns}>{t("move_columns_button")}</StyledButton>
        ) : (
          <StyledButton onClick={handleRestoreColumns}>{t("restore_columns_button")}</StyledButton>
        )}
        <StyledResetButton onClick={resetLesson}>{t("reset_lesson_button")}</StyledResetButton>
      </ButtonContainer>
      {message && <StyledMessage>{message}</StyledMessage>}
    </StyledContainer>
  );
};

export default XlookupFormulaResilience;