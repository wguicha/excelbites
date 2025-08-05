import React, { useState, useEffect } from "react";
import PropTypes from "prop-types";
import Lesson from "./Lesson";
import XlookupIntroduction from "./XlookupIntroduction";
import XlookupFormulaTest from "./XlookupFormulaTest";
import XlookupMultipleSearch from "./XlookupMultipleSearch";
import XlookupErrorHandling from "./XlookupErrorHandling";
import XlookupFormulaResilience from "./XlookupFormulaResilience";
import { useTranslation } from "react-i18next";
import styled from "styled-components";
import { clearAllContentAndFormats, clearAllRangeFills } from "../excelFormatters";

const StyledButton = styled.button`
  background-color: #f44336; /* Red color for reset */
  color: white;
  border: none;
  padding: 8px 15px; /* Reduced padding */
  font-size: 16px; /* Slightly smaller font size */
  cursor: pointer;
  border-radius: 5px;
  margin-top: 15px; /* Reduced margin */

  &:hover {
    background-color: #d32f2f;
  }
`;

const StyledNavButton = styled(StyledButton)`
  background-color: #217346;
  margin: 3px; /* Reduced margin */

  &:hover {
    background-color: #1a5c38;
  }

  &:disabled {
    background-color: #cccccc;
    cursor: not-allowed;
  }
`;

const StyledFooter = styled.div`
  position: fixed;
  bottom: 0;
  left: 0;
  width: 100%; /* Set width to 100% */
  box-sizing: border-box; /* Include padding in width calculation */
  background-color: #f0f0f0; /* Light gray background for footer */
  padding: 10px 15px;
  display: flex;
  justify-content: space-between; /* Distribute items with space between them */
  align-items: center;
  box-shadow: 0 -2px 5px rgba(0, 0, 0, 0.1);
  z-index: 1000;
`;

const StyledNavButtonsContainer = styled.div`
  display: flex;
  gap: 10px; /* Space between nav buttons */
  align-items: center;
`;

const App = (props) => {
  const { t } = useTranslation();
  const { title } = props;
  const [currentStepIndex, setCurrentStepIndex] = useState(0); // Initialize to 0

  const lessonSteps = [
    XlookupIntroduction,
    XlookupFormulaTest,
    XlookupMultipleSearch,
    XlookupErrorHandling,
    XlookupFormulaResilience,
  ];

  // Effect to load currentStepIndex from settings when component mounts and Office is ready
  useEffect(() => {
    Office.onReady((info) => {
      if (info.host === Office.HostType.Excel) {
        if (Office.context && Office.context.document && Office.context.document.settings) {
          const savedIndex = Office.context.document.settings.get("lessonStepIndex");
          if (savedIndex !== null && savedIndex !== undefined) {
            const parsedIndex = parseInt(savedIndex, 10);
            setCurrentStepIndex(parsedIndex);
            console.log("Loaded lesson step:", parsedIndex);
          }
        }
      }
    });
  }, []); // Empty dependency array ensures this runs only once on mount

  // Effect to save currentStepIndex to settings whenever it changes
  useEffect(() => {
    console.log("currentStepIndex changed to:", currentStepIndex);
    if (Office.context && Office.context.document && Office.context.document.settings) {
      Office.context.document.settings.set("lessonStepIndex", currentStepIndex);
      Office.context.document.settings.saveAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Error saving settings:", asyncResult.error.message);
        } else {
          console.log("Lesson step saved:", currentStepIndex);
        }
      });
    }
  }, [currentStepIndex]); // This effect runs whenever currentStepIndex changes

  const goToNextStep = async () => {
    console.log("goToNextStep called. Current index:", currentStepIndex);
    if (currentStepIndex < lessonSteps.length - 1) {
      const nextStepIndex = currentStepIndex + 1;
      const nextStepComponent = lessonSteps[nextStepIndex];

      await Excel.run(async (context) => {
        clearAllRangeFills(context);

        // Special handling for XlookupFormulaResilience step
        if (nextStepComponent.name === "XlookupFormulaResilience") {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          sheet.getRange("F5").values = [[104]]; // Set a valid search ID
        }

        await context.sync();
      });

      setCurrentStepIndex(nextStepIndex);
    }
  };

  const goToPreviousStep = async () => {
    console.log("goToPreviousStep called. Current index:", currentStepIndex);
    if (currentStepIndex > 0) {
      await Excel.run(async (context) => {
        clearAllRangeFills(context);
        await context.sync();
      });
      setCurrentStepIndex(prevIndex => prevIndex - 1);
    }
  };

  const resetLesson = async () => {
    console.log("Resetting lesson.");
    if (Office.context && Office.context.document && Office.context.document.settings) {
      Office.context.document.settings.remove("lessonStepIndex");
      Office.context.document.settings.saveAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Error removing settings:", asyncResult.error.message);
        } else {
          console.log("Lesson step setting removed.");
        }
      });
    }
    await Excel.run(async (context) => {
      clearAllContentAndFormats(context);
      await context.sync();
    });
    setCurrentStepIndex(0); // Start at step 0 (introduction)
  };

  console.log("App.jsx: lessonSteps.length =", lessonSteps.length);

  return (
    <div style={{ paddingBottom: '60px' }}> {/* Add padding to prevent content from being hidden by fixed footer */}
      <Lesson
        steps={lessonSteps}
        currentStepIndex={currentStepIndex}
        setCurrentStepIndex={setCurrentStepIndex}
        goToNextStep={goToNextStep}
        goToPreviousStep={goToPreviousStep}
        resetLesson={resetLesson}
      />
      <StyledFooter>
        <div style={{ flex: 1 }}></div> {/* Left spacer */}
        <StyledNavButtonsContainer>
          <StyledNavButton onClick={goToPreviousStep} disabled={currentStepIndex === 0}>&#9664;</StyledNavButton>
          <StyledNavButton onClick={goToNextStep} disabled={currentStepIndex === lessonSteps.length - 1}>&#9654;</StyledNavButton>
        </StyledNavButtonsContainer>
        <div style={{ flex: 1 }}></div> {/* Right spacer */}
      </StyledFooter>
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;