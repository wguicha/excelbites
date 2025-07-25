import React, { useState, useEffect } from "react";
import { clearAllRangeFills } from "../excelFormatters";

const Lesson = ({ steps }) => {
  const [currentStepIndex, setCurrentStepIndex] = useState(0); // Initialize to 0

  // Effect to load currentStepIndex from settings when component mounts and Office is ready
  useEffect(() => {
    Office.onReady((info) => {
      if (info.host === Office.HostType.Excel) {
        if (Office.context && Office.context.document && Office.context.document.settings) {
          const savedIndex = Office.context.document.settings.get("lessonStepIndex");
          if (savedIndex !== null && savedIndex !== undefined) {
            setCurrentStepIndex(parseInt(savedIndex, 10));
            console.log("Loaded lesson step:", parseInt(savedIndex, 10));
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
    if (currentStepIndex < steps.length - 1) {
      await Excel.run(async (context) => {
        clearAllRangeFills(context);
        await context.sync();
      });
      setCurrentStepIndex(prevIndex => prevIndex + 1);
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

  const CurrentStepComponent = steps[currentStepIndex];

  return (
    <div>
      {React.createElement(steps[currentStepIndex], {
        goToNextStep: goToNextStep,
        goToPreviousStep: goToPreviousStep,
      })}
      {/* Navigation buttons can be added here or within each step component */}
    </div>
  );
};

export default Lesson;
