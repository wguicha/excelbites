import React, { useEffect } from "react";
import { clearAllRangeFills } from "../excelFormatters";

const Lesson = ({ steps, currentStepIndex, setCurrentStepIndex, goToNextStep, goToPreviousStep }) => {
  console.log("Lesson component rendered. currentStepIndex:", currentStepIndex);

  const CurrentStepComponent = steps[currentStepIndex];

  return (
    <div>
      {React.createElement(steps[currentStepIndex], {
        goToNextStep: goToNextStep,
        goToPreviousStep: goToPreviousStep,
        resetLesson: resetLesson,
      })}
      {/* Navigation buttons can be added here or within each step component */}
    </div>
  );
};

export default Lesson;
