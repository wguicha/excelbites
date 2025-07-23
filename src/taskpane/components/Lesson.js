import React, { useState, useEffect } from "react";
import { clearAllRangeFills } from "../excelFormatters";

const Lesson = ({ steps }) => {
  console.log("Lesson component rendered. steps.length:", steps);
  const [currentStepIndex, setCurrentStepIndex] = useState(0);

  useEffect(() => {
    console.log("currentStepIndex changed to:", currentStepIndex);
  }, [currentStepIndex]);

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
