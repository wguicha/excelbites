import React, { useState } from "react";

const Lesson = ({ steps }) => {
  const [currentStepIndex, setCurrentStepIndex] = useState(0);
  const CurrentStepComponent = steps[currentStepIndex];

  const goToNextStep = () => {
    if (currentStepIndex < steps.length - 1) {
      setCurrentStepIndex(currentStepIndex + 1);
    }
  };

  const goToPreviousStep = () => {
    if (currentStepIndex > 0) {
      setCurrentStepIndex(currentStepIndex - 1);
    }
  };

  return (
    <div>
      <CurrentStepComponent goToNextStep={goToNextStep} goToPreviousStep={goToPreviousStep} />
      {/* Navigation buttons can be added here or within each step component */}
    </div>
  );
};

export default Lesson;
