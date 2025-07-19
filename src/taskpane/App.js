import React from "react";
import Lesson from "./components/Lesson";
import XlookupIntroduction from "./components/XlookupIntroduction";
import XlookupFunctionUsage from "./components/XlookupFunctionUsage";

const App = () => {
  const xlookupLessonSteps = [
    XlookupIntroduction,
    XlookupFunctionUsage,
    // Add more steps here for the XLOOKUP lesson
  ];

  return (
    <div>
      <Lesson steps={xlookupLessonSteps} />
    </div>
  );
};

export default App;