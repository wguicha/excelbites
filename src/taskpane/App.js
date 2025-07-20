import React from "react";
import Lesson from "./components/Lesson";
import XlookupIntroduction from "./components/XlookupIntroduction";
import XlookupFormulaTest from "./components/XlookupFormulaTest";

const App = () => {
  const xlookupLessonSteps = [
    XlookupIntroduction,
    XlookupFormulaTest,
    // Add more steps here for the XLOOKUP lesson
  ];

  return (
    <div>
      <Lesson steps={xlookupLessonSteps} />
    </div>
  );
};

export default App;