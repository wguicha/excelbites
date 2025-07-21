import * as React from "react";
import PropTypes from "prop-types";
import Lesson from "./Lesson";
import XlookupIntroduction from "./XlookupIntroduction";
import XlookupFormulaTest from "./XlookupFormulaTest";
import XlookupMultipleSearch from "./XlookupMultipleSearch";

const App = (props) => {
  const { title } = props;



  const lessonSteps = [
    XlookupIntroduction,
    XlookupFormulaTest,
    XlookupMultipleSearch,
  ];

  console.log("App.jsx: lessonSteps.length =", lessonSteps.length);

  return (
    <Lesson steps={lessonSteps} />
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
