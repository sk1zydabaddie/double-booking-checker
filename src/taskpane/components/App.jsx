import * as React from "react";
import PropTypes from "prop-types";
import { makeStyles } from "@fluentui/react-components";
import BookingChecker from "./BookingChecker";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    backgroundColor: "#f5f5f5",
  },
});

const App = (props) => {
  const { title } = props;
  const styles = useStyles();
  return (
    <div className={styles.root}>
      <BookingChecker />
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
