import * as React from "react";
import PropTypes from "prop-types";
import { Spinner, SpinnerSize } from "office-ui-fabric-react";

const Progress = (props) => {
  // render() {
    const { logo, message, title } = props;

    return (
      <section className="ms-welcome__progress ms-u-fadeIn500">
        <img width="90" height="90" src={logo} alt={title} title={title} />
        <h1 className="ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary">{title}</h1>
        <Spinner type={SpinnerSize.large} label={message} />
      </section>
    );
  // }
}

Progress.propTypes = {
  logo: PropTypes.string,
  message: PropTypes.string,
  title: PropTypes.string,
};

export default Progress;