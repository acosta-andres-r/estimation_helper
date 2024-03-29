import * as React from "react";
import PropTypes from "prop-types";

const Header = (props) => {
  // render() {
    const { title, logo, message } = props;

    return (
      <section className=".ms-welcome__header ms-u-fadeIn500">
        {/* <img width="90" height="90" src={logo} alt={title} title={title} /> */}
        <h1 className="ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary" align="center">{message}</h1>
      </section>
    );
  // }
}

Header.propTypes = {
  title: PropTypes.string,
  logo: PropTypes.string,
  message: PropTypes.string,
};

export default Header;