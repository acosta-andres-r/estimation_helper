import * as React from "react";
import PropTypes from "prop-types";

const HeroList = (props) => {
  // render() {
  const { children, items, message, clickExpand, clickCollapse } = props;

  const listItems = items.map((item, index) => (
    <li className="ms-ListItem" key={index}>
      <i className={`ms-Icon ms-Icon--CollapseMenu`}></i>
      <span className="ms-font-m ms-fontColor-neutralPrimary">{item.name}</span>

      <button
        class="ms-Button ms-Button--hero"
        onClick={() => clickExpand(item.from, item.to)}
      >
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--Add"></i></span>
        {/* <span class="ms-Button-label">Expand</span> */}
      </button>

      <button
        class="ms-Button ms-Button--hero"
        onClick={() => clickCollapse(item.from, item.to)}
      >
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--CollapseContentSingle"></i></span>
        {/* <span class="ms-Button-label">Collapse</span> */}
      </button>
    </li>
  ));
  return (
    <main className="ms-welcome__main">
      <h2 className="ms-font-xl ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">{message}</h2>
      <ul className="ms-List ms-welcome__features ms-u-slideUpIn10">{listItems}</ul>
      {children}
    </main>
  );
  // }
}

HeroList.propTypes = {
  children: PropTypes.node,
  items: PropTypes.array,
  message: PropTypes.string,
};

export default HeroList;