import * as React from "react";
import { useState, useEffect } from "react";
import PropTypes from "prop-types";
import { DefaultButton, ButtonType } from "office-ui-fabric-react";
import Header from "./Header";
import HeroList from "./HeroList";
import Progress from "./Progress";
/* global console, Excel */

const ELEMENT_GROUPS_ROWS = [
  // Third Party Vendor
  { row: "14" },
  // Specialty Material
  { row: "22" },
  // Aluminum - Sheet
  { row: "30" },
  // Aluminum Tube/Angle/Rod
  { row: "39" },
  // Grimco
  { row: "48" },
  // Acrylic
  { row: "56" },
  // Sintra
  { row: "80" },
  // MDF/MDO/Lumber
  { row: "82" },
  // Hardware
  { row: "96" },
  // Paint
  { row: "107" },
  // Vinyl
  { row: "114" },
  // Freight
  { row: "123" },
  // Drop
  { row: "128" },
  // Fabrication Services
  { row: "130" },
  // Design Services
  { row: "145" }
]

const App = (props) => {
  const [listElementGroup, setListElementGroup] = useState([]);

  useEffect(() => {

    const getElementGroupText = async () => {
      try {
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();

          let displayRange = sheet.getRange(`B14:B147`).load("values");
          let displayRange2 = sheet.getRange("K10");
          await context.sync(displayRange); // Only when using load or update data from sheet context.sync(sheet)

          // Get Element Groups text
          const elementGroup = ELEMENT_GROUPS_ROWS.map((item, index) => {
            return {
              name: displayRange.values[parseInt(item.row) - 14][0],
              from: parseInt(item.row) + 1,
              to: item.row == "145" ? 147 : parseInt(ELEMENT_GROUPS_ROWS[index + 1].row) - 1,
            };
          });

          await context.sync(displayRange2); // Only when using load or update data from sheet context.sync(sheet)
          // displayRange2.values = [[JSON.stringify(elementGroup)]];

          setListElementGroup(elementGroup);

        });
      } catch (error) {
        console.error(error);
        setListElementGroup([{ name: error }]);
      }
    };

    getElementGroupText();

  }, [])



  const click = async () => {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load("address");

        // Update the fill color
        range.format.fill.color = "yellow";

        // Values
        range.values = "9";

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };

  const hideEmptyRows = async (rowStart = 15, rowEnd = 21, letterColumn = "A") => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // Get quantity in Third Party Vendors
        let range = sheet.getRange(`${letterColumn}${rowStart}:${letterColumn}${rowEnd}`).load("values");
        let displayRange = sheet.getRange("K8");
        let displayRange2 = sheet.getRange("K10");
        await context.sync(range); // Only when using load or update data from sheet context.sync(sheet)

        // Find empty cells and hide the corresponding rows
        const hideRows = range.values.map((element, index) => {
          if (element[0] == "") {
            sheet.getRange(`${letterColumn}${rowStart + index}`).rowHidden = true
            return `Hide row ${rowStart + index}`;
          };
          return "Unhide";
        });


        // Display results in cell K8
        const str = JSON.stringify(hideRows);
        // displayRange.values = [[str]];
        // displayRange2.values = [["test"]]

      });
    } catch (error) {
      console.error(error);
    }
  };

  const unhideEmptyRows = async (rowStart = 15, rowEnd = 21, letterColumn = "A") => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // Unhide all rows in range
        sheet.getRange(`${letterColumn}${rowStart}:${letterColumn}${rowEnd}`).rowHidden = false;

        // sheet.getRange("K10").values = [["unhide"]]
        console.log("test")

      });
    } catch (error) {
      console.error(error);
    }
  };

  const hideActiveRows = async () => {
    try {
      await Excel.run(async (context) => {

        // const range = context.workbook.getSelectedRange();
        const sheet = context.workbook.worksheets.getActiveWorksheet()

        let range = sheet.getRange("A15:A21");
        range.load("values");

        let displayRange = sheet.getRange("K8");

        await context.sync(); // Important sync after get info

        // const rangeFill = range.format.fill.color;
        const rangeFill = range;

        const str = JSON.stringify(rangeFill);

        displayRange.values = [[str]];

        await context.sync();

        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };

  const { title, isOfficeInitialized } = props;

  if (!isOfficeInitialized) {
    return (
      <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
    );
  }

  return (
    <div className="ms-welcome">
      <Header logo="assets/logo-filled.png" title={title} message="Welcome" />
      <HeroList
        message="Hide and Unhide Element Groups."
        items={listElementGroup}
        clickExpand={(start, end) => { unhideEmptyRows(start, end) }}
        clickCollapse={(start, end) => { hideEmptyRows(start, end) }}
      >
        <p className="ms-font-l">
          Modify the source files, then click <b>Run</b>.
        </p>
        <DefaultButton
          className="ms-welcome__action"
          buttonType={ButtonType.hero}
          iconProps={{ iconName: "ChevronRight" }}
          // onClick={() => getElementGroupText()}
          onClick={() => hideEmptyRows()}
        >
          Hide
        </DefaultButton>
        <DefaultButton
          className="ms-welcome__action"
          buttonType={ButtonType.hero}
          iconProps={{ iconName: "ChevronLeft" }}
          // onClick={() => click()}
          onClick={() => unhideEmptyRows()}
        >
          Unhide
        </DefaultButton>
      </HeroList>
    </div >
  );
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};

export default App;