function isValidGroupSelected(selectedGroup, DOB) {
  const dob = new Date(DOB).getTime();
  let res = false;
  switch (selectedGroup) {
    case "Group 1":
      // 25-12-2015 to 24-12-2019
      res = dob >= new Date("2015/12/25").getTime() && dob <= new Date("2019/12/24").getTime();
      return [res, res ? "" : "DOB and selected group mismatch."];
    case "Group 2":
      // 25-12-2012 to 24-12-2015
      res = dob >= new Date("2012/12/25").getTime() && dob <= new Date("2015/12/24").getTime();
      return [res, res ? "" : "DOB and selected group mismatch."];
    case "Group 3":
      // 25-12-2008 to 24-12-2012
      res = dob >= new Date("2008/12/25").getTime() && dob <= new Date("2012/12/24").getTime();
      return [res, res ? "" : "DOB and selected group mismatch."];
    case "General Category":
      // 25-12-2006 to 24-12-2010
      res = dob >= new Date("2006/12/25").getTime() && dob <= new Date("2010/12/24").getTime();
      return [res, res ? "" : "DOB and selected group mismatch."];
    default:
      return [false, "Invalid Group."];
  }
}

function isValidEventRegistration(selectedGroup, passedG2E, g1e1, g1e2, g2e1, g2e2, q) {
  switch (selectedGroup) {
    case "Group 1":
      if (g1e1 == g1e2) {
        return [false, "Same events selected."];
      }

      if ((g1e1 == "Bhajans" && g1e2 == "Tamizh Chants") || (g1e2 == "Bhajans" && g1e1 == "Tamizh Chants")) {
        return [false, "Cannot participate in Bhajans and Tamizh Chants at the same time."];
      }

      return [true, ""];
    case "Group 2":
      if (g2e1 == g2e2) {
        return [false, "Same events selected."];
      }

      if ((g2e1.includes("GROUP") && g2e2.includes("GROUP"))) {
        return [false, "Cannot participate in two group events at the same time."];
      }

      if ((g2e1.includes("Bhajans") && g2e2.includes("Tamizh Chants")) || (g2e2.includes("Bhajans") && g2e1.includes("Tamizh Chants"))) {
        return [false, "Cannot participate in Bhajans and Tamizh Chants at the same time."];
      }

      return [true, ""];
    case "Group 3":
      if (g2e1 == g2e2) {
        return [false, "Same events selected."];
      }

      if ((!g2e1.includes("GROUP") || (g2e2.length > 0)) && passedG2E == "No") {
        return [false, "Cannot participate in Individual event. Did not pass Group 3 exam."];
      }


      if ((g2e1.includes("GROUP") && g2e2.includes("GROUP"))) {
        return [false, "Cannot participate in two group events at the same time."];
      }

      if ((g2e1.includes("Bhajans") && g2e2.includes("Tamizh Chants")) || (g2e2.includes("Bhajans") && g2e1.includes("Tamizh Chants"))) {
        return [false, "Cannot participate in Bhajans and Tamizh Chants at the same time."];
      }

      if ((q == "Yes" && g2e1 == "Drawing") || (q == "Yes" && g2e2 == "Drawing")) {
        return [false, "Cannot participate in Quiz and Drawing at the same time."];
      }

      return [true, ""];
    case "General Category":
      if (g1e1 == "" && g1e2 == "" && g2e1 == "" && g2e2 == "" && (q == "Yes" || q == "No")) {
        return [true, ""];
      }

      return [false, "Invalid Registration"];
    default:
      return [false, "Invalid Group"];
  }
}

function validateSheet() {
  const sheet = SpreadsheetApp.openById('<YOUR_GOOGLE_SHEET_ID>').getSheets()[0];

  let targetRange = sheet.getDataRange();
  let rows = targetRange.getValues();
  
  for(let i = 1; i < rows.length; i++) {
    let lastIndex = rows[i].length - 1;
    rows[i][lastIndex] = "";

    const isValidGroup = isValidGroupSelected(rows[i][1], rows[i][3]);
    rows[i][lastIndex - 3] = isValidGroup[0] ? "Accepted" : "Not Allowed";
    rows[i][lastIndex] += isValidGroup[1];

    const isValidEventReg = isValidEventRegistration(rows[i][1], rows[i][9], rows[i][10], rows[i][11], rows[i][12], rows[i][13], rows[i][14]);
    rows[i][lastIndex - 2] = isValidEventReg[0] ? "Accepted" : "Not Allowed";
    rows[i][lastIndex] += ` ${isValidEventReg[1]}`;

    if (isValidGroup[0] && isValidEventReg[0]) {
      rows[i][lastIndex - 1] = "Accepted";
    } else {
      rows[i][lastIndex - 1] = "Not Allowed";
    }
  }

  targetRange.setValues(rows);
}
