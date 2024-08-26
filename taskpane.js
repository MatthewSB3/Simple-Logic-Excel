/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("Office is ready");
    document.getElementById("insertAndLogic").onclick = insertAndLogic;
    document.getElementById("insertIfLogic").onclick = insertIfLogic;
    document.getElementById("insertOrLogic").onclick = insertOrLogic;
    document.getElementById("insertIfsLogic").onclick = insertIfsLogic;
    document.getElementById("addIfsCondition").onclick = addIfsCondition;
  }
});

function showFunction(functionName) {
  console.log(`Showing function: ${functionName}`);
  document.getElementById("homeScreen").style.display = "none";
  document.getElementById(functionName + "Page").style.display = "block";
}

function showHomeScreen() {
  console.log("Showing home screen");
  document.getElementById("homeScreen").style.display = "block";
  document.getElementById("andLogicPage").style.display = "none";
  document.getElementById("ifLogicPage").style.display = "none";
  document.getElementById("orLogicPage").style.display = "none";
  document.getElementById("ifsLogicPage").style.display = "none";
}

async function insertAndLogic() {
  try {
    console.log("Inserting AND logic");
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const cell = document.getElementById("andCell").value;
      const trueValue = document.getElementById("andTrueValue").value;
      const falseValue = document.getElementById("andFalseValue").value;
      const condition1 = document.getElementById("andCondition1").value;
      const operator1 = document.getElementById("andOperator1").value;
      const value1 = document.getElementById("andValue1").value;
      const condition2 = document.getElementById("andCondition2").value;
      const operator2 = document.getElementById("andOperator2").value;
      const value2 = document.getElementById("andValue2").value;

      const range = sheet.getRange(cell);
      range.formulas = [[`=IF(AND(${condition1}${operator1}${value1}, ${condition2}${operator2}${value2}), "${trueValue}", "${falseValue}")`]];
      await context.sync();
      console.log(`Inserted AND formula at ${cell}`);
    });
  } catch (error) {
    console.error("Error inserting AND logic:", error);
  }
}

async function insertIfLogic() {
  try {
    console.log("Inserting IF logic");
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const cell = document.getElementById("ifCell").value;
      const trueValue = document.getElementById("ifTrueValue").value;
      const falseValue = document.getElementById("ifFalseValue").value;
      const condition = document.getElementById("ifCondition").value;
      const operator = document.getElementById("ifOperator").value;
      const value = document.getElementById("ifValue").value;

      const range = sheet.getRange(cell);
      range.formulas = [[`=IF(${condition}${operator}${value}, "${trueValue}", "${falseValue}")`]];
      await context.sync();
      console.log(`Inserted IF formula at ${cell}`);
    });
  } catch (error) {
    console.error("Error inserting IF logic:", error);
  }
}

async function insertOrLogic() {
  try {
    console.log("Inserting OR logic");
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const cell = document.getElementById("orCell").value;
      const trueValue = document.getElementById("orTrueValue").value;
      const falseValue = document.getElementById("orFalseValue").value;
      const condition1 = document.getElementById("orCondition1").value;
      const operator1 = document.getElementById("orOperator1").value;
      const value1 = document.getElementById("orValue1").value;
      const condition2 = document.getElementById("orCondition2").value;
      const operator2 = document.getElementById("orOperator2").value;
      const value2 = document.getElementById("orValue2").value;

      const range = sheet.getRange(cell);
      range.formulas = [[`=IF(OR(${condition1}${operator1}${value1}, ${condition2}${operator2}${value2}), "${trueValue}", "${falseValue}")`]];
      await context.sync();
      console.log(`Inserted OR formula at ${cell}`);
    });
  } catch (error) {
    console.error("Error inserting OR logic:", error);
  }
}

async function insertIfsLogic() {
  try {
    console.log("Inserting IFS logic");
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const cell = document.getElementById("ifsCell").value;
      let formula = "=IFS(";
      const conditions = document.querySelectorAll("[id^=ifsCondition]");
      const operators = document.querySelectorAll("[id^=ifsOperator]");
      const values = document.querySelectorAll("[id^=ifsValue]");
      const trueValues = document.querySelectorAll("[id^=ifsTrueValue]");

      console.log("Conditions:", conditions);
      console.log("Operators:", operators);
      console.log("Values:", values);
      console.log("True Values:", trueValues);

      for (let i = 0; i < conditions.length; i++) {
        formula += `${conditions[i].value}${operators[i].value}${values[i].value}, "${trueValues[i].value}"`;
        if (i < conditions.length - 1) {
          formula += ", ";
        }
      }
      formula += ")";
      
      console.log("Generated formula:", formula);

      const range = sheet.getRange(cell);
      range.formulas = [[formula]];
      await context.sync();
      console.log(`Inserted IFS formula at ${cell}`);
    });
  } catch (error) {
    console.error("Error inserting IFS logic:", error);
  }
}

let ifsConditionCount = 1;

function addIfsCondition() {
  const container = document.getElementById("ifsConditionsContainer");
  ifsConditionCount++;

  const conditionDiv = document.createElement("div");
  conditionDiv.id = `ifsConditionDiv${ifsConditionCount}`;

  const conditionInput = document.createElement("input");
  conditionInput.type = "text";
  conditionInput.id = `ifsCondition${ifsConditionCount}`;

  const operatorSelect = document.createElement("select");
  operatorSelect.id = `ifsOperator${ifsConditionCount}`;
  operatorSelect.innerHTML = `
    <option value=">">&gt; greater than</option>
    <option value="<">&lt; less than</option>
    <option value=">=">&gt;= greater than or equal to</option>
    <option value="<=">&lt;= less than or equal to</option>
    <option value="<>">&lt;&gt; not equal to</option>
    <option value="=">= equal to</option>
  `;

  const valueInput = document.createElement("input");
  valueInput.type = "text";
  valueInput.id = `ifsValue${ifsConditionCount}`;

  const trueValueInput = document.createElement("input");
  trueValueInput.type = "text";
  trueValueInput.id = `ifsTrueValue${ifsConditionCount}`;

  const deleteButton = document.createElement("button");
  deleteButton.className = "ms-Button";
  deleteButton.id = `ifsDelete${ifsConditionCount}`;
  deleteButton.innerHTML = `<span class="ms-Button-label">Delete</span>`;
  deleteButton.onclick = () => deleteIfsCondition(ifsConditionCount);

  conditionDiv.appendChild(conditionInput);
  conditionDiv.appendChild(operatorSelect);
  conditionDiv.appendChild(valueInput);
  conditionDiv.appendChild(trueValueInput);
  conditionDiv.appendChild(deleteButton);
  conditionDiv.appendChild(document.createTextNode(", "));

  container.appendChild(conditionDiv);
  container.appendChild(document.createElement("br"));

  console.log(`Added IFS condition ${ifsConditionCount}`);
}

function deleteIfsCondition(index) {
  const container = document.getElementById("ifsConditionsContainer");
  const conditionDiv = document.getElementById(`ifsConditionDiv${index}`);

  if (conditionDiv) {
    container.removeChild(conditionDiv.nextSibling); // remove the <br> element
    container.removeChild(conditionDiv);
    console.log(`Deleted IFS condition ${index}`);
  } else {
    console.error(`Condition div not found for index ${index}`);
  }
}

