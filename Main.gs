////////////////////////////
//        VARIABLES       //______________________________________________________________________________________________________________________________________
////////////////////////////

///// SHEETS ///// ____________________________________________________________________________________________________________________________

const ss = SpreadsheetApp.getActiveSpreadsheet();

const sheetFont = 'Roboto'

const setupName = 'Setup';
const backlogName = 'Backlog';
const planningName = 'Planning';
const sprintSheetName = 'Current Sprint';
const scopeName = 'Scope';
const timeManagementName = 'Time Management';
const budgetSetupName = 'Budget Setup';
const costsName = 'Costs';
const incomesName = 'Incomes';
const timelineName = 'Budget Timeline';
const annualBudgetName = 'Annual Budgets';
const productionBudgetName = 'Stage Budgets';
const customBudgetName = 'Custom Budget'
const importanceScopeName = 'ImportanceScope';
const productionScopeName = 'ProductionScope'

///// SETUP ///// ____________________________________________________________________________________________________________________________

const setupSheet = ss.getSheetByName(setupName);

const setupMinRow = 100;
const setupLength = Math.max(setupMinRow, setupSheet.getMaxRows());


const setupTabsRow = 1;
const setupTabsCol = 2;
const setupTabsWidth = 7;

const parameterName = 'Parameter';
const jobsName = 'Jobs';
const teamName = 'Team';
const productionStageName = 'Production Stage';
const importanceName = 'Importance'
const satusName = 'Status';
const monthsName = 'Months';

const darkModeName = 'Dark Mode';
const highlightColorsName = 'Hilight Colors';
const autoBorderName = 'Backlog Autoborders';
const autoFillName = 'Backlog Auto Fill Empty Row';



///// BACKLOG ///// ____________________________________________________________________________________________________________________________

const backlogSheet = ss.getSheetByName(backlogName);

const backlogTabsRow = 1;
const backlogTabsCol = 1;
const backlogTabsWidth = 13;
const backlogMinRow = 1000;
const backlogLength = Math.max(backlogMinRow, backlogSheet.getMaxRows());

const epicName = 'Epic';
const storiesName = 'Stories';
const taskName = 'Tasks'
const nbDayEstName = 'Nb of days estimated';
const nbDayEst30Name = 'Nb of days estimated + 30%';
const nbDayWorkedName = 'Nb of days worked';
const priorityName = 'Priority';
const sprintName = 'Sprint';
const backlogTeamName = 'In Charge';

///// PLANNING ///// ____________________________________________________________________________________________________________________________

const planningSheet = ss.getSheetByName(planningName);

const planningRow = 1;
const planningCol = 1;

const planningMinLength = 100
const planningLength = Math.max(planningMinLength, planningSheet.getMaxColumns());
const planningWidth = 14
const planningWeeksRow = 1;
const planningStartingDateRow = 2;
const planningTitleColumn = 'A';
const planningValueColumn = 'B';

const planningfirstRow = 4;

const planningCellsWidth = 101;
const planningQuarterCellsHeigth = 50;
const planningCellsHeigth = 30;

const planningStartingDaysName = 'Starting             Day'

///// SPRINT ///// ____________________________________________________________________________________________________________________________

const sprintSheet = ss.getSheetByName(sprintSheetName);

const sprintRow = 1;
const sprintCol = 1;
const sprintMinLength = 100;
const sprintLength = Math.max(sprintMinLength, sprintSheet.getMaxRows());
const sprintTabsWidth = 6;

const sprintStartingDayColumn = 'E';
const sprintTitleColumn = 'D';
const sprintDayGradientColumn = 'F';
const sprintDayGradientStart = 1
const sprintDayGradientEnd = 2

const sprintTitleRow = 1;
const sprintValueRow = 2;
const sprintTabsRow = 3;

const sprintNbWeeks = 5;

///// SCOPE ///// ____________________________________________________________________________________________________________________________

const scopeSheet = ss.getSheetByName(scopeName);

const scopeRow = 1;
const scopeCol = 2;
const scopeMinLength = 500;
const scopeMinWidth = 26

const scopeLength = Math.max(sprintMinLength, scopeSheet.getMaxRows());
const scopeWidth = Math.max(scopeMinWidth, scopeSheet.getMaxColumns());;
const scopeTabsRow = 3;
const scopeVerticalTabsWidth = 3;
const scopeColorFadingCoeff = 0.5;

const scopeTitleColumn = 'B';
const scopeValueColumn = 'C';
const scopeBorderRow = 499;
const scopeBackgroundStyleRow = 500;

const grandTotalName = 'Grand Total'

///// SECONDARY SCOPES ///// ____________________________________________________________________________________________________________________________

const importanceScopeSheet = ss.getSheetByName(importanceScopeName);

const productionScopeSheet = ss.getSheetByName(productionScopeName);

///// TIME MANAGEMENT ///// ____________________________________________________________________________________________________________________________

const timeManagementSheet = ss.getSheetByName(timeManagementName);

const timeManagementLength = 500;
const timeManagementWidth = 27;
const timeManagementCol = 2;

const timeManagementMainBarRow = 4;
const timeManagementTabsRow = 7;
const timeManagementMainBarColumn = 'G';
const timeManagementMainBarWidth = 8;

const timemanagementSprintParamTitleColumn = 'R';
const timeManagementSprintParamValueColumn = 'S';
const timeManagementSprintParamIdealRow = 2;
const timeManagementSprintParamMaxRow = 3;

///// BUDGET SETUP ///// ____________________________________________________________________________________________________________________________

const budgetSetupSheet = ss.getSheetByName(budgetSetupName);

const budgetSetupMinLength = 100;
const budgetSetupLength = Math.max(budgetSetupMinLength, budgetSetupSheet.getMaxRows());
const budgetSetupWidth = 7;
const budgetSetupTabsWidth = 4;
const budgetSetupCol = 2;
const budgetSetupRow = 1;

const probabilityName = 'Probability';
const costSourcesName = 'Cost Sources';
const incomeSourcesName = 'Income Sources';
const sourcesName = 'Sources';
const totalsName = 'Totals';
const stageName = 'Stage';

///// COSTS & INCOMES ///// ____________________________________________________________________________________________________________________________

const startingDateName = 'Starting Date';
const endingDateName = 'Ending Date';
const dateName = 'Date';

///// COSTS ///// ____________________________________________________________________________________________________________________________

const costsSheet = ss.getSheetByName(costsName);

const costsMinWidth = 26;
const costsWidth = Math.max(costsMinWidth, costsSheet.getMaxColumns());

const costsMinLength = 100;
const costsLength = Math.max(costsMinLength, costsSheet.getMaxRows());

const costsCol = 2;
const costsTabsRow = 4;

const costsSalaryWidth = 9;
const costsPonctualWidth = 5;
const fixCostsWidth = 6;

const salariesName = 'Salaries';
const ponctualCostsName = 'Ponctual Costs';
const fixCostsName = 'Monthly Costs : Fix';
const variableCostsName = 'Monthly Costs : Variable';
const permanentContractName = 'Permanent Contract';

const costNameName = 'Cost Name'
const costValueName = 'Cost'

const salariesMonthlyCostName = 'Monthly Cost';
const salariesTeamName = 'Member or Team';

///// INCOMES ///// ____________________________________________________________________________________________________________________________

const incomesSheet = ss.getSheetByName(incomesName);

const incomesMinWidth = 26;
const incomesWidth = Math.max(costsMinWidth, costsSheet.getMaxColumns());

const incomesMinLength = 100;
const incomesLength = Math.max(costsMinLength, costsSheet.getMaxRows());

const incomesCol = 2;
const incomesTabsRow = 4;

const incomesPonctualWidth = 5;
const fixIncomesWidth = 6;

const ponctualIncomesName = 'Ponctual Incomes';
const fixIncomesName = 'Monthly Incomes : Fix';
const variableIncomesName = 'Monthly Incomes : Variable';

const incomeNameName = 'Income Name'
const incomeValueName = 'Income'

///// BUDGET TIMELINE ///// ____________________________________________________________________________________________________________________________

const budgetTimelineSheet = ss.getSheetByName(timelineName);

const budgetTimelineRow = 2;
const timelineFirstColumn = 'G'

const budgetTimelineMinLength = 100;
const budgetTimelineLength = Math.max(budgetTimelineMinLength, budgetTimelineSheet.getMaxRows());

const budgetTimelineMinWidth = 10;
const budgetTimelineWidth = Math.max(budgetTimelineMinWidth, budgetTimelineSheet.getMaxColumns());

const budgetTimelineTitleColumn = 'B';
const budgetTimelineParam1Column = 'C';
const budgetTimelineParam2Column = 'D';
const budgetTimelineParam3Column = 'E';
const budgetTimelineBarLength = 8;
const budgetTimelineGradientLength = 2;
const budgetTimelineCurveLength = 3;

const budgetTimelineStartingDateName = 'Starting Date :';
const budgetTimelineDurationName = 'Duration :';
const budgetTimelineInitialFundName = 'Initial fund :';
const budgetTimelineFundGradientName = 'Fund Gradient';

const fundsName = 'Funds';
const gainsName = 'Gains';

///// BUDGETS ///// ____________________________________________________________________________________________________________________________

const budgetWidth = 8;
const budgetChartWidth = 6;
const budgetChartLength = 12;

const namesName = 'Names';
const totalName = 'Total';
const totalCostsName = 'Total Costs';
const totalIncomesName = 'Total Incomes';
const balanceName = 'Balance :';

///// ANNUAL BUDGETS ///// ____________________________________________________________________________________________________________________________

const annualBudgetSheet = ss.getSheetByName(annualBudgetName);

const annualBudgetRow = 2;
const annualBudgetFirstCol = 6;
const annualBudgetParamTitleColumn = 'B';
const annualBudgetParam1Column = 'C';
const annualBudgetParam2Column = 'D';
const annualBudgetLength = 100;

///// PRODUCTION BUDGETS ///// ____________________________________________________________________________________________________________________________

const productionBudgetSheet = ss.getSheetByName(productionBudgetName);

const productionBudgetRow = 2;
const productionBudgetFirstCol = 8;
const productionBudgetParamTitleColumn = 'B';
const productionBudgetParam1Column = 'C';
const productionBudgetParam2Column = 'D';
const productionBudgetParam3Column = 'E';
const productionBudgetLength = 100;

///// CUSTOM BUDGETS ///// ____________________________________________________________________________________________________________________________

const customBudgetSheet = ss.getSheetByName(customBudgetName);

const customBudgeRow = 2;
const customBudgetFirstCol = 6;
const customBudgetParamTitleColumn = 'B';
const customBudgetParam1Column = 'C';
const customBudgetParam2Column = 'D';
const customBudgetParam3Column = 'E';
const customBudgetLength = 100;

///// COLORS ///// ____________________________________________________________________________________________________________________________

const getHighlightcolor1 = function () {
  const paramColumn = getColumnFromName(setupSheet, parameterName, setupTabsRow, setupTabsCol, setupTabsWidth);
  const paramCol = getColumnFromA1(paramColumn);
  const highlightRow = getRowFromName(setupSheet, highlightColorsName, setupTabsRow, paramCol, setupLength);
  return setupSheet.getRange(paramColumn + (highlightRow + 1)).getBackground().toString();
}

const getHighlightcolor2 = function () {
  const paramColumn = getColumnFromName(setupSheet, parameterName, setupTabsRow, setupTabsCol, setupTabsWidth);
  const paramCol = getColumnFromA1(paramColumn);
  const highlightRow = getRowFromName(setupSheet, highlightColorsName, setupTabsRow, paramCol, setupLength);
  return setupSheet.getRange(paramColumn + (highlightRow + 2)).getBackground().toString();
}

const priotityMax = '10';
const priotityMid = '5';
const priorityMin = '1';

const priorityMaxColor = 'red';
const priorityMidColor = '#b500ff';
const priorityMinColor = 'blue';

const lightGray = '#efefef';
const darkGray = '#191919';
const middleGray = '#2f2f2f';

///// PARAMETERS ///// ____________________________________________________________________________________________________________________________

const isAutoBorder = function () {
  const paramColumn = getColumnFromName(setupSheet, parameterName, setupTabsRow, setupTabsCol, setupTabsWidth);
  const paramCol = getColumnFromA1(paramColumn);
  const autoBorderRow = getRowFromName(setupSheet, autoBorderName, setupTabsRow, paramCol, setupLength);
  return setupSheet.getRange(paramColumn + autoBorderRow).getValue();
}

const isAutoFill = function () {
  const paramColumn = getColumnFromName(setupSheet, parameterName, setupTabsRow, setupTabsCol, setupTabsWidth);
  const paramCol = getColumnFromA1(paramColumn);
  const autoFillRow = getRowFromName(setupSheet, autoFillName, setupTabsRow, paramCol, setupLength);
  return setupSheet.getRange(paramColumn + autoFillRow).getValue();
}

////////////////////////////
//          MAIN          //_______________________________________________________________________________________________________________________________________
////////////////////////////

function Main() {
  costsSheet.getRange(1, 1, costsSheet.getMaxRows(), costsSheet.getMaxColumns()).breakApart()
}

////////////////////////////
//          MENUS         //________________________________________________________________________________________________________________________________________
////////////////////////////

function AddMenu() {
  const menu = SpreadsheetApp.getUi().createMenu('Video Game Project Manager');
  menu
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Global')
      .addItem('Refresh Showdowns', 'RefreshShowdowns')
      .addItem('Refresh Highlights', 'RefreshHighlights')
      .addItem('Switch Dark Mode', 'SwitchDarkMode')
      .addItem('Refresh Dark Mode', 'RefreshDarkMode')
    )
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu(backlogName)
      .addItem('Refesh Backlog Borders', 'TraceBacklogBorders')
      .addItem('Switch Auto Border', 'SwitchAutoBorder')
      .addItem('Switch Auto Fill Empty Row', 'SwitchAutoEmptyRow')
    )

    .addSubMenu(SpreadsheetApp.getUi().createMenu(planningName)
      .addItem('Create Planning', 'CreatePlanning')
    )

    .addSubMenu(SpreadsheetApp.getUi().createMenu('Sprints')
      .addItem('Reset Gantt Dates', 'ResetGanttDates')
      .addItem('Reset Gantt', 'SetSprintConditionnalFomatRules')
    )

    .addSubMenu(SpreadsheetApp.getUi().createMenu(scopeName)
      .addItem('Remove Scope Borders', 'RemoveScopeBorders')
      .addItem('Set Scope Borders', 'CreateScopeBorders')
      .addItem('Change Scope Background Style', 'SwitchScopeBordersStyle')
    )

    .addSubMenu(SpreadsheetApp.getUi().createMenu('Time Management')
      .addItem('Reset Bar Style', 'SetTimeManagementBars')
    )
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Budget Menus')
      .addItem('Refresh Menus', 'RefreshBudgetMenus')
    )
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Budget Timeline')
      .addItem('Create Budget Timeline', 'CreatebudgetTimeline')
      .addItem('Recalculate Timeline Values', 'CalulateBudgetTimelineValues')
    )
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Annual Budgets')
      .addItem('Create Annual Budgets', 'CreateAnnualBudgets')
    )
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Production Budgets')
      .addItem('Create Production Budgets', 'CreateProductionBudgets')
    )
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Custom Budget')
      .addItem('Create Custom Budget', 'CreateCustomBudget')
    )
    .addToUi();
}

////////////////////////////
//     TRIGGER EVENTS     //________________________________________________________________________________________________________________________________________
////////////////////////////

function onOpen(e) {
  AddMenu();
}

function onEdit(e) {
  let sheet = e.source;
  let sheetName = sheet.getActiveSheet().getName();
  let range = e.range;
  let col = range.getColumn();
  let row = range.getRow();
  let value = e.value;

  ///// SETUP /////___________________________________________________________________________

  //Modify the jobs list

  const setupParameterColumn = getColumnFromName(setupSheet, parameterName, setupTabsRow, setupTabsCol, setupTabsWidth);
  const setupParameterCol = getColumnFromA1(setupParameterColumn);

  const setupJobColumn = getColumnFromName(setupSheet, jobsName, setupTabsRow, setupTabsCol, setupTabsWidth);
  const setupJobCol = getColumnFromA1(setupJobColumn);

  const setupTeamColumn = getColumnFromName(setupSheet, teamName, setupTabsRow, setupTabsCol, setupTabsWidth);
  const setupTeamCol = getColumnFromA1(setupTeamColumn);

  const setupDarkModeRow = getRowFromName(setupSheet, darkModeName, setupTabsRow, setupTabsCol, setupTabsWidth);

  //Update Showdowns
  if (sheetName == setupName && (col == setupJobCol || col == setupTeamCol) && row > setupTabsRow) {
    RefreshShowdowns();
  }

  //Dark Mode
  if (sheetName == setupName && col == setupParameterCol && row == setupDarkModeRow) {
    SetDarkModes(GetDarkModeValue());
  }

  ///// BACKLOG /////___________________________________________________________________________

  const backlogEpicColumn = getColumnFromName(backlogSheet, epicName, backlogTabsRow, backlogTabsCol, backlogTabsWidth);
  const backlogEpicCol = getColumnFromA1(backlogEpicColumn);

  const backlogStoriesColumn = getColumnFromName(backlogSheet, storiesName, backlogTabsRow, backlogTabsCol, backlogTabsWidth);
  const backlogStoriesCol = getColumnFromA1(backlogStoriesColumn);

  const backlogSprintColumn = getColumnFromName(backlogSheet, sprintName, backlogTabsRow, backlogTabsCol, backlogTabsWidth);
  const backlogSprintCol = getColumnFromA1(backlogSprintColumn);

  //Auto Borders
  if ((sheetName == backlogName && (col == backlogEpicCol || col == backlogStoriesCol) && row > 1 && GetAutoBorderValue())) {
    const epicValue = backlogSheet.getRange(backlogEpicColumn + row).getValue();
    const storyValue = backlogSheet.getRange(backlogStoriesColumn + row).getValue();

    if (epicValue != '' || storyValue != '') TraceBacklogBorders();
  }

  //Sprint List Update
  if (sheetName == backlogName && col == backlogSprintCol && row > 1) {
    SetSprintShowdown();
  }

  //AutoFill
  if (sheetName == backlogName && col == backlogEpicCol && row > 1 && GetAutoFillValue() == true) {
    AutoFillBacklogEmptyRow(row)
    if (GetAutoBorderValue()) TraceBacklogBorders();
  }

  ///// SPRINT /////___________________________________________________________________________

  const sprintTitleCol = getColumnFromA1(sprintTitleColumn)
  const sprintStartingDayCol = getColumnFromA1(sprintStartingDayColumn);

  //Changing current sprint
  if (sheetName == sprintSheetName && col == sprintTitleCol && row == sprintValueRow) {
    ChangePivotTableFilter();
    SetSprintDuration();
    ResetGanttDates();
  }

  //Update Dates
  if (sheetName == sprintSheetName && col == sprintStartingDayCol && row == sprintValueRow) {
    SetMonday();
    ResetGanttDates();
  }

  if (sheetName == sprintSheetName && col <= sprintTabsWidth && row > sprintTabsRow) {
    SetSprintConditionnalFomatRules();
  }

  ///// SCOPE /////___________________________________________________________________________


  const scopeValueCol = getColumnFromA1(scopeValueColumn)

  //Switch Scope Border
  if (sheetName == scopeName && col == scopeValueCol && row == scopeBorderRow) {
    SetScopeBorders();
  }

  //Change Background Style
  if (sheetName == scopeName && col == scopeValueCol && row == scopeBackgroundStyleRow) {
    SetScopeConditionnalFormatingRules();
  }
}

////////////////////////////
//        UTILITIES       //_______________________________________________________________________________________________________________________________________
////////////////////////////

///// DATAS /////___________________________________________________________________________

function transposeArray(array2D) {
  return Object.keys(array2D[0]).map(function (column) {
    return array2D.map(function (row) {
      return row[column];
    });
  });
}

function cleanData(columnRangeValues) {
  return transposeArray(columnRangeValues)[0];
}

///// A1NOTATION /////___________________________________________________________________________

const getA1Notation = (row, column) => {
  row--;
  column--;
  const a1Notation = [`${row + 1}`];
  const totalAlphabets = 'Z'.charCodeAt() - 'A'.charCodeAt() + 1;
  let block = column;
  while (block >= 0) {
    a1Notation.unshift(String.fromCharCode((block % totalAlphabets) + 'A'.charCodeAt()));
    block = Math.floor(block / totalAlphabets) - 1;
  }
  return a1Notation.join('');
};

const getR1C1Notation = (row, column, numberOfRow, numberOfColumn) => {
  return (getA1Notation(row, column) + ':' + getA1Notation(row + numberOfRow - 1, column + numberOfColumn - 1));
}

const getA1FromCol = (column) => {
  column--;
  const a1Notation = [];
  const totalAlphabets = 'Z'.charCodeAt() - 'A'.charCodeAt() + 1;
  let block = column;
  while (block >= 0) {
    a1Notation.unshift(String.fromCharCode((block % totalAlphabets) + 'A'.charCodeAt()));
    block = Math.floor(block / totalAlphabets) - 1;
  }
  return a1Notation.join('');
}

const getColumnFromA1 = (a1Notation) => {
  const totalAlphabets = 'Z'.charCodeAt() - 'A'.charCodeAt() + 1;
  let column = 0;
  const letters = a1Notation.toUpperCase().split('').reverse();
  letters.forEach((letter, index) => {
    const value = letter.charCodeAt() - 'A'.charCodeAt() + 1;
    column += value * Math.pow(totalAlphabets, index);
  });
  return column;
}

const getColumnFromName = (sheet, name, row, column, numberOfColumn) => {
  const range = sheet.getRange(row, column, 1, numberOfColumn);

  const values = range.getValues();

  return getA1FromCol(column + values[0].indexOf(name));
}

const getRowFromName = (sheet, name, row, column, numberOfRow) => {
  const range = sheet.getRange(getR1C1Notation(row, column, numberOfRow, 1))
  let values = cleanData(range.getValues());
  return (values.indexOf(name) + row);
}

const getNumbersFromR1C1 = (r1c1Notation) => {

  var regex = /([A-Z]+)(\d+):([A-Z]+)(\d+)/;

  // Extracting row and column information using regex
  var matches = r1c1Notation.match(regex);

  // Extracted values
  const startColumn = matches[1];
  const startRow = parseInt(matches[2]);
  const endColumn = matches[3];
  const endRow = parseInt(matches[4]);

  const startColumnNumber = getColumnFromA1(startColumn);
  const endColumnNumber = getColumnFromA1(endColumn);

  // Calculating the number of rows and columns
  const numRows = endRow - startRow + 1;
  const numColumns = endColumnNumber - startColumnNumber + 1;

  return [startRow, startColumnNumber, numRows, numColumns];
}

const getColumnOffset = (column, offset) => {
  return getA1FromCol(getColumnFromA1(column) + offset);
}

///// SEQUENCE /////___________________________________________________________________________

function GetSequence(values) {
  const occurrences = values.reduce(function (acc, curr) {
    acc[curr] = (acc[curr] || 0) + 1;
    return acc;
  }, {});
  const sequence = Object.values(occurrences);
  return sequence;
}

function GetSequenceWithEmpty(values) {
  const sequence = values.reduce(function (acc, curr) {
    if (curr[0] !== '') {
      acc.push(1);
    } else if (acc.length > 0) {
      acc[acc.length - 1]++;
    };
    return acc;
  }, []);

  return sequence;
};

function GetSequenceIgnoreEmpty(values) {
  const sequence = [];
  let acc = 0;
  let lastValue = values[0][0];
  for (let k = 0; k < values.length; k++) {
    if (values[k][0] == lastValue || values[k][0] == '') {
      acc++;
    }
    else {
      sequence.push(acc);
      acc = 0;
      if (k != (values.length - 1)) {
        lastValue = values[k + 1][0];
      };
    };
  };
  sequence.push(0);
  return sequence;
}

///// MISC /////___________________________________________________________________________

function ResetBorders(range) {
  range.setBorder(true, true, true, true, true, true, GetDarkModeColor(), SpreadsheetApp.BorderStyle.SOLID);
}

const pSBC = (p, c0, c1, l) => {
  let r, g, b, P, f, t, h, i = parseInt, m = Math.round, a = typeof (c1) == "string";
  if (typeof (p) != "number" || p < -1 || p > 1 || typeof (c0) != "string" || (c0[0] != 'r' && c0[0] != '#') || (c1 && !a)) return null;
  if (!this.pSBCr) this.pSBCr = (d) => {
    let n = d.length, x = {};
    if (n > 9) {
      [r, g, b, a] = d = d.split(","), n = d.length;
      if (n < 3 || n > 4) return null;
      x.r = i(r[3] == "a" ? r.slice(5) : r.slice(4)), x.g = i(g), x.b = i(b), x.a = a ? parseFloat(a) : -1;
    } else {
      if (n == 8 || n == 6 || n < 4) return null;
      if (n < 6) d = "#" + d[1] + d[1] + d[2] + d[2] + d[3] + d[3] + (n > 4 ? d[4] + d[4] : "");
      d = i(d.slice(1), 16);
      if (n == 9 || n == 5) x.r = d >> 24 & 255, x.g = d >> 16 & 255, x.b = d >> 8 & 255, x.a = m((d & 255) / 0.255) / 1000;
      else x.r = d >> 16, x.g = d >> 8 & 255, x.b = d & 255, x.a = -1;
    } return x;
  };
  h = c0.length > 9, h = a ? c1.length > 9 ? true : c1 == "c" ? !h : false : h, f = this.pSBCr(c0), P = p < 0, t = c1 && c1 != "c" ? this.pSBCr(c1) : P ? { r: 0, g: 0, b: 0, a: -1 } : { r: 255, g: 255, b: 255, a: -1 }, p = P ? p * -1 : p, P = 1 - p;
  if (!f || !t) return null;
  if (l) r = m(P * f.r + p * t.r), g = m(P * f.g + p * t.g), b = m(P * f.b + p * t.b);
  else r = m((P * f.r ** 2 + p * t.r ** 2) ** 0.5), g = m((P * f.g ** 2 + p * t.g ** 2) ** 0.5), b = m((P * f.b ** 2 + p * t.b ** 2) ** 0.5);
  a = f.a, t = t.a, f = a >= 0 || t >= 0, a = f ? a < 0 ? t : t < 0 ? a : a * P + t * p : 0;
  if (h) return "rgb" + (f ? "a(" : "(") + r + "," + g + "," + b + (f ? "," + m(a * 1000) / 1000 : "") + ")";
  else return "#" + (4294967296 + r * 16777216 + g * 65536 + b * 256 + (f ? m(a * 255) : 0)).toString(16).slice(1, f ? undefined : -2);
}

function lightenDarkenColor(color, coeff) {
  const isDarkMode = GetDarkModeValue();
  if (isDarkMode) return pSBC(-coeff, color);
  else return pSBC(coeff, color);
}

/* let color1 = "rgb(20,60,200)";
let color2 = "rgba(20,60,200,0.67423)";
let color3 = "#67DAF0";
let color4 = "#5567DAF0";
let color5 = "#F3A";
let color6 = "#F3A9";
let color7 = "rgb(200,60,20)";
let color8 = "rgba(200,60,20,0.98631)";

// Tests:

// Log Blending
// Shade (Lighten or Darken)
pSBC ( 0.42, color1 ); // rgb(20,60,200) + [42% Lighter] => rgb(166,171,225)
pSBC ( -0.4, color5 ); // #F3A + [40% Darker] => #c62884
pSBC ( 0.42, color8 ); // rgba(200,60,20,0.98631) + [42% Lighter] => rgba(225,171,166,0.98631)

// Shade with Conversion (use "c" as your "to" color)
pSBC ( 0.42, color2, "c" ); // rgba(20,60,200,0.67423) + [42% Lighter] + [Convert] => #a6abe1ac

// RGB2Hex & Hex2RGB Conversion Only (set percentage to zero)
pSBC ( 0, color6, "c" ); // #F3A9 + [Convert] => rgba(255,51,170,0.6)

// Blending
pSBC ( -0.5, color2, color8 ); // rgba(20,60,200,0.67423) + rgba(200,60,20,0.98631) + [50% Blend] => rgba(142,60,142,0.83)
pSBC ( 0.7, color2, color7 ); // rgba(20,60,200,0.67423) + rgb(200,60,20) + [70% Blend] => rgba(168,60,111,0.67423)
pSBC ( 0.25, color3, color7 ); // #67DAF0 + rgb(200,60,20) + [25% Blend] => rgb(134,191,208)
pSBC ( 0.75, color7, color3 ); // rgb(200,60,20) + #67DAF0 + [75% Blend] => #86bfd0

// Linear Blending
// Shade (Lighten or Darken)
pSBC ( 0.42, color1, false, true ); // rgb(20,60,200) + [42% Lighter] => rgb(119,142,223)
pSBC ( -0.4, color5, false, true ); // #F3A + [40% Darker] => #991f66
pSBC ( 0.42, color8, false, true ); // rgba(200,60,20,0.98631) + [42% Lighter] => rgba(223,142,119,0.98631)

// Shade with Conversion (use "c" as your "to" color)
pSBC ( 0.42, color2, "c", true ); // rgba(20,60,200,0.67423) + [42% Lighter] + [Convert] => #778edfac

// RGB2Hex & Hex2RGB Conversion Only (set percentage to zero)
pSBC ( 0, color6, "c", true ); // #F3A9 + [Convert] => rgba(255,51,170,0.6)

// Blending
pSBC ( -0.5, color2, color8, true ); // rgba(20,60,200,0.67423) + rgba(200,60,20,0.98631) + [50% Blend] => rgba(110,60,110,0.83)
pSBC ( 0.7, color2, color7, true ); // rgba(20,60,200,0.67423) + rgb(200,60,20) + [70% Blend] => rgba(146,60,74,0.67423)
pSBC ( 0.25, color3, color7, true ); // #67DAF0 + rgb(200,60,20) + [25% Blend] => rgb(127,179,185)
pSBC ( 0.75, color7, color3, true ); // rgb(200,60,20) + #67DAF0 + [75% Blend] => #7fb3b9

// Other Stuff 
// Error Checking
pSBC ( 0.42, "#FFBAA" ); // #FFBAA + [42% Lighter] => null  (Invalid Input Color)
pSBC ( 42, color1, color5 ); // rgb(20,60,200) + #F3A + [4200% Blend] => null  (Invalid Percentage Range)
pSBC ( 0.42, {} ); // [object Object] + [42% Lighter] => null  (Strings Only for Color)
pSBC ( "42", color1 ); // rgb(20,60,200) + ["42"] => null  (Numbers Only for Percentage)
pSBC ( 0.42, "salt" ); // salt + [42% Lighter] => null  (A Little Salt is No Good...)

// Error Check Fails (Some Errors are not Caught)
pSBC ( 0.42, "#salt" ); // #salt + [42% Lighter] => #a5a5a500  (...and a Pound of Salt is Jibberish)

// Ripping
pSBCr ( color4 ); // #5567DAF0 + [Rip] => [object Object] => {'r':85,'g':103,'b':218,'a':0.941} */

////////////////////////////
//        FORMATTING      //_______________________________________________________________________________________________________________________________________
////////////////////////////

function GetConditionnalFormattingRulesExact(setupRange, targetRange) {

  const values = GetValues(setupRange);
  const colors = GetColors(setupRange);
  const textStyles = GetTextStyles(setupRange);

  const rules = [];

  for (let index in values) {
    if (values[index] != null) {
      const newFormat = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(values[index])
        .setBackground(colors[index])
        .setFontColor(textStyles[index].getForegroundColor())
        .setBold(textStyles[index].isBold())
        .setRanges([targetRange])
        .build();
      rules.push(newFormat);
    };
  };
  return rules;
}

function GetConditionnalFormattingRulesContain(setupRange, targetRange) {

  const values = GetValues(setupRange);
  const colors = GetColors(setupRange);
  const textStyles = GetTextStyles(setupRange);

  const rules = [];

  for (let index in values) {
    if (values[index] != null) {
      const newFormat = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains(values[index])
        .setBackground(colors[index])
        .setFontColor(textStyles[index].getForegroundColor())
        .setBold(textStyles[index].isBold())
        .setRanges([targetRange])
        .build()
      rules.push(newFormat);
    };
  };

  return rules;
}

function GetColors(range) {
  const colors = cleanData(range.getBackgrounds());
  colors.length = GetValues(range).length;
  return colors;
}

function GetTextStyles(range) {
  const textStyles = cleanData(range.getTextStyles());
  textStyles.length = GetValues(range).length;
  return textStyles;
}

function GetValues(range) {
  const values = cleanData(range.getValues().filter(String));
  return values;
}

////////////////////////////
//        DARK MODE       //___________________________________________________________________________________________________________________________________________
////////////////////////////

///// HIGHLIGHT /////___________________________________________________________________________________________________________________________________________________________________________

function SetHighlight(sheet, textRanges1, borderRanges1, textRanges2, borderRanges2) {

  // Text Highlights 1
  if (textRanges1.length != 0) {
    const textRangeList1 = sheet.getRangeList(textRanges1);
    textRangeList1.setFontColor(getHighlightcolor1());
  };

  // Text Highlights 2
  if (textRanges2.length != 0) {
    const textRangeList2 = sheet.getRangeList(textRanges2);
    textRangeList2.setFontColor(getHighlightcolor2());
  };

  // Border Highlights 1
  if (borderRanges2.length != 0) {
    const borderRangeList2 = sheet.getRangeList(borderRanges2);
    borderRangeList2.setBorder(true, true, true, true, null, null, getHighlightcolor2(), SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  };

  // Border Highlights 2
  if (borderRanges1.length != 0) {
    const borderRangeList1 = sheet.getRangeList(borderRanges1);
    borderRangeList1.setBorder(true, true, true, true, null, null, getHighlightcolor1(), SpreadsheetApp.BorderStyle.SOLID_THICK);
  };
}

///// BACKGROUND /////___________________________________________________________________________________________________________________________________________________________________________

function ClearBandings(sheet) {
  bandings = sheet.getBandings()
  for (let banding of bandings) {
    banding.remove();
  }
}

function SetDarkModeBandings(sheet, ranges) {
  ClearBandings(sheet);
  for (let range of ranges) {
    sheet.getRange(range).applyRowBanding()
      .setHeaderRowColor(GetDarkModeColor())
      .setFirstRowColor(GetDarkModeColor())
      .setSecondRowColor(GetDarkModeGrayColor())
  };
}

function SetBackgroundDarkMode(sheet, backgroundRanges, grayBackgroundRanges, textRanges) {

  // BackGround
  if (backgroundRanges.length != 0)
    sheet.getRangeList(backgroundRanges).setBackground(GetDarkModeColor());

  // Gray Area
  if (grayBackgroundRanges.length != 0)
    sheet.getRangeList(grayBackgroundRanges).setBackground(GetDarkModeGrayColor());

  // Texts
  if (textRanges.length != 0)
    sheet.getRangeList(textRanges).setFontColor(GetDarkModeTextColor());
}

///// BORDER /////___________________________________________________________________________________________________________________________________________________________________________

function SetBordersDarkMode(sheet, borderRanges, backgroundBorderRanges, grayBorderRanges, rowsBorderRanges, columnBorderRanges) {

  // Get whole sheet Range
  const sheetRange = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());

  // Set background borders
  sheetRange.setBorder(true, true, true, true, true, true, GetDarkModeColor(), SpreadsheetApp.BorderStyle.SOLID);

  // Set background borders
  if (backgroundBorderRanges.length != 0) {
    const backgroundBorderRangesList = sheet.getRangeList(backgroundBorderRanges);
    backgroundBorderRangesList.setBorder(true, true, true, true, null, null, GetDarkModeColor(), SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  };

  // Set gray borders
  if (grayBorderRanges.length != 0) {
    const grayBorderRangeRangeList = sheet.getRangeList(grayBorderRanges);
    grayBorderRangeRangeList.setBorder(true, true, true, true, null, null, GetDarkModeGrayBorderColor(), SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  };

  // Set borders
  if (borderRanges.length != 0) {
    const borderRangeList = sheet.getRangeList(borderRanges);
    borderRangeList.setBorder(true, true, true, true, null, null, GetDarkModeTextColor(), SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  };

  // Set rows borders
  if (rowsBorderRanges.length != 0) {
    const rowsBorderRangeRangeList = sheet.getRangeList(rowsBorderRanges);
    rowsBorderRangeRangeList.setBorder(null, null, null, null, null, true, GetDarkModeGrayBorderColor(), SpreadsheetApp.BorderStyle.SOLID);
  };

  // Set columns borders
  if (columnBorderRanges.length != 0) {
    const columnBorderRangeList = sheet.getRangeList(columnBorderRanges);
    columnBorderRangeList.setBorder(null, null, null, null, true, null, GetDarkModeTextColor(), SpreadsheetApp.BorderStyle.SOLID);
  };

}

///// COLORS /////___________________________________________________________________________________________________________________________________________________________________________

function GetDarkModeColor() {
  range = setupSheet.getRange(GetDarkModeCellA1());
  isDarkMode = range.getValue();
  if (isDarkMode) return '#000000';
  else return '#ffffff';
}

function GetDarkModeTextColor() {
  range = setupSheet.getRange(GetDarkModeCellA1());
  isDarkMode = range.getValue();
  if (isDarkMode) return '#ffffff';
  else return '#000000';
}

function GetDarkModeGrayColor() {
  range = setupSheet.getRange(GetDarkModeCellA1());
  isDarkMode = range.getValue();
  if (isDarkMode) return darkGray;
  else return lightGray;
}

function GetDarkModeGrayBorderColor() {
  range = setupSheet.getRange(GetDarkModeCellA1());
  isDarkMode = range.getValue();
  if (isDarkMode) return middleGray;
  else return lightGray;
}
///// SETUP /////___________________________________________________________________________________________________________________________________________________________________________

function SetSetupHighlight() {
  const sheet = setupSheet;
  const row = setupTabsRow;
  const col = setupTabsCol;
  const width = setupTabsWidth
  const length = setupLength;

  // Get Columns & Rows

  const lastColumn = getA1FromCol(col + width - 1);

  const parameterColumn = getColumnFromName(sheet, parameterName, row, col, width);
  const teamColumn = getColumnFromName(sheet, teamName, row, col, width);
  const monthsColumn = getColumnFromName(sheet, monthsName, row, col, width);
  const importanceColumn = getColumnFromName(sheet, importanceName, row, col, width);
  const statusColumn = getColumnFromName(sheet, satusName, row, col, width);
  const jobColumn = getColumnFromName(sheet, jobsName, row, col, width);
  const productionColumn = getColumnFromName(sheet, productionStageName, row, col, width);

  const firstColumn = getColumnOffset(parameterColumn, 1);

  const blackModeRow = getRowFromName(sheet, darkModeName, row, col, length);
  const highlightRow = getRowFromName(sheet, highlightColorsName, row, col, setupLength);
  const autoBorderRow = getRowFromName(sheet, autoBorderName, row, col, setupLength);
  const autofillRow = getRowFromName(sheet, autoFillName, row, col, setupLength);

  // Get Ranges

  const textRanges1 = [
    // Parameter Titles
    parameterColumn + blackModeRow,
    parameterColumn + autoBorderRow,
    parameterColumn + autofillRow,
    parameterColumn + highlightRow,
  ];

  const borderRanges1 = [
    // Titles
    parameterColumn + row,
    teamColumn + row,
    monthsColumn + row,
    importanceColumn + row,
    statusColumn + row,
    jobColumn + row,
    productionColumn + row,

    // Tab
    firstColumn + row + ':' + lastColumn + setupLength,
    parameterColumn + row + ':' + parameterColumn + setupLength,

    // Parameter
    parameterColumn + blackModeRow + ':' + parameterColumn + (blackModeRow + 1),
    parameterColumn + highlightRow + ':' + parameterColumn + (highlightRow + 2),
    parameterColumn + autoBorderRow+ ':' + parameterColumn + (autoBorderRow + 1),
    parameterColumn + autofillRow+ ':' + parameterColumn + (autofillRow + 1),
  ];

  const textRanges2 = [
    // Titles
    parameterColumn + row + ':' + lastColumn + row,

    // Parameters
    parameterColumn + (blackModeRow+1),
    parameterColumn + (highlightRow+1),
    parameterColumn + (autoBorderRow+1),
    parameterColumn + (autofillRow+1),
  ];

  const borderRanges2 = [];

  // Set Highlights

  SetHighlight(sheet, textRanges1, borderRanges1, textRanges2, borderRanges2);
}

function SetSetupDarkMode() {
  const sheet = setupSheet
  const row = setupTabsRow;
  const col = setupTabsCol;
  const width = setupTabsWidth
  const length = setupLength;

  // Get Columns & Rows

  const firstColumn = getA1FromCol(col)
  const lastColumn = getA1FromCol(col + width - 1)
  const maxColumn = getA1FromCol(sheet.getMaxColumns())

  const parameterColumn = getColumnFromName(sheet, parameterName, row, col, width);
  const teamColumn = getColumnFromName(sheet, teamName, row, col, width);
  const monthsColumn = getColumnFromName(sheet, monthsName, row, col, width);
  const importanceColumn = getColumnFromName(sheet, importanceName, row, col, width);
  const statusColumn = getColumnFromName(sheet, satusName, row, col, width);
  const jobColumn = getColumnFromName(sheet, jobsName, row, col, width);
  const productionColumn = getColumnFromName(sheet, productionStageName, row, col, width);

  const blackModeRow = getRowFromName(sheet, darkModeName, row, col, length);

  const highlightRow = getRowFromName(sheet, highlightColorsName, row, col, length);
  const autoBorderRow = getRowFromName(sheet, autoBorderName, row, col, length);
  const autofillRow = getRowFromName(sheet, autoFillName, row, col, length);

  // Get Lengths

  const parameterRange = sheet.getRange(parameterColumn + row + ':' + parameterColumn + length)

  const parameterLength = GetValues(parameterRange).length;
  const jobsLength = GetValues(GetJobsRange()).length;
  const teamLength = GetValues(GetTeamRange()).length;
  const productionLength = GetValues(GetProductionRange()).length;
  const monthsLength = GetValues(GetMonthsRange()).length;
  const importanceLength = GetValues(GetImportanceRange()).length;
  const statusLength = GetValues(GetStatusRange()).length;

  // Get Ranges

  const backgroundRanges = [
    firstColumn + row + ':' + lastColumn + row, // Title Rows
    getColumnOffset(firstColumn, -1) + row + ':' + getColumnOffset(firstColumn, -1) + length, // Left Space
    getColumnOffset(lastColumn, 1) + row + ':' + maxColumn + length, // Right Space

    // Parameters
    parameterColumn + (blackModeRow + 1),
    parameterColumn + (autoBorderRow + 1),
    parameterColumn + (autofillRow + 1),
  ];

  const grayBackgroundRanges = [

    // Parameters
    parameterColumn + blackModeRow,
    parameterColumn + autoBorderRow,
    parameterColumn + autofillRow,
    parameterColumn + highlightRow,

    // Lists Bottoms
    parameterColumn + (parameterLength + row + 2) + ':' + parameterColumn + length,
    jobColumn + (jobsLength + row + 1) + ':' + jobColumn + length,
    teamColumn + (teamLength + row + 1) + ':' + teamColumn + length,
    productionColumn + (productionLength + row + 1) + ':' + productionColumn + length,
    monthsColumn + (monthsLength + row + 1) + ':' + monthsColumn + length,
    importanceColumn + (importanceLength + row + 1) + ':' + importanceColumn + length,
    statusColumn + (statusLength + row + 1) + ':' + statusColumn + length,
  ];

  const textRanges = [];

  const borderRanges = [];
  const backgroundBorderRanges = [];
  const grayBorderRanges = [];
  const rowsBorderRanges = [];
  const columnBorderRanges = [];

  // Set Dark Mode

  SetBackgroundDarkMode(sheet, backgroundRanges, grayBackgroundRanges, textRanges);
  SetBordersDarkMode(sheet, borderRanges, backgroundBorderRanges, grayBorderRanges, rowsBorderRanges, columnBorderRanges);
}

///// BACKLOG /////___________________________________________________________________________________________________________________________________________________________________________

function SetBacklogHighlight() {
  const sheet = backlogSheet;
  const row = backlogTabsRow;
  const col = backlogTabsCol;
  const width = backlogTabsWidth;
  const length = backlogLength;

  // Get Columns & Rows

  const firstColumn = getA1FromCol(col);
  const lastColumn = getA1FromCol(col + width - 1);

  const dayWorkedColumn = getColumnFromName(sheet, nbDayWorkedName, row, col, width);
  const dayEstimated30Column = getColumnFromName(sheet, nbDayEst30Name, row, col, width);
  const epicColumn = getColumnFromName(sheet, epicName, row, col, width);
  const storiesColumn = getColumnFromName(sheet, storiesName, row, col, width);

  // Get Ranges

  const textRanges1 = [
    dayWorkedColumn + row + ':' + dayWorkedColumn + length // Days
  ];

  const borderRanges1 = [
    firstColumn + row + ':' + lastColumn + row, // Tabs

    epicColumn + row, // Epic
    storiesColumn + row // Stories
  ];

  const textRanges2 = [
    firstColumn + row + ':' + lastColumn + row, // Titles
    dayEstimated30Column + row + ':' + dayEstimated30Column + length // Days
  ];

  const borderRanges2 = [];

  // Set Highlights

  SetHighlight(sheet, textRanges1, borderRanges1, textRanges2, borderRanges2);
}

function SetBacklogDarkMode() {
  const sheet = backlogSheet;
  const row = backlogTabsRow;
  const col = backlogTabsCol;
  const width = backlogTabsWidth;
  const length = backlogLength;



  // Get Columns & Rows

  const firstColumn = getA1FromCol(col);
  const lastColumn = getA1FromCol(col + width - 1);

  const epicColumn = getColumnFromName(sheet, epicName, row, col, width);
  const storiesColumn = getColumnFromName(sheet, storiesName, row, col, width);

  const maxWidth = sheet.getMaxColumns();
  const maxColumn = getA1FromCol(maxWidth);

  // Get Ranges

  const backgroundRanges = [
    epicColumn + row + ':' + epicColumn + length, // Epic
    getColumnOffset(lastColumn, 1) + row + ':' + maxColumn + length // Right Space
  ];

  const grayBackgroundRanges = [
    storiesColumn + row + ':' + storiesColumn + length // Stories
  ];

  const textRanges = [];

  const alternatingColorsRanges = [
    firstColumn + row + ':' + lastColumn + length // Datas
  ];

  const borderRanges = [];
  const backgroundBorderRanges = [];
  const grayBorderRanges = [];
  const rowsBorderRanges = [];
  const columnBorderRanges = [];


  // Set Dark Mode

  SetDarkModeBandings(sheet, alternatingColorsRanges);
  SetBackgroundDarkMode(sheet, backgroundRanges, grayBackgroundRanges, textRanges);
  SetBordersDarkMode(sheet, borderRanges, backgroundBorderRanges, grayBorderRanges, rowsBorderRanges, columnBorderRanges);
}

///// PLANNING /////___________________________________________________________________________________________________________________________________________________________________________

function SetPlanningHighlight() {
  const sheet = planningSheet;

  // Get Columns & Rows

  const monthRow = planningWeeksRow;
  const startingDateRow = planningStartingDateRow;
  const titleColumn = planningTitleColumn;
  const valueColumn = planningValueColumn;

  // Get Ranges

  const textRanges1 = [titleColumn + monthRow + ':' + titleColumn + startingDateRow]; // Parameter Title

  const borderRanges1 = [valueColumn + monthRow, valueColumn + startingDateRow]; // Parameter Values

  const textRanges2 = [valueColumn + monthRow + ':' + valueColumn + startingDateRow]; // Parameter Values

  const borderRanges2 = [];

  SetHighlight(sheet, textRanges1, borderRanges1, textRanges2, borderRanges2);
}

function SetPlanningDarkMode() {
  const sheet = planningSheet;
  const row = planningRow;
  const col = planningCol;
  const width = planningWidth;
  const length = planningLength;
  const planningLimit = sheet.getLastRow()

  // Get Columns & Rows

  const firstColumn = getA1FromCol(col)
  const lastColumn = getA1FromCol(col + width - 1);
  const maxColumn = getA1Notation(sheet.getMaxColumns())
  const monthRow = planningWeeksRow;
  const startingDateRow = planningStartingDateRow;
  const titleColumn = planningTitleColumn;
  const valueColumn = planningValueColumn;

  // Get Ranges

  const backgroundRanges = [firstColumn + row + ':' + maxColumn + length]; // Whole Sheet

  const grayBackgroundRanges = [valueColumn + monthRow + ':' + valueColumn + startingDateRow]; // Parameter Values

  // Alternating Colors
  for (let k = 4; k < planningLimit; k += 4) {
    backgroundRanges.push(firstColumn + k + ':' + lastColumn + k);
    grayBackgroundRanges.push(firstColumn + (k + 1) + ':' + lastColumn + (k + 3));
  }

  const textRanges = [];

  const borderRanges = [];

  // Get borders cell from properties
  var planningBordersCells = PropertiesService.getScriptProperties().getProperty('planningBordersCells');
  planningBordersCells = JSON.parse(planningBordersCells)
  const backgroundBorderRanges = planningBordersCells;

  const grayBorderRanges = [];
  const rowBorderRanges = [];
  const columnBorderRanges = [];


  // Set Dark Mode

  SetBackgroundDarkMode(sheet, backgroundRanges, grayBackgroundRanges, textRanges);
  SetBordersDarkMode(sheet, borderRanges, backgroundBorderRanges, grayBorderRanges, rowBorderRanges, columnBorderRanges);
}

///// SPRINT /////___________________________________________________________________________________________________________________________________________________________________________

function SetSprintHighlight() {
  const sheet = sprintSheet;
  const row = sprintRow;
  const col = sprintCol;
  const length = sprintLength;
  const tabsWidth = sprintTabsWidth;

  // Get Columns & Rows

  const firstColumn = getA1FromCol(col);
  const lastColumn = getA1FromCol(col + length - 1);
  const tabsLastColumn = getA1FromCol(col + tabsWidth - 1);

  const startingDayColumn = sprintStartingDayColumn;
  const titleColumn = sprintTitleColumn;

  const titleRow = sprintTitleRow;
  const valueRow = sprintValueRow;
  const tabsRow = sprintTabsRow;

  const datesColumn = getColumnFromName(sheet, planningStartingDaysName, tabsRow, col, sheet.getLastColumn());

  // Get Ranges

  const textRanges1 = [
    firstColumn + tabsRow + ':' + tabsLastColumn + tabsRow, // Tabs Titles
    sprintTitleColumn + row + ':' + sprintStartingDayColumn + row, // Params Titles
  ];

  const borderRanges1 = [
    firstColumn + tabsRow + ':' + tabsLastColumn + length, // Tabs
    startingDayColumn + valueRow + ':' + titleColumn + valueRow, // Param Values
    datesColumn + titleRow + ':' + datesColumn + valueRow, // Sprint Duration
    datesColumn + tabsRow + ':' + lastColumn + tabsRow, // Gantt Dates X
    datesColumn + tabsRow + ':' + datesColumn + length, // Gantt Dates Y
  ];

  //Get Weeks Borders
  let startColumn = getColumnOffset(datesColumn, 1)
  for (let k = 0; k < sprintNbWeeks; k++) {
    let endColumn = getColumnOffset(startColumn, 4)
    borderRanges1.push(startColumn + tabsRow + ':' + endColumn + length)
    startColumn = getColumnOffset(endColumn, 1)
  };

  const textRanges2 = [
    startingDayColumn + valueRow + ':' + titleColumn + valueRow, // Param Values
    datesColumn + (tabsRow + 1) + ':' + datesColumn + length, // Gantt Dates Y
  ];

  const borderRanges2 = [];

  // Set Highlights

  SetHighlight(sheet, textRanges1, borderRanges1, textRanges2, borderRanges2);
}

function SetSprintDarkMode() {
  const sheet = sprintSheet;
  const row = sprintRow;
  const col = sprintCol;
  const length = sprintLength;
  const tabsWidth = sprintTabsWidth;

  // Get Columns & Rows

  const firstColumn = getA1FromCol(col);
  const lastColumn = getA1FromCol(col + length - 1);
  const tabsLastColumn = getA1FromCol(col + tabsWidth - 1);

  const startingDayColumn = sprintStartingDayColumn;
  const titleColumn = sprintTitleColumn;

  const titleRow = sprintTitleRow;
  const valueRow = sprintValueRow;
  const tabsRow = sprintTabsRow;

  const priorityColumn = getColumnFromName(sheet, priorityName, tabsRow, col, sheet.getLastColumn());
  const datesColumn = getColumnFromName(sheet, planningStartingDaysName, tabsRow, col, sheet.getLastColumn());
  const taskColumn = getColumnFromName(sheet, taskName, tabsRow, col, length);

  // Get Ranges

  const backgroundRanges = [
    firstColumn + row + ':' + lastColumn + row, // First Row
    firstColumn + (row + 1) + ':' + getColumnOffset(sprintDayGradientColumn, -1) + (row + 1), // Second Row left
    getColumnOffset(datesColumn, 6) + (row + 1) + ':' + lastColumn + (row + 1), // Second Row right

    firstColumn + tabsRow + ':' + lastColumn + tabsRow, // Tabs Titles & Gantt X
    datesColumn + row + ':' + datesColumn + length // Gantt Y
  ];

  const grayBackgroundRanges = [
    startingDayColumn + titleRow + ':' + titleColumn + titleRow, // Param Titles
    getColumnOffset(datesColumn, 1) + valueRow + ':' + getColumnOffset(datesColumn, 5) + valueRow // Weeks
  ];

  const textRanges = [
    firstColumn + (tabsRow + 1) + ':' + tabsLastColumn + length, // Tabs Datas
    datesColumn + sprintValueRow // Sprint Duration
  ];

  const alternatingColorsRanges = [
    firstColumn + tabsRow + ':' + lastColumn + length
  ];

  const borderRanges = [];
  const backgroundBorderRanges = [];
  const grayBorderRanges = [];
  const rowsBorderRanges = [];
  const columnBorderRanges = [];

  // Set Dark Mode

  SetBackgroundDarkMode(sheet, backgroundRanges, grayBackgroundRanges, textRanges);
  SetDarkModeBandings(sheet, alternatingColorsRanges)
  SetBordersDarkMode(sheet, borderRanges, backgroundBorderRanges, grayBorderRanges, rowsBorderRanges, columnBorderRanges)
}

///// SCOPE /////___________________________________________________________________________________________________________________________________________________________________________

function SetScopeHighlights() {
  const sheet = scopeSheet
  const row = scopeRow;
  const col = scopeCol;

  // Get Columns & Rows

  const titleColumn = scopeTitleColumn;
  const valueColumn = scopeValueColumn;
  const borderRow = scopeBorderRow;
  const backgroundStyleRow = scopeBackgroundStyleRow;

  const firstColumn = getA1FromCol(col)
  const tabsLastColumn = getColumnOffset(firstColumn, 3);

  // Get Ranges

  const textRanges1 = [
    valueColumn + borderRow, valueColumn + backgroundStyleRow, // Params Values
  ];

  const borderRanges1 = [
    titleColumn + borderRow + ':' + valueColumn + borderRow, // Border Param
    titleColumn + backgroundStyleRow + ':' + valueColumn + backgroundStyleRow // Style Param
  ];

  const textRanges2 = [];

  const borderRanges2 = [];

  // Set Highlight

  SetHighlight(sheet, textRanges1, borderRanges1, textRanges2, borderRanges2);
}

function SetScopeDarkMode() {
  const sheet = scopeSheet;
  const row = scopeRow;
  const col = scopeCol;
  const length = scopeLength;
  const width = scopeWidth;

  // Get Columns & Rows

  const firstColumn = getA1FromCol(col);
  const lastColumn = getA1FromCol(width);
  const lastRow = getRowFromName(scopeSheet, grandTotalName, scopeRow, scopeCol, scopeLength);
  const tabsRow = scopeTabsRow + 1;

  const tabsColumn = getA1FromCol(scopeCol + scopeVerticalTabsWidth);


  const titleColumn = scopeTitleColumn;
  const valueColumn = scopeValueColumn;
  const borderRow = scopeBorderRow;
  const backgroundStyleRow = scopeBackgroundStyleRow;

  // Get Ranges

  const backgroundRanges = ['A1:' + lastColumn + length]; // Whole Sheet

  const grayBackgroundRanges = [valueColumn + borderRow + ':' + valueColumn + backgroundStyleRow];

  const textRanges = [
    firstColumn + row + ':' + lastColumn + lastRow, // Whole Sheet
  ];

  const borderRanges = [];
  const backgroundBorderRanges = [];
  const grayBorderRanges = [];
  const rowsBorderRanges = [];
  const columnBorderRanges = [];

  // Set Dark Mode

  SetBackgroundDarkMode(sheet, backgroundRanges, grayBackgroundRanges, textRanges);
  SetBordersDarkMode(sheet, borderRanges, backgroundBorderRanges, grayBorderRanges, rowsBorderRanges, columnBorderRanges);
}

///// TIME MANAGEMENT /////______________________________________________________________________________________________________________________________________________________________________

function SetTimeManagementHighlights() {
  const sheet = timeManagementSheet;
  const length = timeManagementLength;
  const col = timeManagementCol;
  const width = sheet.getMaxColumns();

  // Get Columns & Rows
  const tabsRow = timeManagementTabsRow;

  // Main Bar
  const mainBarRow = timeManagementMainBarRow;
  const mainBarColum = timeManagementMainBarColumn;
  const mainBarLastColumn = getColumnOffset(mainBarColum, timeManagementMainBarWidth);

  // Progression Tabs
  const progressionColumn = getColumnFromName(sheet, importanceName, tabsRow, col, width);
  const progressionCol = getColumnFromA1(progressionColumn);

  const progressionToDoColumn = getColumnOffset(progressionColumn, 1);
  const progressionDoneColumn = getColumnOffset(progressionColumn, 3);
  const progressionTotalColumn = getColumnOffset(progressionColumn, 4);

  const importanceRow = tabsRow;
  const productionRow = getRowFromName(sheet, productionStageName, tabsRow, progressionCol, length);
  const teamRow = getRowFromName(sheet, backlogTeamName, tabsRow, progressionCol, length);

  const importanceLength = GetValues(GetImportanceRange()).length;
  const productionLength = GetValues(GetProductionRange()).length;
  const teamLength = GetValues(GetTeamRange()).length;

  // Job Tab
  const jobColumn = getColumnFromName(sheet, jobsName, tabsRow, col, width);
  const jobToDoColumn = getColumnOffset(jobColumn, 1);
  const jobDoneColumn = getColumnOffset(jobColumn, 3);
  const jobTotalColumn = getColumnOffset(jobColumn, 4);

  // Sprint Tab
  const sprintParamBorderTitleColumn = getColumnOffset(timemanagementSprintParamTitleColumn, -1);

  const sprintColumn = getColumnFromName(sheet, sprintName, tabsRow, col, width);
  const sprintDayColumn = getColumnOffset(sprintColumn, 2);

  // Epic Tab
  const epicColumn = getColumnFromName(sheet, epicName, tabsRow, col, width);
  const storiesColumn = getColumnOffset(epicColumn, 1);
  const storiesDayColumn = getColumnOffset(epicColumn, 2);
  const storiesTotalColumn = getColumnOffset(epicColumn, 3);

  // Get Ranges

  const textRanges1 = [
    getColumnOffset(mainBarColum, -3) + mainBarRow, // Main Bar Total
    getColumnOffset(mainBarColum, -1) + mainBarRow, // Main Bar Total Value


    // Progression Bar Title & To Do
    progressionColumn + importanceRow + ':' + progressionToDoColumn + importanceRow,
    progressionColumn + productionRow + ':' + progressionToDoColumn + productionRow,
    progressionColumn + teamRow + ':' + progressionToDoColumn + teamRow,
    jobColumn + tabsRow + ':' + jobToDoColumn + tabsRow,

    // Progression Bar Totals
    progressionTotalColumn + importanceRow,
    progressionTotalColumn + productionRow,
    progressionTotalColumn + teamRow,
    jobTotalColumn + tabsRow,

    // Progressions Bar Totals Values
    progressionToDoColumn + (importanceRow + 1) + ':' + progressionToDoColumn + (importanceRow + importanceLength),
    progressionToDoColumn + (productionRow + 1) + ':' + progressionToDoColumn + (productionRow + productionLength),
    progressionToDoColumn + (teamRow + 1) + ':' + progressionToDoColumn + (teamRow + teamLength),
    jobToDoColumn + (tabsRow + 1) + ':' + jobToDoColumn + (length - tabsRow),

    // Sprint Tab
    sprintColumn + tabsRow + ':' + sprintDayColumn + tabsRow, // Sprint Titles Row
    sprintColumn + tabsRow + ':' + sprintColumn + (length - tabsRow), // Sprint Names

    // Epic Tab
    epicColumn + tabsRow + ':' + epicColumn + (length - tabsRow), //Epic Title & Names
  ];

  const borderRanges1 = [

    // Main Bar
    mainBarColum + mainBarRow + ':' + mainBarLastColumn + mainBarRow, // Main Bar
    getColumnOffset(mainBarColum, -3) + mainBarRow + ':' + mainBarColum + mainBarRow, // Main Bar values

    // Sprint Parameter
    sprintParamBorderTitleColumn + timeManagementSprintParamIdealRow + ':' + timemanagementSprintParamTitleColumn + timeManagementSprintParamMaxRow, // Titles
    timeManagementSprintParamValueColumn + timeManagementSprintParamIdealRow + ':' + timeManagementSprintParamValueColumn + timeManagementSprintParamMaxRow, // Values

  ];

  const textRanges2 = [

    getColumnOffset(mainBarColum, -2) + mainBarRow, // Main Bar To Do

    // Progression Bar Done Values
    progressionDoneColumn + importanceRow + ':' + progressionDoneColumn + (importanceRow + importanceLength),
    progressionDoneColumn + productionRow + ':' + progressionDoneColumn + (productionRow + productionLength),
    progressionDoneColumn + teamRow + ':' + progressionDoneColumn + (teamRow + teamLength),
    jobDoneColumn + tabsRow + ':' + jobDoneColumn + (length - tabsRow),

    storiesColumn + tabsRow + ':' + storiesColumn + length, // Stories Title & Values
    storiesDayColumn + tabsRow + ':' + storiesTotalColumn + tabsRow, // Stories Days & Total Title
  ];

  const borderRanges2 = [];

  // Set Highlight

  SetHighlight(sheet, textRanges1, borderRanges1, textRanges2, borderRanges2);
}

function SetTimeManagementDarkMode() {
  const sheet = timeManagementSheet;
  const length = timeManagementLength;
  const col = timeManagementCol - 1;
  const firstColumn = getA1FromCol(col);
  const width = sheet.getLastColumn();
  const lastColumn = getA1FromCol(width);

  // Get Columns & Rows

  const tabsRow = timeManagementTabsRow;

  const progressionColumn = getColumnFromName(sheet, importanceName, tabsRow, col, width);
  const progressionCol = getColumnFromA1(progressionColumn);

  const progressionDoingColumn = getColumnOffset(progressionColumn, 2);
  const progressionDoneColumn = getColumnOffset(progressionColumn, 3);
  const progressionTotalColumn = getColumnOffset(progressionColumn, 4);
  const progressionBarColumn = getColumnOffset(progressionColumn, 5);

  const jobColumn = getColumnFromName(sheet, jobsName, tabsRow, col, width);

  const jobDoingColumn = getColumnOffset(jobColumn, 2);
  const jobDoneColumn = getColumnOffset(jobColumn, 3);
  const jobTotalColumn = getColumnOffset(jobColumn, 4);
  const jobBarColumn = getColumnOffset(jobColumn, 5);

  const sprintColumn = getColumnFromName(sheet, sprintName, tabsRow, col, width);

  const sprintTeamColumn = getColumnOffset(sprintColumn, 1);
  const sprintDaysColumn = getColumnOffset(sprintColumn, 2);
  const sprintBarColumn = getColumnOffset(sprintColumn, 3);

  const epicColumn = getColumnFromName(sheet, epicName, tabsRow, col, width);

  const storiesColumn = getColumnOffset(epicColumn, 1);
  const storiesDayColumn = getColumnOffset(epicColumn, 2);
  const storiesTotalColumn = getColumnOffset(epicColumn, 3);
  const storiesBarColumn = getColumnOffset(epicColumn, 4);

  // Sprint Parameters
  const firstParameterRow = timeManagementSprintParamIdealRow;
  const secondParameterRow = timeManagementSprintParamMaxRow;

  // Main Bar
  const mainBarRow = timeManagementMainBarRow;

  // Progression Tabs

  const importanceRow = tabsRow;
  const productionRow = getRowFromName(sheet, productionStageName, tabsRow, progressionCol, length);
  const teamRow = getRowFromName(sheet, backlogTeamName, tabsRow, progressionCol, length);


  // Lengths
  const importanceLength = GetValues(GetImportanceRange()).length;
  const productionLength = GetValues(GetProductionRange()).length;
  const teamLength = GetValues(GetTeamRange()).length;
  const jobsLength = GetValues(GetJobsRange()).length;

  const storiesRange = sheet.getRange(storiesColumn + tabsRow + ':' + storiesColumn + length)
  const sprintTabRange = sheet.getRange(sprintTeamColumn + tabsRow + ':' + sprintTeamColumn + length);

  const storiesLength = GetValues(storiesRange).length;
  const sprintTabLength = GetValues(sprintTabRange).length;

  // Get Sequences

  const sprintValues = sheet.getRange(sprintColumn + tabsRow + ':' + sprintColumn + length).getValues();
  const sprintSequence = GetSequenceWithEmpty(sprintValues);
  sprintSequence.shift();
  sprintSequence.pop();

  const storiesValues = sheet.getRange(epicColumn + tabsRow + ':' + epicColumn + length).getValues();
  const storiesSequence = GetSequenceWithEmpty(storiesValues);
  storiesSequence.shift();
  storiesSequence.pop();

  // Get Ranges

  const backgroundRanges = [
    firstColumn + 1 + ':' + getColumnOffset(lastColumn, 1) + 1, // First Row

    firstColumn + firstParameterRow + ':' + sprintTeamColumn + firstParameterRow, // Second Row
    getColumnOffset(sprintDaysColumn, 1) + firstParameterRow + ':' + getColumnOffset(lastColumn, 1) + firstParameterRow,

    firstColumn + secondParameterRow + ':' + sprintTeamColumn + secondParameterRow, // Third Row
    getColumnOffset(sprintDaysColumn, 1) + secondParameterRow + ':' + getColumnOffset(lastColumn, 1) + secondParameterRow,

    firstColumn + mainBarRow + ':' + getColumnOffset(lastColumn, 1) + length, // Rest of the sheet
  ];

  const grayBackgroundRanges = [

    // Tabs Titles
    progressionColumn + importanceRow + ':' + progressionTotalColumn + importanceRow,
    progressionColumn + productionRow + ':' + progressionTotalColumn + productionRow,
    progressionColumn + teamRow + ':' + progressionTotalColumn + teamRow,
    jobColumn + tabsRow + ':' + jobTotalColumn + tabsRow,
    sprintColumn + tabsRow + ':' + sprintDaysColumn + tabsRow,
    epicColumn + tabsRow + ':' + storiesTotalColumn + tabsRow
  ];

  const borderRanges = [

    // Importance 
    progressionColumn + importanceRow + ':' + progressionTotalColumn + importanceRow, // Title
    progressionColumn + (importanceRow + 1) + ':' + progressionBarColumn + (importanceRow + importanceLength), // Datas
    progressionBarColumn + (importanceRow + 1) + ':' + progressionBarColumn + (importanceRow + importanceLength), // Bars

    // Production Bar 
    progressionColumn + productionRow + ':' + progressionTotalColumn + productionRow, // Title
    progressionColumn + (productionRow + 1) + ':' + progressionBarColumn + (productionRow + productionLength), // Datas
    progressionBarColumn + (productionRow + 1) + ':' + progressionBarColumn + (productionRow + productionLength), // Bars

    // Team 
    progressionColumn + teamRow + ':' + progressionTotalColumn + teamRow, // Titles
    progressionColumn + (teamRow + 1) + ':' + progressionBarColumn + (teamRow + teamLength - 1), // Datas
    progressionBarColumn + (teamRow + 1) + ':' + progressionBarColumn + (teamRow + teamLength - 1), // Bars

    // Jobs 
    jobColumn + tabsRow + ':' + jobTotalColumn + tabsRow, // Title
    jobColumn + (tabsRow + 1) + ':' + jobBarColumn + (jobsLength + tabsRow - 2), // Datas
    jobBarColumn + (tabsRow + 1) + ':' + jobBarColumn + (jobsLength + tabsRow - 2), // Bars

    // Sprint
    sprintColumn + tabsRow + ':' + sprintDaysColumn + tabsRow, // Title
    sprintColumn + (tabsRow + 1) + ':' + sprintBarColumn + (sprintTabLength + tabsRow - 1), // Datas
    sprintBarColumn + (tabsRow + 1) + ':' + sprintBarColumn + (sprintTabLength + tabsRow - 1), // Bars

    // Epic
    epicColumn + tabsRow + ':' + storiesTotalColumn + tabsRow,// Title
    epicColumn + (tabsRow + 1) + ':' + storiesBarColumn + (storiesLength + tabsRow - 1), // Datas
    storiesBarColumn + (tabsRow + 1) + ':' + storiesBarColumn + (storiesLength + tabsRow - 1), // Bars
  ];

  const textRanges = [
    progressionDoingColumn + tabsRow + ':' + progressionDoneColumn + length, // Progression Bar Datas
    jobDoingColumn + tabsRow + ':' + jobDoneColumn + length, // Jobs Datas
    storiesDayColumn + tabsRow + ':' + storiesDayColumn + length // Epic Datas
  ];

  const backgroundBorderRanges = [];

  const grayBorderRanges = [];

  // Get Sprint Sequences Ranges
  let startRow = tabsRow + 1;

  for (let index in sprintSequence) {
    grayBorderRanges.push(sprintTeamColumn + startRow + ':' + sprintBarColumn + (startRow + sprintSequence[index] - 1));
    startRow = startRow + sprintSequence[index];
  };
  grayBorderRanges.push(sprintTeamColumn + (tabsRow + 1) + ':' + sprintBarColumn + (tabsRow + sprintTabLength - 1));

  // Get Epic Sequences Ranges
  startRow = tabsRow + 1;
  for (let index in storiesSequence) {
    grayBorderRanges.push(epicColumn + startRow + ':' + storiesBarColumn + (startRow + storiesSequence[index] - 1));
    startRow = startRow + storiesSequence[index];
  };
  grayBorderRanges.push(epicColumn + (tabsRow + 1) + ':' + storiesBarColumn + (tabsRow + storiesLength - 1));

  const rowsBorderRanges = [
    progressionColumn + (importanceRow + 1) + ':' + progressionBarColumn + (importanceRow + importanceLength), // Importance Bars
    progressionColumn + (productionRow + 1) + ':' + progressionBarColumn + (productionRow + productionLength), // Production Bars
    progressionColumn + (teamRow + 1) + ':' + progressionBarColumn + (teamRow + teamLength - 1), // Team Bars
    jobColumn + (tabsRow + 1) + ':' + jobBarColumn + (jobsLength + tabsRow - 2), // Jobs Bars
  ];

  const columnBorderRanges = []

  // Set Dark Mode

  SetBackgroundDarkMode(sheet, backgroundRanges, grayBackgroundRanges, textRanges);
  SetBordersDarkMode(sheet, borderRanges, backgroundBorderRanges, grayBorderRanges, rowsBorderRanges, columnBorderRanges);
}

///// BUDGET SETUP /////___________________________________________________________________________________________________________________________________________________________________________

function SetBudgetSetupHighlights() {
  const sheet = budgetSetupSheet;
  const row = budgetSetupRow;
  const col = budgetSetupCol;
  const length = budgetSetupLength;
  const width = budgetSetupWidth;
  const tabsWidth = budgetSetupTabsWidth;
  const firstColumn = getA1FromCol(col);
  const lastColumn = getA1FromCol(col + tabsWidth - 1);

  // Get Columns & Rows

  const parameterColumn = getColumnFromName(sheet, parameterName, row, col, width);
  const probabibilityColumn = getColumnFromName(sheet, probabilityName, row, col, width);
  const costsColumn = getColumnFromName(sheet, costSourcesName, row, col, width);
  const incomesColumn = getColumnFromName(sheet, incomeSourcesName, row, col, width);

  // Get Ranges

  const textRanges1 = []

  const borderRanges1 = [
    firstColumn + row + ':' + lastColumn + length, // Whole Tabs
    parameterColumn + row + ':' + parameterColumn + length, // Parameters

    // Titles
    parameterColumn + row,
    probabibilityColumn + row,
    costsColumn + row,
    incomesColumn + row,
  ]

  const textRanges2 = [
    firstColumn + row + ':' + lastColumn + row, //Titles
  ]

  const borderRanges2 = []

  // Set Highlight

  SetHighlight(sheet, textRanges1, borderRanges1, textRanges2, borderRanges2);
}

function SetBudgetSetupDarkMode() {
  const sheet = budgetSetupSheet;
  const row = budgetSetupRow;
  const col = budgetSetupCol;
  const length = budgetSetupLength;
  const width = budgetSetupWidth;
  const tabsWidth = budgetSetupTabsWidth;

  const firstColumn = getA1FromCol(col);
  const lastColumn = getA1FromCol(col + tabsWidth - 1);

  // Get Columns & Rows

  const parameterColumn = getColumnFromName(sheet, parameterName, row, col, width);
  const probabibilityColumn = getColumnFromName(sheet, probabilityName, row, col, width);
  const costsColumn = getColumnFromName(sheet, costSourcesName, row, col, width);
  const incomesColumn = getColumnFromName(sheet, incomeSourcesName, row, col, width);

  // Get Lengths
  const parameterRange = GetValues(sheet.getRange(parameterColumn + row + ':' + parameterColumn + length));
  const parameterLength = parameterRange.length;

  const probabilityRange = sheet.getRange(probabibilityColumn + row + ':' + probabibilityColumn + length)
  const probabilityLength = GetValues(probabilityRange).length;

  const costsRange = sheet.getRange(costsColumn + row + ':' + costsColumn + length)
  const costsLength = GetValues(costsRange).length;

  const incomesRange = sheet.getRange(incomesColumn + row + ':' + incomesColumn + length)
  const incomesLength = GetValues(incomesRange).length;

  // Get Ranges

  const backgroundRanges = [
    getColumnOffset(firstColumn, -1) + row + ':' + getColumnOffset(firstColumn, -1) + length, // Left Space
    getColumnOffset(lastColumn, 1) + row + ':' + getA1FromCol(col + budgetSetupWidth - 1) + length, // Right Space
    firstColumn + row + ':' + lastColumn + row, // Titles
  ];

  const grayBackgroundRanges = [
    // Bottom Spaces
    parameterColumn + (row + parameterLength) + ':' + parameterColumn + length,
    probabibilityColumn + (row + probabilityLength) + ':' + probabibilityColumn + length,
    costsColumn + (row + costsLength) + ':' + costsColumn + length,
    incomesColumn + (row + incomesLength) + ':' + incomesColumn + length,
    
  ];

  const textRanges = [];
  const borderRanges = [];
  const backgroundBorderRanges = [];
  const grayBorderRanges = [];
  const rowsBorderRanges = [];

  // Set Dark Mode

  SetBackgroundDarkMode(sheet, backgroundRanges, grayBackgroundRanges, textRanges);
  SetBordersDarkMode(sheet, borderRanges, backgroundBorderRanges, grayBorderRanges, rowsBorderRanges, []);
}

///// COSTS /////___________________________________________________________________________________________________________________________________________________________________________

function SetCostsHighlights() {
  const sheet = costsSheet;
  const row = costsTabsRow;
  const col = costsCol;
  const length = costsLength;
  const width = costsWidth;

  // Get Columns & Rows

  const titleRow = row - 1;

  const salaryColumn = getColumnFromName(sheet, salariesName, titleRow, col, width)
  const ponctualCostsColumn = getColumnFromName(sheet, ponctualCostsName, titleRow, col, width)
  const fixCostColumn = getColumnFromName(sheet, fixCostsName, titleRow, col, width)
  const variableCostColumn = getColumnFromName(sheet, variableCostsName, titleRow, col, width)

  const lastColumn = getA1FromCol(width)

  // Get Ranges

  const textRanges1 = [
    // Costs Titles
    salaryColumn + titleRow,
    ponctualCostsColumn + titleRow,
    fixCostColumn + titleRow,
    variableCostColumn + titleRow,

    // Costs Names
    salaryColumn + row + ':' + salaryColumn + length,
    ponctualCostsColumn + row + ':' + ponctualCostsColumn + length,
    fixCostColumn + row + ':' + fixCostColumn + length,
    variableCostColumn + row + ':' + variableCostColumn + (row + 4),

    // Salaries
    getColumnOffset(salaryColumn, 5) + row + ':' + getColumnOffset(salaryColumn, 6) + length,

  ]
  const borderRanges1 = []

  const textRanges2 = [

    // Costs Costs Title
    getColumnOffset(salaryColumn, 2) + titleRow + ':' + getColumnOffset(salaryColumn, 4) + titleRow,
    getColumnOffset(ponctualCostsColumn, 1) + titleRow,
    getColumnOffset(fixCostColumn, 1) + titleRow,

    // Costs Values
    getColumnOffset(salaryColumn, 2) + row + ':' + getColumnOffset(salaryColumn, 4) + length,
    getColumnOffset(ponctualCostsColumn, 1) + row + ':' + getColumnOffset(ponctualCostsColumn, 1) + length,
    getColumnOffset(fixCostColumn, 1) + row + ':' + getColumnOffset(fixCostColumn, 1) + length,
    variableCostColumn + (row + 5) + ':' + lastColumn + length,
    getColumnOffset(variableCostColumn, 1) + row,
  ]

  const borderRanges2 = []

  // Set Highlight

  SetHighlight(sheet, textRanges1, borderRanges1, textRanges2, borderRanges2);
}

function SetCostsDarkMode() {
  const sheet = costsSheet;
  const length = costsLength;
  const width = costsWidth;
  const tabsRow = costsTabsRow;
  const titleRow = tabsRow - 1;
  const col = costsCol;
  const dataRow = tabsRow + 1;

  // Get Columns & Rows

  const salaryColumn = getColumnFromName(sheet, salariesName, titleRow, col, width)
  const ponctualCostsColumn = getColumnFromName(sheet, ponctualCostsName, titleRow, col, width)
  const fixCostColumn = getColumnFromName(sheet, fixCostsName, titleRow, col, width)
  const variableCostColumn = getColumnFromName(sheet, variableCostsName, titleRow, col, width)

  const lastColumn = getA1FromCol(width)
  const maxColumn = getA1FromCol(sheet.getMaxColumns())

  const variableCostWidth = width - getColumnFromA1(variableCostColumn);

  // Get Ranges

  const backgroundRanges = ['A1:' + maxColumn + length]


  const grayBackgroundRanges = []
  const textRanges = []

  const borderRanges = [

    // Costs Tabs
    salaryColumn + tabsRow + ':' + getColumnOffset(salaryColumn, costsSalaryWidth - 1) + length,
    ponctualCostsColumn + tabsRow + ':' + getColumnOffset(ponctualCostsColumn, costsPonctualWidth - 1) + length,
    fixCostColumn + tabsRow + ':' + getColumnOffset(fixCostColumn, fixCostsWidth - 1) + length,
    variableCostColumn + tabsRow + ':' + getColumnOffset(variableCostColumn, variableCostWidth) + length,

    // Costs Tabs Titles
    salaryColumn + tabsRow + ':' + getColumnOffset(salaryColumn, costsSalaryWidth - 1) + tabsRow,
    ponctualCostsColumn + tabsRow + ':' + getColumnOffset(ponctualCostsColumn, costsPonctualWidth - 1) + tabsRow,
    fixCostColumn + tabsRow + ':' + getColumnOffset(fixCostColumn, fixCostsWidth - 1) + tabsRow,
    variableCostColumn + tabsRow + ':' + getColumnOffset(variableCostColumn, variableCostWidth) + tabsRow,

    // Costs Names
    salaryColumn + tabsRow + ':' + salaryColumn + length,
    ponctualCostsColumn + tabsRow + ':' + ponctualCostsColumn + length,
    fixCostColumn + tabsRow + ':' + fixCostColumn + length,
    variableCostColumn + tabsRow + ':' + variableCostColumn + length,
    variableCostColumn + tabsRow + ':' + lastColumn + (tabsRow + 3),
  ]

  const backgroundBorderRanges = []

  const grayBorderRanges = []

  const rowsBorderRanges = []

  const columnBorderRanges = [
    variableCostColumn + dataRow + ':' + getColumnOffset(variableCostColumn, variableCostWidth) + length, // Variable Costs Datas
  ]

  const alternatingColorsRanges = [
    salaryColumn + tabsRow + ':' + getColumnOffset(salaryColumn, costsSalaryWidth - 1) + length, // Salary Tabs
    ponctualCostsColumn + tabsRow + ':' + getColumnOffset(ponctualCostsColumn, costsPonctualWidth - 1) + length, // Ponctual Tabs
    fixCostColumn + tabsRow + ':' + getColumnOffset(fixCostColumn, fixCostsWidth - 1) + length, // Fix Tabs
    variableCostColumn + tabsRow + ':' + getColumnOffset(variableCostColumn, variableCostWidth) + length, // Variable Costs
  ]

  // Merge Varibal Costs Title
  sheet.getRange(variableCostColumn + titleRow + ':' + maxColumn + titleRow).breakApart();
  sheet.getRange(variableCostColumn + titleRow + ':' + lastColumn + titleRow).merge();

  SetBackgroundDarkMode(sheet, backgroundRanges, grayBackgroundRanges, textRanges);
  SetDarkModeBandings(sheet, alternatingColorsRanges)
  SetBordersDarkMode(sheet, borderRanges, backgroundBorderRanges, grayBorderRanges, rowsBorderRanges, columnBorderRanges);
}

///// INCOMES /////___________________________________________________________________________________________________________________________________________________________________________

function SetIncomesHighlights() {
  const sheet = incomesSheet;
  const length = incomesLength;
  const width = incomesWidth;
  const row = incomesTabsRow;
  const titleRow = row - 1;
  const col = incomesCol;

  const ponctualIncomesColumn = getColumnFromName(sheet, ponctualIncomesName, titleRow, col, width)
  const fixIncomeColumn = getColumnFromName(sheet, fixIncomesName, titleRow, col, width)
  const variableIncomeColumn = getColumnFromName(sheet, variableIncomesName, titleRow, col, width)

  const lastColumn = getA1FromCol(width)

  const textRanges1 = [
    // Incomes Titles
    ponctualIncomesColumn + titleRow,
    fixIncomeColumn + titleRow,
    variableIncomeColumn + titleRow,

    // Incomes Names
    ponctualIncomesColumn + row + ':' + ponctualIncomesColumn + length,
    fixIncomeColumn + row + ':' + fixIncomeColumn + length,
    variableIncomeColumn + row + ':' + variableIncomeColumn + (row + 4),

  ]
  const borderRanges1 = []

  const textRanges2 = [
    // Incomes Incomes Title
    getColumnOffset(ponctualIncomesColumn, 1) + titleRow,
    getColumnOffset(fixIncomeColumn, 1) + titleRow,

    // Incomes Values
    getColumnOffset(ponctualIncomesColumn, 1) + row + ':' + getColumnOffset(ponctualIncomesColumn, 1) + length,
    getColumnOffset(fixIncomeColumn, 1) + row + ':' + getColumnOffset(fixIncomeColumn, 1) + length,
    variableIncomeColumn + (row + 5) + ':' + lastColumn + length,
    getColumnOffset(variableIncomeColumn, 1) + row,
  ]

  const borderRanges2 = []

  // Set Highlight

  SetHighlight(sheet, textRanges1, borderRanges1, textRanges2, borderRanges2);
}

function SetincomesDarkMode() {
  const sheet = incomesSheet;
  const length = incomesLength;
  const width = incomesWidth;
  const tabsRow = incomesTabsRow;
  const titleRow = tabsRow - 1;
  const col = incomesCol;
  const dataRow = tabsRow + 1;

  // Get Columns & Rows

  const ponctualIncomesColumn = getColumnFromName(sheet, ponctualIncomesName, titleRow, col, width)
  const fixIncomeColumn = getColumnFromName(sheet, fixIncomesName, titleRow, col, width)
  const variableIncomeColumn = getColumnFromName(sheet, variableIncomesName, titleRow, col, width)

  const lastColumn = getA1FromCol(width)
  const maxColumn = getA1FromCol(sheet.getMaxColumns())

  const variableIncomeWidth = width - getColumnFromA1(variableIncomeColumn);


  const backgroundRanges = ['A1:' + maxColumn + length]

  // Get Ranges

  const grayBackgroundRanges = []
  const textRanges = []

  const borderRanges = [

    // Incomes Tabs
    ponctualIncomesColumn + tabsRow + ':' + getColumnOffset(ponctualIncomesColumn, incomesPonctualWidth - 1) + length,
    fixIncomeColumn + tabsRow + ':' + getColumnOffset(fixIncomeColumn, fixIncomesWidth - 1) + length,
    variableIncomeColumn + tabsRow + ':' + getColumnOffset(variableIncomeColumn, variableIncomeWidth) + length,

    // Incomes Tabs Title
    ponctualIncomesColumn + tabsRow + ':' + getColumnOffset(ponctualIncomesColumn, incomesPonctualWidth - 1) + tabsRow,
    fixIncomeColumn + tabsRow + ':' + getColumnOffset(fixIncomeColumn, fixIncomesWidth - 1) + tabsRow,
    variableIncomeColumn + tabsRow + ':' + getColumnOffset(variableIncomeColumn, variableIncomeWidth) + tabsRow,

    // Incomes Names
    ponctualIncomesColumn + tabsRow + ':' + ponctualIncomesColumn + length,
    fixIncomeColumn + tabsRow + ':' + fixIncomeColumn + length,
    variableIncomeColumn + tabsRow + ':' + variableIncomeColumn + length,
    variableIncomeColumn + tabsRow + ':' + lastColumn + (tabsRow + 3),
  ]

  const backgroundBorderRanges = []

  const grayBorderRanges = []

  const rowsBorderRanges = []

  const columnBorderRanges = [
    variableIncomeColumn + dataRow + ':' + getColumnOffset(variableIncomeColumn, variableIncomeWidth) + length, // Variable Incomes Datas
  ]

  const alternatingColorsRanges = [
    ponctualIncomesColumn + tabsRow + ':' + getColumnOffset(ponctualIncomesColumn, incomesPonctualWidth - 1) + length, // Ponctual Incomes Datas
    fixIncomeColumn + tabsRow + ':' + getColumnOffset(fixIncomeColumn, fixIncomesWidth - 1) + length, // Fix Incomes Datas
    variableIncomeColumn + tabsRow + ':' + getColumnOffset(variableIncomeColumn, variableIncomeWidth) + length, // Variable Incomes Datas
  ]

  sheet.getRange(variableIncomeColumn + titleRow + ':' + maxColumn + titleRow).breakApart();
  sheet.getRange(variableIncomeColumn + titleRow + ':' + lastColumn + titleRow).merge();

  // Set Dark Mode

  SetBackgroundDarkMode(sheet, backgroundRanges, grayBackgroundRanges, textRanges);
  SetDarkModeBandings(sheet, alternatingColorsRanges)
  SetBordersDarkMode(sheet, borderRanges, backgroundBorderRanges, grayBorderRanges, rowsBorderRanges, columnBorderRanges);
}

///// BUDGET PLANNING /////_________________________________________________________________________________________________________________________________________________________________

function SetbudgetTimelineHighlights() {
  const sheet = budgetTimelineSheet;
  const row = budgetTimelineRow;
  const length = budgetTimelineLength;
  const width = budgetTimelineMinWidth
  const lastColumn = getA1FromCol(width);
  const firstColumn = getColumnFromName(sheet, monthsName, row, 1, width)

  // Get Columns & Rows

  const monthRow = row;
  const dataRow = row + 1;

  const secondColumn = getColumnOffset(firstColumn, 1);
  const seconCol = getColumnFromA1(secondColumn);

  const startingDateRow = getRowFromName(sheet, budgetTimelineStartingDateName, row, getColumnFromA1(budgetTimelineTitleColumn), length);
  const nbMonthsRow = getRowFromName(sheet, budgetTimelineDurationName, row, getColumnFromA1(budgetTimelineTitleColumn), length);

  const initialFundRow = getRowFromName(sheet, budgetTimelineInitialFundName, row, getColumnFromA1(budgetTimelineTitleColumn), length);
  const probabilityRow = getRowFromName(sheet, probabilityName, row, getColumnFromA1(budgetTimelineTitleColumn), length);
  const gradientRow = getRowFromName(sheet, budgetTimelineFundGradientName, row, getColumnFromA1(budgetTimelineTitleColumn), length);
  const scopeFundRow = getRowFromName(sheet, 'Importance Scope', row, getColumnFromA1(budgetTimelineTitleColumn), length);

  const titleColumn = budgetTimelineTitleColumn;
  const paramColumn = budgetTimelineParam1Column;
  const valueColumn = getColumnOffset(budgetTimelineParam1Column, 1)

  // Get Lentghs

  const probabilityLength = GetValues(GetProbabilityRange()).length - 1;
  const importanceLength = GetValues(GetImportanceRange()).length + 1;
  const fundGradientLength = budgetTimelineGradientLength;

  // Get Setup Data
  const importances = cleanData(sheet.getRange(budgetTimelineParam1Column+(scopeFundRow+1) + ':' + budgetTimelineParam1Column+sheet.getLastRow()).getValues().filter(String));
  const nbImportances = importances.length;
  const setupImportances = GetValues(GetImportanceRange());
  const setupImportancesColors = GetColors(GetImportanceRange());

  // Get Ranges

  const textRanges1 = [
    // Param Titles
    titleColumn + initialFundRow,
    titleColumn + startingDateRow,
    titleColumn + nbMonthsRow,
    titleColumn + probabilityRow,
    titleColumn + gradientRow,
    titleColumn + scopeFundRow + ':' + budgetTimelineParam3Column + scopeFundRow,
  ]

  const borderRanges1 = [
    // Main Parameters
    titleColumn + initialFundRow + ':' + valueColumn + initialFundRow, // Initial Fund
    titleColumn + startingDateRow + ':' + valueColumn + startingDateRow, // Starting Date
    titleColumn + nbMonthsRow + ':' + valueColumn + nbMonthsRow, // Nb Months

    // probability
    titleColumn + probabilityRow + ':' + titleColumn + (probabilityRow + probabilityLength - 1), // Title
    paramColumn + probabilityRow + ':' + valueColumn + (probabilityRow + probabilityLength - 1), // Values

    // gradient
    titleColumn + gradientRow + ':' + paramColumn + (gradientRow + fundGradientLength - 1), // Title
    valueColumn + gradientRow + ':' + valueColumn + (gradientRow + fundGradientLength - 1), // Values

    // scope
    titleColumn + scopeFundRow + ':' + titleColumn + (scopeFundRow + importanceLength - 1), // Title
    paramColumn + scopeFundRow + ':' + getColumnOffset(valueColumn, 1) + (scopeFundRow + importanceLength - 1), // Values
    titleColumn + scopeFundRow + ':' + budgetTimelineParam3Column + scopeFundRow, // Title
  ]

  const textRanges2 = [
    // Param Values
    paramColumn + initialFundRow + ':' + valueColumn + initialFundRow, // Initial Funds
    paramColumn + startingDateRow + ':' + valueColumn + startingDateRow, // Starting Date
    paramColumn + nbMonthsRow + ':' + valueColumn + nbMonthsRow, // Nb Months

    paramColumn + probabilityRow + ':' + valueColumn + (probabilityRow + probabilityLength - 1), // Probability
    paramColumn + (scopeFundRow + 1) + ':' + getColumnOffset(valueColumn, 1) + (scopeFundRow + importanceLength + 1), // Scope
  ]

  const borderRanges2 = [];

  // Set Highlight
  SetHighlight(sheet, textRanges1, borderRanges1, textRanges2, borderRanges2);

  // Scope Style

  // Get Ranges
  const importanceRanges = [];

  for(let k = 0 ; k < importances.length ; k++){
    importanceRanges.push(budgetTimelineParam1Column+(scopeFundRow+k+1) +':'+ budgetTimelineParam3Column+(scopeFundRow+k+1))
  }

  // Set Scope Colors
  for(let index in importanceRanges){
    const importanceRange = sheet.getRange(importanceRanges[index])
    const importanceValue = importanceRange.getValues()[0][0];
    const importanceColor = setupImportancesColors[setupImportances.indexOf(importanceValue)];
    importanceRange.setFontColor(importanceColor);
  }
  
}

function SetbudgetTimelineDarkMode() {
  const sheet = budgetTimelineSheet;
  const row = budgetTimelineRow;
  const dataRow = row + 1;
  const monthRow = row;
  const length = budgetTimelineLength;
  const width = budgetTimelineMinWidth

  // Get Param Datas

  const nbMonthsRow = getRowFromName(sheet, budgetTimelineDurationName, row, getColumnFromA1(budgetTimelineTitleColumn), length);
  const nbMonths = sheet.getRange(budgetTimelineParam1Column + nbMonthsRow).getValue();

  // Get Columns & Rows

  const firstColumn = getColumnFromName(sheet, monthsName, monthRow, 1, width)
  const firstCol = getColumnFromA1(firstColumn);

  const lastCol = firstCol + nbMonths;
  const lastColumn = getA1FromCol(lastCol);

  const maxRow = sheet.getMaxRows()
  const maxColumn = getA1FromCol(sheet.getMaxColumns())

  const secondColumn = getColumnOffset(firstColumn, 1);

  const gradientRow = getRowFromName(sheet, budgetTimelineFundGradientName, row, getColumnFromA1(budgetTimelineTitleColumn), length)


  const fundRow = getRowFromName(sheet, fundsName, row, firstCol, length)
  const costsRow = getRowFromName(sheet, costsName, row, firstCol, length)
  const incomesRow = getRowFromName(sheet, incomesName, row, firstCol, length)

  const gainsRow = getRowFromName(sheet, gainsName, budgetTimelineRow, firstCol, length);
  const sparklineFundStartRow = fundRow - budgetTimelineBarLength;
  const sparklineCurveStartRow = gainsRow - budgetTimelineCurveLength;

  const fundGradientLength = budgetTimelineGradientLength;
  const costsLength = GetValues(GetCostRange()).length + 1;
  const incomesLength = GetValues(GetIncomeRange()).length + 1;

  const paramColumn = budgetTimelineParam1Column;
  const valueColumn = getColumnOffset(budgetTimelineParam1Column, 1)

  // Get Ranges

  const backgroundRanges = [
    'A1:' + paramColumn + maxRow,
    valueColumn + '1:' + valueColumn + (gradientRow - 1),
    valueColumn + (gradientRow + fundGradientLength) + ':' + valueColumn + maxRow,
    getColumnOffset(valueColumn, 1) + '1:' + maxColumn + maxRow
  ];

  const grayBackgroundRanges = []

  const textRanges = [
    secondColumn + dataRow + ':' + maxColumn + maxRow,
  ];

  const borderRanges = [
    firstColumn + monthRow + ':' + lastColumn + monthRow, // Months

    // Costs
    firstColumn + costsRow + ':' + lastColumn + costsRow, // Costs Row
    firstColumn + costsRow + ':' + lastColumn + (costsRow + costsLength - 1), // Sources

    // Incomes
    firstColumn + incomesRow + ':' + lastColumn + incomesRow, // Incomes Row
    firstColumn + incomesRow + ':' + lastColumn + (incomesRow + incomesLength - 1),// Sources

    // Sparklines
    firstColumn + sparklineFundStartRow + ':' + lastColumn + sparklineFundStartRow, // Fund Sparkline
    firstColumn + sparklineCurveStartRow + ':' + lastColumn + sparklineCurveStartRow, // Gains Curve
    firstColumn + gainsRow + ':' + lastColumn + gainsRow, // Gains
    firstColumn + fundRow + ':' + lastColumn + fundRow, // Funds
  ];

  const backgroundBorderRanges = [
  ];

  const grayBorderRanges = [];

  const rowsBorderRanges = [];

  const columnBorderRanges = [
    secondColumn + dataRow + ':' + lastColumn + (fundRow - (budgetTimelineBarLength + budgetTimelineCurveLength) - 2) // Datas
  ];

  // Set Dark Mode
  SetBackgroundDarkMode(sheet, backgroundRanges, grayBackgroundRanges, textRanges);
  SetBordersDarkMode(sheet, borderRanges, backgroundBorderRanges, grayBorderRanges, rowsBorderRanges, columnBorderRanges);
}

///// ANNUAL BUDGET /////_______________________________________________________________________________________________________________________________________________________________________

function SetAnnualBudgetHighlights() {
  const sheet = annualBudgetSheet;
  const row = annualBudgetRow;
  const col = annualBudgetFirstCol;
  const length = annualBudgetLength;

  // Get Columns & Rows

  const paramTitleColumn = annualBudgetParamTitleColumn;
  const param1Column = annualBudgetParam1Column;
  const param2Column = annualBudgetParam2Column;
  const parameterCol = getColumnFromA1(paramTitleColumn)

  const startingYearRow = getRowFromName(sheet, startingDateName, row, parameterCol, length)
  const endingYearRow = startingYearRow + 1
  const probaRow = getRowFromName(sheet, probabilityName, row, parameterCol, length)

  const probaLentgh = GetValues(GetProbabilityRange()).length - 1;

  // Get Ranges

  const textRanges1 = [
    paramTitleColumn + startingYearRow,
    paramTitleColumn + endingYearRow,
    paramTitleColumn + probaRow,
  ];

  const borderRanges1 = [
    paramTitleColumn + startingYearRow + ':' + param1Column + endingYearRow,
    paramTitleColumn + probaRow + ':' + param2Column + (probaRow + probaLentgh - 1),
  ];

  const textRanges2 = [
  ];

  const borderRanges2 = [
  ];

  // Set Highlight

  SetHighlight(sheet, textRanges1, borderRanges1, textRanges2, borderRanges2);
}

function SetAnnualBudgetDarkMode() {
  const sheet = annualBudgetSheet;

  // Get Columns & Rows

  // Get Ranges

  const backgroundRanges = [
    'A1:'+getA1FromCol(sheet.getMaxColumns())+sheet.getMaxRows()
  ];
  const grayBackgroundRanges = [];
  const textRanges = [];
  const borderRanges = [];
  const backgroundBorderRanges = [];
  const grayBorderRanges = [];
  const rowsBorderRanges = [];
  const columnBorderRanges = [];

  // Set Dark Mode

  SetBackgroundDarkMode(sheet, backgroundRanges, grayBackgroundRanges, textRanges);
  SetBordersDarkMode(sheet, borderRanges, backgroundBorderRanges, grayBorderRanges, rowsBorderRanges, columnBorderRanges);
}

///// PRODUCTION BUDGET /////___________________________________________________________________________

function SetProductionBudgetHighlights() {
  const sheet = productionBudgetSheet
  const row = productionBudgetRow;
  const length = productionBudgetLength;

  // Get Columns & Rows

  const probabilityRow = getRowFromName(sheet, probabilityName, row, getColumnFromA1(productionBudgetParamTitleColumn), length);
  const scopeFundRow = getRowFromName(sheet, 'Production Scope', row, getColumnFromA1(productionBudgetParamTitleColumn), length);

  const nbProba = GetValues(GetProbabilityRange()).length - 1;

  const stages = cleanData(sheet.getRange(productionBudgetParam1Column+(scopeFundRow+1) + ':' + productionBudgetParam1Column+sheet.getLastRow()).getValues().filter(String));
  const nbStage = stages.length;

  const setupStages = GetValues(GetProductionRange());
  const setupStagesColors = GetColors(GetProductionRange());

  const datesFirstRow = scopeFundRow+nbStage+2;

  // Get Columns & Rows

  // Get Ranges

  const textRanges1 = [
    productionBudgetParamTitleColumn+probabilityRow, // Probability Title
    productionBudgetParamTitleColumn+scopeFundRow, // Scope Title
  ];

  const borderRanges1 = [
    productionBudgetParamTitleColumn+probabilityRow+':'+productionBudgetParam2Column+(probabilityRow+nbProba-1), // Proba Border
    productionBudgetParamTitleColumn+scopeFundRow+':'+productionBudgetParam3Column+(scopeFundRow+nbStage), // Scope Border
    productionBudgetParamTitleColumn+scopeFundRow+':'+productionBudgetParamTitleColumn+(scopeFundRow+nbStage), // Scope Title
    productionBudgetParamTitleColumn+scopeFundRow+':'+productionBudgetParam3Column+scopeFundRow, // Scope Param
  ];

  const textRanges2 = [
    productionBudgetParam1Column+probabilityRow+':'+productionBudgetParam2Column+(probabilityRow+nbProba-1), // Proba Param
    productionBudgetParam1Column+scopeFundRow+':'+productionBudgetParam3Column+scopeFundRow, // Scope Param
  ];

  const stageRanges = [];
  const dateRanges = [];
  const dateTitleCells = [];

  for(let k = 0 ; k < stages.length ; k++){
    stageRanges.push(productionBudgetParam1Column+(scopeFundRow+k+1) +':'+ productionBudgetParam3Column+(scopeFundRow+k+1))
    const dateRow = datesFirstRow + k*4;

    dateRanges.push(productionBudgetParamTitleColumn+dateRow+':'+productionBudgetParam3Column+(dateRow+2));
    dateTitleCells.push(productionBudgetParamTitleColumn+dateRow);
    textRanges1.push(productionBudgetParamTitleColumn+(dateRow+1)+':'+productionBudgetParamTitleColumn+(dateRow+2));
    textRanges2.push(productionBudgetParam2Column+(dateRow+1)+':'+productionBudgetParam2Column+(dateRow+2));
  }

  for(let index in dateRanges){
    const stageRange = sheet.getRange(stageRanges[index])
    const dateRange = sheet.getRange(dateRanges[index]);
    const dateTitleCell = sheet.getRange(dateTitleCells[index]);
    const stageValue = dateTitleCell.getValues()[0][0];
    const stageColor = setupStagesColors[setupStages.indexOf(stageValue)];

    dateRange.setBorder(true,true,true,true,null,null,stageColor,SpreadsheetApp.BorderStyle.SOLID_THICK);
    dateTitleCell.setBorder(true,true,true,true,null,null,stageColor,SpreadsheetApp.BorderStyle.SOLID_THICK);
    dateTitleCell.setFontColor(stageColor);
    stageRange.setFontColor(stageColor);
  }

  const borderRanges2 = [];

  // Set Highlight

  SetHighlight(sheet, textRanges1, borderRanges1, textRanges2, borderRanges2);
}

function SetProductionBudgetDarkMode() {
  const sheet = productionBudgetSheet;

  // Get Columns & Rows

  // Get Ranges

  const backgroundRanges = [
    'A1:'+getA1FromCol(sheet.getMaxColumns())+sheet.getMaxRows()
  ];
  const grayBackgroundRanges = [];
  const textRanges = [];
  const borderRanges = [];
  const backgroundBorderRanges = [];
  const grayBorderRanges = [];
  const rowsBorderRanges = [];
  const columnBorderRanges = [];

  // Set Dark Mode

  SetBackgroundDarkMode(sheet, backgroundRanges, grayBackgroundRanges, textRanges);
  SetBordersDarkMode(sheet, borderRanges, backgroundBorderRanges, grayBorderRanges, rowsBorderRanges, columnBorderRanges);
}

///// CUSTOM BUDGET /////___________________________________________________________________________

function SetCustomBudgetHighlights(){
  const sheet = customBudgetSheet;
  const row = customBudgeRow;
  const col = customBudgetFirstCol;
  const length = customBudgetLength;

  // Get Columns & Rows

  const paramTitleColumn = customBudgetParamTitleColumn;
  const param1Column = customBudgetParam1Column;
  const param2Column = customBudgetParam2Column;
  const parameterCol = getColumnFromA1(paramTitleColumn)

  const startingYearRow = getRowFromName(sheet, startingDateName, row, parameterCol, length)
  const endingYearRow = startingYearRow + 1
  const probaRow = getRowFromName(sheet, probabilityName, row, parameterCol, length)

  const probaLentgh = GetValues(GetProbabilityRange()).length - 1;

  // Get Ranges

  const textRanges1 = [
    paramTitleColumn + startingYearRow,
    paramTitleColumn + endingYearRow,
    paramTitleColumn + probaRow,
  ];

  const borderRanges1 = [
    paramTitleColumn + startingYearRow + ':' + param1Column + endingYearRow,
    paramTitleColumn + probaRow + ':' + param2Column + (probaRow + probaLentgh - 1),
  ];

  const textRanges2 = [
  ];

  const borderRanges2 = [
  ];

  // Set Highlight

  SetHighlight(sheet, textRanges1, borderRanges1, textRanges2, borderRanges2);
}

function SetCustomBudgetDarkMode() {
  const sheet = customBudgetSheet;

  // Get Columns & Rows

  // Get Ranges

  const backgroundRanges = [
    'A1:'+getA1FromCol(sheet.getMaxColumns())+sheet.getMaxRows()
  ];
  const grayBackgroundRanges = [];
  const textRanges = [];
  const borderRanges = [];
  const backgroundBorderRanges = [];
  const grayBorderRanges = [];
  const rowsBorderRanges = [];
  const columnBorderRanges = [];

  // Set Dark Mode

  SetBackgroundDarkMode(sheet, backgroundRanges, grayBackgroundRanges, textRanges);
  SetBordersDarkMode(sheet, borderRanges, backgroundBorderRanges, grayBorderRanges, rowsBorderRanges, columnBorderRanges);
}

///// EXECUTION /////___________________________________________________________________________

function RefreshHighlights() {

  // Highlights
  SetSetupHighlight();
  SetBacklogHighlight();
  SetSprintHighlight();
  SetPlanningHighlight();
  SetScopeHighlights();
  SetScopeBorders()
  SetTimeManagementHighlights();
  SetBudgetSetupHighlights();
  SetCostsHighlights();
  SetIncomesHighlights();
  SetbudgetTimelineHighlights();
  SetAnnualBudgetHighlights();
  SetProductionBudgetHighlights();
  SetCustomBudgetHighlights();

  // Backlog
  TraceBacklogBorders();

  // Time Management
  SetTimeManagementBars();

  // Budget Timeline
  CreatebudgetTimeline();
}

function SetDarkModes() {

  // Message
  SpreadsheetApp.getActive().toast('Changing Dark Mode', '', 40);

  // Change Theme
  const theme = ss.getPredefinedSpreadsheetThemes()[0];
  const backgroundColor = SpreadsheetApp.newColor().setRgbColor(GetDarkModeColor()).build();
  const textColor = SpreadsheetApp.newColor().setRgbColor(GetDarkModeTextColor()).build();

  theme
    .setConcreteColor(SpreadsheetApp.ThemeColorType.TEXT, textColor)
    .setConcreteColor(SpreadsheetApp.ThemeColorType.BACKGROUND, backgroundColor)
    .setFontFamily(sheetFont);

  ss.setSpreadsheetTheme(theme);

  // Dark Modes
  SetSetupDarkMode();
  SetBacklogDarkMode();
  SetSprintDarkMode();
  SetPlanningDarkMode();
  SetScopeDarkMode();
  SetTimeManagementDarkMode();
  SetBudgetSetupDarkMode();
  SetCostsDarkMode();
  SetincomesDarkMode();
  SetbudgetTimelineDarkMode();
  SetAnnualBudgetDarkMode();
  SetProductionBudgetDarkMode();
  SetCustomBudgetDarkMode();

  // Highlights
  RefreshHighlights();

  // Showdowns
  RefreshShowdowns();
}

function SwitchDarkMode() {
  range = setupSheet.getRange(GetDarkModeCellA1());
  isDarkMode = range.getValue();
  if (isDarkMode == true) {
    range.setValue(false);
    SetDarkModes(false);
  } else {
    range.setValue(true);
    SetDarkModes(true);
  };
}

function RefreshDarkMode() {
  SetDarkModes();
}

////////////////////////////
//         SHOWDOWNS      //_______________________________________________________________________________________________________________________________________
////////////////////////////

///// FUNCTION /////___________________________________________________________________________

function SetShowdownsValidationRule(setupRange, targetRange) {
  const values = GetValues(setupRange);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values)
    .build();
  targetRange.setDataValidation(rule);
}

function SetShowdowns(setupRanges, targetRanges) {
  rules = [];
  for (let index in setupRanges) {
    rules.concat(GetConditionnalFormattingRulesExact(setupRanges[index], targetRanges[index]));
    SetShowdownsValidationRule(setupRange[index], targetRange[index]);
  };
  targetRanges[0].getSheet().setConditionalFormatRules(rules);
  return;
}

function SetShowdownRanges(setupRange, targetRanges) {
  const values = GetValues(setupRange);
  for (let range of targetRanges) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(values)
      .build()
    range.setDataValidation(rule);
  };
}

///// BACKLOG /////___________________________________________________________________________

function SetBacklogShowdowns() {
  const sheet = backlogSheet;
  const row = backlogTabsRow;
  const col = backlogTabsCol;
  const width = backlogTabsWidth;
  const length = backlogLength;

  const dataRow = row + 1;

  // Clear Data Validation
  sheet.getRange(row + col + ':' + width + length).clearDataValidations();

  // Get Columns

  const statusColumn = getColumnFromName(sheet, satusName, row, col, width);
  const priorityColumn = getColumnFromName(sheet, priorityName, row, col, width);
  const importanceColumn = getColumnFromName(sheet, importanceName, row, col, width);
  const jobColumn = getColumnFromName(sheet, jobsName, row, col, width);
  const teamColumn = getColumnFromName(sheet, backlogTeamName, row, col, width);
  const productionColumn = getColumnFromName(sheet, productionStageName, row, col, width);

  // Get Ranges

  const backlogStatusRange = sheet.getRange(statusColumn + dataRow + ':' + statusColumn + backlogLength);
  const backlogPriorityRange = sheet.getRange(priorityColumn + dataRow + ':' + priorityColumn + backlogLength);
  const backlogImportanceRange = sheet.getRange(importanceColumn + dataRow + ':' + importanceColumn + backlogLength);
  const backlogJobRange = sheet.getRange(jobColumn + dataRow + ':' + jobColumn + backlogLength);
  const backlogTeamRange = sheet.getRange(teamColumn + dataRow + ':' + teamColumn + backlogLength);
  const backlogProductionRange = sheet.getRange(productionColumn + dataRow + ':' + productionColumn + backlogLength);

  let setupRanges = [GetTeamRange(), GetJobsRange(), GetProductionRange(), GetStatusRange(), GetImportanceRange()];
  let targetRanges = [backlogTeamRange, backlogJobRange, backlogProductionRange, backlogStatusRange, backlogImportanceRange];

  // Validation & Conditionnal Rules Team, Job, Production, Status

  let rules = [];
  for (let index in setupRanges) {
    SetShowdownsValidationRule(setupRanges[index], targetRanges[index]);
    const showdownsRules = GetConditionnalFormattingRulesExact(setupRanges[index], targetRanges[index]);
    rules = rules.concat(showdownsRules);
  };

  // Priotity Rules

  const priorityFormat = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue(priorityMaxColor, SpreadsheetApp.InterpolationType.NUMBER, priotityMax)
    .setGradientMidpointWithValue(priorityMidColor, SpreadsheetApp.InterpolationType.NUMBER, priotityMid)
    .setGradientMinpointWithValue(priorityMinColor, SpreadsheetApp.InterpolationType.NUMBER, priorityMin)
    .setRanges([backlogPriorityRange])
    .build();

  rules.push(priorityFormat);


  // Set Conditionnal Rules

  sheet.setConditionalFormatRules(rules);
}

///// SPRINT /////___________________________________________________________________________

function SetSprintShowdown() {
  const setupSheet = backlogSheet;
  const targetSheet = sprintSheet;

  const backlogSprintColumn = getColumnFromName(setupSheet, sprintName, backlogTabsRow, backlogTabsCol, backlogTabsWidth);
  const backlogSprintRange = setupSheet.getRange(backlogSprintColumn + (backlogTabsRow + 1) + ':' + backlogSprintColumn + backlogLength);

  const sprintParameterRange = targetSheet.getRange(sprintTitleColumn + sprintValueRow);

  SetShowdownsValidationRule(backlogSprintRange, sprintParameterRange);
}

function SetSprintConditionnalFomatRules() {
  const sheet = sprintSheet;

  // Get Columns

  const priorityColumn = getColumnFromName(sheet, priorityName, sprintTabsRow, sprintCol, sprintTabsWidth);
  const statusColumn = getColumnFromName(sheet, satusName, sprintTabsRow, sprintCol, sprintTabsWidth);
  const teamColumn = getColumnFromName(sheet, teamName, sprintTabsRow, sprintCol, sprintTabsWidth);
  const jobColumn = getColumnFromName(sheet, jobsName, sprintTabsRow, sprintCol, sprintTabsWidth);
  const dayColumn = getColumnFromName(sheet, nbDayEst30Name, sprintTabsRow, sprintCol, sprintTabsWidth);
  const startingDayColumn = getColumnFromName(sheet, planningStartingDaysName, sprintTabsRow, sprintCol, sheet.getLastColumn());

  const dataRow = sprintTabsRow + 1;
  const dataColumn = getColumnOffset(startingDayColumn, 1);
  const lastColumn = getA1FromCol(sheet.getLastColumn());

  const dateCell = dataColumn + '$' + sprintTabsRow;
  const taskDateCell = '$' + startingDayColumn + dataRow
  const durationCell = '$' + dayColumn + dataRow
  const teamCell = '$' + teamColumn + dataRow

  // Get Ranges

  const priorityRanges = sheet.getRange(priorityColumn + dataRow + ':' + priorityColumn + sprintLength);
  const statusRange = sheet.getRange(statusColumn + dataRow + ':' + statusColumn + sprintLength);
  const teamRange = sheet.getRange(teamColumn + dataRow + ':' + teamColumn + sprintLength);
  const jobRange = sheet.getRange(jobColumn + dataRow + ':' + jobColumn + sprintLength);
  const daysRange = sheet.getRange(dayColumn + dataRow + ':' + dayColumn + sprintLength);
  const ganttRange = sheet.getRange(dataColumn + dataRow + ':' + lastColumn + sprintLength);

  const dayGradientStartCell = sheet.getRange(sprintDayGradientColumn + sprintDayGradientStart);
  const dayGradientEndCell = sheet.getRange(sprintDayGradientColumn + sprintDayGradientEnd);

  // Get Setup Datas

  const setupTeamRange = GetTeamRange();
  const setupTeamColors = setupTeamRange.getBackgrounds();
  const setupTeamMembers = GetValues(setupTeamRange);

  // Get Conditionnal Format Rules 

  let rules = [];

  // Priority Formats
  const priorityFormat = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue(priorityMaxColor, SpreadsheetApp.InterpolationType.NUMBER, priotityMax)
    .setGradientMidpointWithValue(priorityMidColor, SpreadsheetApp.InterpolationType.NUMBER, priotityMid)
    .setGradientMinpointWithValue(priorityMinColor, SpreadsheetApp.InterpolationType.NUMBER, priorityMin)
    .setRanges([priorityRanges])
    .build();

  // Team & Job Formats
  const statusRules = GetConditionnalFormattingRulesExact(GetStatusRange(), statusRange);
  const teamRules = GetConditionnalFormattingRulesExact(GetTeamRange(), teamRange);
  const jobRules = GetConditionnalFormattingRulesExact(GetJobsRange(), jobRange);

  // Days Format
  const daysFormat = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue(dayGradientEndCell.getBackground(), SpreadsheetApp.InterpolationType.NUMBER, dayGradientEndCell.getValue())
    .setGradientMinpointWithValue(dayGradientStartCell.getBackground(), SpreadsheetApp.InterpolationType.NUMBER, dayGradientStartCell.getValue())
    .setRanges([daysRange])
    .build();

  // Gantt
  const ganttRules = [];

  for (let index in setupTeamMembers) {

    const ganttFormula = '=IF(AND(AND((' + dateCell + '>=' + taskDateCell + ');(' + dateCell + '<=' + taskDateCell + '+(ROUND(' + durationCell + ')+((WORKDAY(' + taskDateCell + ';' + durationCell + ')-' + taskDateCell + ')-' + durationCell + '))));' + teamCell + '=\"' + setupTeamMembers[index] + '\");True;False)';

    const ganttFormat = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(ganttFormula)
      .setBackground(setupTeamColors[index])
      .setRanges([ganttRange])
      .build();

    ganttRules.push(ganttFormat);
  }

  // Concat Rules
  rules.push(priorityFormat);
  rules = rules.concat(statusRules);
  rules = rules.concat(teamRules);
  rules = rules.concat(jobRules);
  rules = rules.concat(ganttRules);
  rules.push(daysFormat);


  //Set the rules
  sheet.setConditionalFormatRules(rules);
}

///// PLANNING /////___________________________________________________________________________

function SetPlanningShowdown(ranges) {
  const setupSheet = backlogSheet;

  const sprintColumn = getColumnFromName(setupSheet, sprintName, backlogTabsRow, backlogTabsCol, backlogTabsWidth);
  const backlogDataRow = backlogTabsRow + 1;
  const sprintRange = setupSheet.getRange(sprintColumn + backlogDataRow + ':' + sprintColumn + backlogLength);

  for (let range of ranges) range.setFontColor(getHighlightcolor1());

  SetShowdownRanges(sprintRange, ranges);
}

///// SCOPE /////___________________________________________________________________________

function GetScopeImportanceFormat(importance, color, format, targetRange) {
  const sheet = scopeSheet;

  // Get Rows & Columns

  const refColumn = getA1FromCol(scopeCol + scopeVerticalTabsWidth);
  const refRow = scopeTabsRow + 1;

  const statusRow = '$' + parseInt(refRow - 1);
  const importanceRow = '$' + parseInt(refRow - 2);

  const productionColumn = getColumnFromName(sheet, refRow, scopeTabsRow, scopeCol, scopeVerticalTabsWidth);
  const teamColumn = getColumnFromName(sheet, backlogTeamName, scopeTabsRow, scopeCol, scopeVerticalTabsWidth);

  // Get Formula

  const doneValue = GetStatusRange().getValues()[2][0];

  const importanceFormula = '=AND(OR(AND(' + getColumnOffset(refColumn, -1) + statusRow + '=\"' + doneValue + '\";OR(OR(;' + getColumnOffset(refColumn, -3) + importanceRow + '=\"' + importance + '\");OR(' + getColumnOffset(refColumn, -2) + importanceRow + '=\"' + importance + '\";' + getColumnOffset(refColumn, -1) + importanceRow + '=\"' + importance + '\")));AND(IFERROR(IF(SEARCH(\"Total\"; ' + refColumn + importanceRow + '); True); False);IFERROR(IF(SEARCH(\"' + importance + '\"; ' + refColumn + importanceRow + '); True); False)));NOT(AND(' + '$' + productionColumn + refRow + ' = 0;' + '$' + teamColumn + refRow + ' =0)))';

  // Make the format Rule

  const newImportanceFormat = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(importanceFormula)
    .setBackground(color)
    .setFontColor(format.getForegroundColor())
    .setRanges([targetRange])
    .build();

  return newImportanceFormat;
}

// PROTOTYPE AI GENERATED NOT TESTED // =============================================================================================================
function generateFormulas(refColumn, statusRow, importanceRow, status, importance, productionColumn, teamColumn, refRow, length) {
  const formulas = [];
  let offset = 0;

  for (let i = 0; i < length; i++) {
    const importanceColumn = getColumnOffset(refColumn, offset);

    let formula = '=AND(AND(' + refColumn + statusRow + '=\"' + status + '\";OR(';

    for (let j = 0; j <= i; j++) {
      const importanceColumnOffset = getColumnOffset(refColumn, offset - j);
      formula += importanceColumnOffset + importanceRow + '=\"' + importance + '\"';
      if (j < i) {
        formula += ';';
      }
    }

    formula += '));NOT(AND(' + productionColumn + refRow + ' = 0;' + teamColumn + refRow + ' =0)))';

    formulas.push(formula);
    offset++;
  }

  return formulas;
}
//===================================================================================================================================================

function GetScopeStatusFormat(status, importance, color, targetRange) {
  const sheet = scopeSheet;

  // Get Columns & Rows

  const refColumn = getA1FromCol(scopeCol + scopeVerticalTabsWidth);
  const refRow = scopeTabsRow + 1;

  const statusRow = '$' + parseInt(refRow - 1);
  const importanceRow = '$' + parseInt(refRow - 2);

  const productionColumn = '$' + getColumnFromName(sheet, refRow, scopeTabsRow, scopeCol, scopeVerticalTabsWidth);
  const teamColumn = '$' + getColumnFromName(sheet, backlogTeamName, scopeTabsRow, scopeCol, scopeVerticalTabsWidth);

  // Get Setup Datas

  const statusList = GetValues(GetStatusRange());

  const statusFormulas = [
    // To Do Formula
    '=AND(AND(' + refColumn + statusRow + '=\"' + status + '\";' + refColumn + importanceRow + '=\"' + importance + '\");NOT(AND(' + productionColumn + refRow + ' = 0;' + teamColumn + refRow + ' =0)))',

    // Doing Formula
    '=AND(AND(' + refColumn + statusRow + '=\"' + status + '\";OR(' + refColumn + importanceRow + '=\"' + importance + '\";' + getColumnOffset(refColumn, -1) + importanceRow + '=\"' + importance + '\"));NOT(AND(' + productionColumn + refRow + ' = 0;' + teamColumn + refRow + ' =0)))',

    // Done Formula
    '=AND(AND(' + refColumn + statusRow + '=\"' + status + '\";OR(OR(' + refColumn + importanceRow + '=\"' + importance + '\";' + getColumnOffset(refColumn, -1) + importanceRow + '=\"' + importance + '\");' + getColumnOffset(refColumn, -2) + importanceRow + '=\"' + importance + '\"));NOT(AND(' + productionColumn + refRow + ' = 0;' + teamColumn + refRow + ' =0)))',
  ];

  for (let index in statusList) {

    // Check the status
    if (status == statusList[index]) {

      const newStatusFormat = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(statusFormulas[index])
        .setBackground(color)
        .setFontColor(GetDarkModeTextColor())
        .setRanges([targetRange])
        .build();

      return newStatusFormat;

    };
  };
}

function SetScopeConditionnalFormatingRules() {
  const sheet = scopeSheet;

  // Get Columns & Rows

  const lastColumn = getA1FromCol(scopeWidth);
  const lastRow = sheet.getRange(getColumnFromName(sheet, productionStageName, scopeTabsRow, scopeCol, scopeVerticalTabsWidth) + '1:' + getColumnFromName(sheet, productionStageName, scopeTabsRow, scopeCol, scopeVerticalTabsWidth) + scopeLength).getValues().filter(String).length + 10;

  const tabsRefRow = '$' + (scopeTabsRow + 1);
  const statusRow = '$' + parseInt(scopeTabsRow);
  const importanceRow = '$' + parseInt(scopeTabsRow - 1);
  const tabsRow = scopeTabsRow + 1;

  const productionColumn = '$' + getColumnFromName(sheet, productionStageName, scopeTabsRow, scopeCol, scopeVerticalTabsWidth);
  const teamColumn = '$' + getColumnFromName(sheet, backlogTeamName, scopeTabsRow, scopeCol, scopeVerticalTabsWidth);
  const jobColumn = '$' + getColumnFromName(sheet, jobsName, scopeTabsRow, scopeCol, scopeVerticalTabsWidth);
  const tabsColumn = getA1FromCol(scopeCol + scopeVerticalTabsWidth);

  // Get Setup Datas

  const setupStatusRange = GetStatusRange();
  const setupImportanceRange = GetImportanceRange();
  const setupTeamRange = GetTeamRange();
  const setupProductionRange = GetProductionRange();

  const status = GetValues(setupStatusRange);
  const importance = GetValues(setupImportanceRange);
  const teamValues = GetValues(setupTeamRange);
  const productionValues = GetValues(setupProductionRange);

  const importanceColors = GetColors(setupImportanceRange);
  const teamColors = GetColors(setupTeamRange);
  const productionColors = GetColors(setupProductionRange);

  const importanceTextStyle = GetTextStyles(setupImportanceRange);
  const teamTextStyle = GetTextStyles(setupTeamRange);
  const productionTextStyles = GetTextStyles(setupProductionRange);


  // Get Ranges

  const productionColumnRange = sheet.getRange(productionColumn + tabsRow + ':' + productionColumn + lastRow);
  const teamColumnRange = sheet.getRange(teamColumn + tabsRow + ':' + teamColumn + lastRow);
  const jobColumnRange = sheet.getRange(jobColumn + tabsRow + ':' + jobColumn + lastRow);

  const titleRange = sheet.getRange(tabsColumn + importanceRow + ':' + lastColumn + lastRow);
  const tabsRange = sheet.getRange(tabsColumn + tabsRow + ':' + lastColumn + lastRow);

  // Prepare Rules

  let rules = [];
  let totalRules = [];
  let dataRules = [];
  let titlesRules = [];
  let dataAlternativeRule = [];
  let teamRules = [];
  let productionRules = [];
  let jobRules = [];

  // Get Rules

  // Basic Rules
  teamRules = teamRules.concat(GetConditionnalFormattingRulesContain(setupTeamRange, teamColumnRange));
  productionRules = productionRules.concat(GetConditionnalFormattingRulesContain(GetProductionRange(), productionColumnRange));
  jobRules = GetConditionnalFormattingRulesContain(GetJobsRange(), jobColumnRange);

  // Data
  for (let i in importance) {
    const importanceFormat = importanceTextStyle[i];
    const newImportanceFormat = GetScopeImportanceFormat(importance[i], importanceColors[i], importanceFormat, tabsRange)
    dataRules.push(newImportanceFormat);
    let coeff = scopeColorFadingCoeff;
    for (let j in status) {
      const newStatusFormat = GetScopeStatusFormat(status[j], importance[i], lightenDarkenColor(importanceColors[i], 1 - coeff), tabsRange);
      coeff = coeff * coeff;
      dataRules.push(newStatusFormat);
    }
  }

  // Title
  let count = 0;
  for (let index in importanceColors) {

    tabsFormula = '=AND(AND(AND(NOT(AND(' + tabsColumn + importanceRow + '=0;' + tabsColumn + statusRow + '=0));OR(OR(' + tabsColumn + importanceRow + ' = "' + importance[count] + '";' + getColumnOffset(tabsColumn, -1) + importanceRow + ' = "' + importance[count] + '");OR(' + getColumnOffset(tabsColumn, -2) + importanceRow + ' = "' + importance[count] + '";' + getColumnOffset(tabsColumn, -3) + importanceRow + ' = "' + importance[count] + '")));NOT(AND(IFERROR(IF(SEARCH(\"Total\"; ' + tabsColumn + importanceRow + '); True); False);NOT(OR(IFERROR(IF(SEARCH(\"Must have\"; ' + tabsColumn + importanceRow + '); True); False);OR(IFERROR(IF(SEARCH(\"Nice to have\"; ' + tabsColumn + importanceRow + '); True); False);IFERROR(IF(SEARCH(\"Withlist\"; ' + tabsColumn + importanceRow + '); True); False)))))));AND(NOT(AND(' + productionColumn + parseInt(scopeTabsRow - 1) + '=0;' + teamColumn + parseInt(scopeTabsRow - 1) + '=0));NOT(AND(' + tabsColumn + importanceRow + '=0;' + tabsColumn + statusRow + '=0))))';

    const newTitleFormat = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(tabsFormula)
      .setBackground(importanceColors[index])
      .setFontColor(importanceTextStyle[index].getForegroundColor())
      .setBold(true)
      .setRanges([titleRange])
      .build();
    count++;
    titlesRules.push(newTitleFormat);
  }

  // Team  
  for (let index in teamValues) {
    // Get Formattings
    const teamValue = teamValues[index];
    const teamColor = teamColors[index];
    const teamtextStyle = teamTextStyle[index];

    // Get Formulas
    const teamBlanckFormula = '=AND(IFERROR(IF(SEARCH("' + teamValue + '"; ' + teamColumn + tabsRow + '); True); False);' + jobColumn + tabsRow + '=0)'

    const teamTotalFormula = '=AND(AND(IFERROR(IF(SEARCH("Total"; ' + teamColumn + tabsRow + '); True); False);IFERROR(IF(SEARCH("' + teamValue + '"; ' + teamColumn + tabsRow + '); True); False));NOT(AND(' + tabsColumn + importanceRow + '=0;' + tabsColumn + statusRow + '=0)))';

    //Fill the blanck in the job column
    let newTeamFormat = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(teamBlanckFormula)
      .setBackground(teamColor)
      .setBold(teamtextStyle.isBold())
      .setFontColor(teamtextStyle.getForegroundColor())
      .setRanges([jobColumnRange])
      .build()
    teamRules.push(newTeamFormat);


    //Team Total Row
    newTeamFormat = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(teamTotalFormula)
      .setBackground(teamColor)
      .setBold(teamtextStyle.isBold())
      .setFontColor(teamtextStyle.getForegroundColor())
      .setRanges([tabsRange])
      .build()
    teamRules.push(newTeamFormat);
  }

  //Production

  for (let k = 0; k < productionValues.length; k++) {
    // Get Formattings
    const productionValue = productionValues[k];
    const productionColor = productionColors[k];
    const productionTextStyle = productionTextStyles[k]

    // Get Formulas
    const productionColumnBlanckFormula = '=AND(' + productionColumn + tabsRow + '=0;OR(IFERROR(IF(SEARCH("' + productionValue + '"; ' + productionColumn + (tabsRow + 1) + '); True); False);IFERROR(IF(SEARCH("' + productionValue + '"; ' + productionColumn + (tabsRow - 1) + '); True); False)))'

    const productionTeamBlanckFormula = '=AND(' + teamColumn + tabsRow + '=0;IFERROR(IF(SEARCH("' + productionValue + '"; ' + productionColumn + tabsRow + '); True); False))'

    const productionJobBlanckFormula = '=AND(' + jobColumn + tabsRow + '=0;IFERROR(IF(SEARCH("' + productionValue + '"; ' + productionColumn + tabsRow + '); True); False))'

    const productionTotalFormula = '=AND(AND(IFERROR(IF(SEARCH("Total"; ' + productionColumn + tabsRow + '); True); False);IFERROR(IF(SEARCH("' + productionValue + '"; ' + productionColumn + tabsRow + '); True); False));NOT(AND(' + tabsColumn + statusRow + '=0;' + tabsColumn + importanceRow + '=0)))'

    //Blanck in the production Column
    newTeamFormat = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(productionColumnBlanckFormula)
      .setBackground(productionColor)
      .setBold(productionTextStyle.isBold())
      .setFontColor(productionTextStyle.getForegroundColor())
      .setRanges([productionColumnRange])
      .build()
    teamRules.push(newTeamFormat)

    //Blanck in the team Column
    newProductionFormat = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(productionTeamBlanckFormula)
      .setBackground(productionColor)
      .setBold(productionTextStyle.isBold())
      .setFontColor(productionTextStyle.getForegroundColor())
      .setRanges([teamColumnRange])
      .build()
    productionRules.push(newProductionFormat)

    //Blanck in the job Column
    newProductionFormat = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(productionJobBlanckFormula)
      .setBackground(productionColor)
      .setBold(productionTextStyle.isBold())
      .setFontColor(productionTextStyle.getForegroundColor())
      .setRanges([jobColumnRange])
      .build()
    productionRules.push(newProductionFormat)

    //Production totals
    newProductionFormat = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(productionTotalFormula)
      .setBackground(productionColor)
      .setBold(productionTextStyle.isBold())
      .setFontColor(productionTextStyle.getForegroundColor())
      .setRanges([tabsRange])
      .build()
    totalRules.push(newProductionFormat)
  }

  // Background Alternative Style
  if (sheet.getRange(scopeValueColumn + scopeBackgroundStyleRow).getValue()) {

    const perFormula = '=AND(AND(MOD(ROW(' + tabsColumn + tabsRow + ');2)=0;OR(AND(' + tabsColumn + tabsRow + '=0;NOT(OR(OR(IFERROR(IF(SEARCH("Total"; ' + tabsColumn + importanceRow + '); True); False) ; IFERROR(IF(SEARCH("Total"; ' + teamColumn + tabsRow + '); True); False));IFERROR(IF(SEARCH("Total"; ' + tabsColumn + tabsRefRow + '); True); False) )));AND(IFERROR(IF(SEARCH("Total"; ' + tabsColumn + importanceRow + '); True); False);NOT(OR(IFERROR(IF(SEARCH("Must Have"; ' + tabsColumn + importanceRow + '); True); False);OR(IFERROR(IF(SEARCH("Nice To Have"; ' + tabsColumn + importanceRow + '); True); False);IFERROR(IF(SEARCH("Withlist"; ' + tabsColumn + importanceRow + '); True); False)))))));AND(NOT(AND(' + productionColumn + tabsRow + '=0;' + teamColumn + tabsRow + '=0));NOT(AND(' + tabsColumn + importanceRow + '=0;' + tabsColumn + statusRow + '=0))))'

    const oddFormula = '=AND(AND(MOD(ROW(' + tabsColumn + tabsRow + ');2)=1;OR(AND(' + tabsColumn + tabsRow + '=0;NOT(OR(OR(IFERROR(IF(SEARCH("Total"; ' + tabsColumn + importanceRow + '); True); False) ; IFERROR(IF(SEARCH("Total"; ' + teamColumn + tabsRow + '); True); False));IFERROR(IF(SEARCH("Total"; ' + tabsColumn + tabsRefRow + '); True); False) )));AND(IFERROR(IF(SEARCH("Total"; ' + tabsColumn + importanceRow + '); True); False);NOT(OR(IFERROR(IF(SEARCH("Must Have"; ' + tabsColumn + importanceRow + '); True); False);OR(IFERROR(IF(SEARCH("Nice To Have"; ' + tabsColumn + importanceRow + '); True); False);IFERROR(IF(SEARCH("Withlist"; ' + tabsColumn + importanceRow + '); True); False)))))));AND(NOT(AND(' + productionColumn + tabsRow + '=0;' + teamColumn + tabsRow + '=0));NOT(AND(' + tabsColumn + importanceRow + '=0;' + tabsColumn + statusRow + '=0))))'

    const perRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(perFormula)
      .setRanges([tabsRange])
      .build()
    dataAlternativeRule.push(perRule);

    const oddRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(oddFormula)
      .setBackground(GetDarkModeGrayColor())
      .setRanges([tabsRange])
      .build()
    dataAlternativeRule.push(oddRule);
  }

  // Make the rules

  rules = [
    ...jobRules,
    ...productionRules,
    ...teamRules,
    ...totalRules,
    ...dataAlternativeRule,
    ...dataRules,
    ...titlesRules,
  ]

  ///// Set scope Conditionnal Format Rules_____________________________________________________________________________________

  sheet.setConditionalFormatRules(rules);
}

///// TIME MANAGEMENT /////___________________________________________________________________________

function SetTimeManagementConditionnalFormattingRules() {
  const sheet = timeManagementSheet;

  // Get Columns & Rows 

  const length = sheet.getLastRow();
  const tabsRow = timeManagementTabsRow;
  const col = timeManagementCol;
  const width = sheet.getLastColumn();
  const dataRow = tabsRow + 1;

  // Epic Tab
  const epicColumn = getColumnFromName(sheet, epicName, tabsRow, col, width);
  const storiesColumn = getColumnOffset(epicColumn, 1);
  const storiesTotalColumn = getColumnOffset(epicColumn, 3);

  // Sprint Tab
  const sprintColumn = getColumnFromName(sheet, sprintName, tabsRow, col, width);
  const sprintDaysColumn = getColumnOffset(sprintColumn, 2);
  const storiesRange = sheet.getRange(storiesColumn + tabsRow + ':' + storiesColumn + length);

  // Get Length
  const storiesLength = GetValues(storiesRange).length;

  // Get Ranges
  const sprintDayRange = sheet.getRange(sprintDaysColumn + tabsRow + ':' + sprintDaysColumn + (dataRow + sprintLength - 1))
  const epicDayRange = sheet.getRange(storiesTotalColumn + tabsRow + ':' + storiesTotalColumn + (dataRow + storiesLength - 1))

  const idealSprintDurationRange = sheet.getRange(timeManagementSprintParamValueColumn + timeManagementSprintParamIdealRow);
  const maxSprintDurationRange = sheet.getRange(timeManagementSprintParamValueColumn + timeManagementSprintParamMaxRow);

  // Get Sheet Data 
  const idealSprintDurationValue = idealSprintDurationRange.getValue();
  const maxSprintDurationValue = maxSprintDurationRange.getValue();

  const idealSprintDurationColor = idealSprintDurationRange.getBackground();
  const maxSprintDurationColor = maxSprintDurationRange.getBackground();

  // Get Conditionnal Format Rules 

  const rules = [];

  const epicDayValues = cleanData(epicDayRange.getValues()).filter(num => {
    if (typeof num === "number" && parseInt(num) !== num) {
      return true;
    } return false;
  })
  const epicMaxDayValue = Math.ceil(Math.max(...epicDayValues))

  const sprintFormat = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue(maxSprintDurationColor, SpreadsheetApp.InterpolationType.NUMBER, maxSprintDurationValue)
    .setGradientMidpointWithValue(lightenDarkenColor(idealSprintDurationColor, 0.5), SpreadsheetApp.InterpolationType.NUMBER, idealSprintDurationValue)
    .setGradientMinpointWithValue(GetDarkModeColor(), SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setRanges([sprintDayRange])
    .build();

  const epicFormat = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue(maxSprintDurationColor, SpreadsheetApp.InterpolationType.NUMBER, epicMaxDayValue)
    .setGradientMinpointWithValue(lightenDarkenColor(idealSprintDurationColor, 0.5), SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setRanges([epicDayRange])
    .build();

  // Make Rules

  rules.push(sprintFormat);
  rules.push(epicFormat);

  // Set Rules

  sheet.setConditionalFormatRules(rules);
}

///// COSTS /////___________________________________________________________________________

function SetCostShowdowns() {
  const sheet = costsSheet;
  const length = costsLength;
  const width = costsWidth;
  const row = costsTabsRow;
  const col = costsCol;

  sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).clearDataValidations()

  // Get Column & Rows

  const titleRow = row - 1;
  const dataRow = row + 1;

  const salaryColumn = getColumnFromName(sheet, salariesName, titleRow, col, width);
  const ponctualCostsColumn = getColumnFromName(sheet, ponctualCostsName, titleRow, col, width);
  const fixCostColumn = getColumnFromName(sheet, fixCostsName, titleRow, col, width);
  const variableCostColumn = getColumnFromName(sheet, variableCostsName, titleRow, col, width);
  const variableCostCol = getColumnFromA1(variableCostColumn);

  const salaryContractColumn = getColumnFromName(sheet, permanentContractName, row, getColumnFromA1(salaryColumn), costsSalaryWidth);

  const ponctualProbabilityColumn = getColumnFromName(sheet, probabilityName, row, getColumnFromA1(ponctualCostsColumn), costsPonctualWidth);
  const ponctualSourceColumn = getColumnFromName(sheet, sourcesName, row, getColumnFromA1(ponctualCostsColumn), costsPonctualWidth);

  const fixProbabilityColumn = getColumnFromName(sheet, probabilityName, row, getColumnFromA1(fixCostColumn), fixCostsWidth);
  const fixSourceColumn = getColumnFromName(sheet, sourcesName, row, getColumnFromA1(fixCostColumn), fixCostsWidth);

  const variableProbabilityRow = getRowFromName(sheet, probabilityName, row, variableCostCol, length);
  const variableSourceRow = getRowFromName(sheet, sourcesName, row, variableCostCol, length);

  const lastColumn = getA1FromCol(sheet.getLastColumn());

  // Get Length

  const teamValues = GetTeamRange().getValues();
  const teamLength = teamValues.length;

  // Get Ranges

  const salaryContractRange = sheet.getRange(salaryContractColumn + dataRow + ':' + salaryContractColumn + length);
  const costTeamRange = sheet.getRange(salaryColumn + dataRow + ':' + salaryColumn + (dataRow + teamLength - 1));

  const ponctualProbabilityRange = sheet.getRange(ponctualProbabilityColumn + dataRow + ':' + ponctualProbabilityColumn + length);
  const ponctualSourceRange = sheet.getRange(ponctualSourceColumn + dataRow + ':' + ponctualSourceColumn + length);

  const fixProbabilityRange = sheet.getRange(fixProbabilityColumn + dataRow + ':' + fixProbabilityColumn + length);
  const fixSourceRange = sheet.getRange(fixSourceColumn + dataRow + ':' + fixSourceColumn + length);

  const variableMonthRange = sheet.getRange(getColumnOffset(variableCostColumn, 1) + (dataRow + 3) + ':' + fixProbabilityColumn + length);
  const variableProbabilityRange = sheet.getRange(getColumnOffset(variableCostColumn, 1) + variableProbabilityRow + ':' + lastColumn + variableProbabilityRow);
  const variableSourceRange = sheet.getRange(getColumnOffset(variableCostColumn, 1) + variableSourceRow + ':' + lastColumn + variableSourceRow);


  // Set Showdowns

  SetShowdownRanges(GetProbabilityRange(), [ponctualProbabilityRange, fixProbabilityRange, variableProbabilityRange]);
  SetShowdownRanges(GetCostRange(), [ponctualSourceRange, fixSourceRange, variableSourceRange]);

  // Get Conditionnal Format Rules

  // Set Salary Conditionnal Formatting
  costTeamRange.setValues(teamValues);
  const teamRules = GetConditionnalFormattingRulesExact(GetTeamRange(), costTeamRange);
  const contractRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();
  salaryContractRange.setDataValidation(contractRule);

  // Ponctual Showdowns
  const ponctualProbabilityRules = GetConditionnalFormattingRulesExact(GetProbabilityRange(), ponctualProbabilityRange);
  const ponctualSourceRules = GetConditionnalFormattingRulesExact(GetCostRange(), ponctualSourceRange);

  // Fix Showdowns
  const fixProbabilityRules = GetConditionnalFormattingRulesExact(GetProbabilityRange(), fixProbabilityRange);
  const fixSourceRules = GetConditionnalFormattingRulesExact(GetCostRange(), fixSourceRange);

  // Variable ShowDown
  const monthRules = GetConditionnalFormattingRulesContain(GetMonthsRange(), variableMonthRange);
  const variableProbabilityRules = GetConditionnalFormattingRulesExact(GetProbabilityRange(), variableProbabilityRange);
  const variableSourceRules = GetConditionnalFormattingRulesExact(GetCostRange(), variableSourceRange);

  //Make Rules

  const rules = [
    ...teamRules,
    ...ponctualProbabilityRules,
    ...ponctualSourceRules,
    ...fixProbabilityRules,
    ...fixSourceRules,
    ...monthRules,
    ...variableProbabilityRules,
    ...variableSourceRules
  ]

  // Set Conditionnal Format Rules

  sheet.setConditionalFormatRules(rules);
}

///// INCOMES /////___________________________________________________________________________

function SetIncomeShowdowns() {
  const sheet = incomesSheet;
  const length = incomesLength;
  const width = incomesWidth;
  const row = incomesTabsRow;
  const col = incomesCol;

  // Get Columns & Rows

  const titleRow = row - 1;
  const dataRow = row + 1;

  const ponctualIncomesColumn = getColumnFromName(sheet, ponctualIncomesName, titleRow, col, width);
  const fixIncomeColumn = getColumnFromName(sheet, fixIncomesName, titleRow, col, width);
  const variableIncomeColumn = getColumnFromName(sheet, variableIncomesName, titleRow, col, width);
  const variableIncomeCol = getColumnFromA1(variableIncomeColumn);

  const ponctualProbabilityColumn = getColumnFromName(sheet, probabilityName, row, getColumnFromA1(ponctualIncomesColumn), incomesPonctualWidth);
  const ponctualSourceColumn = getColumnFromName(sheet, sourcesName, row, getColumnFromA1(ponctualIncomesColumn), incomesPonctualWidth);

  const fixProbabilityColumn = getColumnFromName(sheet, probabilityName, row, getColumnFromA1(fixIncomeColumn), fixIncomesWidth);
  const fixSourceColumn = getColumnFromName(sheet, sourcesName, row, getColumnFromA1(fixIncomeColumn), fixIncomesWidth);

  const variableProbabilityRow = getRowFromName(sheet, probabilityName, row, variableIncomeCol, length);
  const variableSourceRow = getRowFromName(sheet, sourcesName, row, variableIncomeCol, length);

  const lastColumn = getA1FromCol(width);

  // Get Ranges

  const ponctualProbabilityRange = sheet.getRange(ponctualProbabilityColumn + dataRow + ':' + ponctualProbabilityColumn + length);
  const ponctualSourceRange = sheet.getRange(ponctualSourceColumn + dataRow + ':' + ponctualSourceColumn + length);

  const fixProbabilityRange = sheet.getRange(fixProbabilityColumn + dataRow + ':' + fixProbabilityColumn + length);
  const fixSourceRange = sheet.getRange(fixSourceColumn + dataRow + ':' + fixSourceColumn + length);

  const variableMonthRange = sheet.getRange(getColumnOffset(variableIncomeColumn, 1) + (dataRow + 3) + ':' + fixProbabilityColumn + length);
  const variableProbabilityRange = sheet.getRange(getColumnOffset(variableIncomeColumn, 1) + variableProbabilityRow + ':' + lastColumn + variableProbabilityRow);
  const variableSourceRange = sheet.getRange(getColumnOffset(variableIncomeColumn, 1) + variableSourceRow + ':' + lastColumn + variableSourceRow);

  // Set Showdowns

  SetShowdownRanges(GetProbabilityRange(), [ponctualProbabilityRange, fixProbabilityRange, variableProbabilityRange]);
  SetShowdownRanges(GetIncomeRange(), [ponctualSourceRange, fixSourceRange, variableSourceRange]);

  // Get Conditionnal Format Rules

  // Ponctual Showdowns
  const ponctualProbabilityRules = GetConditionnalFormattingRulesExact(GetProbabilityRange(), ponctualProbabilityRange);
  const ponctualSourceRules = GetConditionnalFormattingRulesExact(GetIncomeRange(), ponctualSourceRange);

  // Fix Showdowns
  const fixProbabilityRules = GetConditionnalFormattingRulesExact(GetProbabilityRange(), fixProbabilityRange);
  const fixSourceRules = GetConditionnalFormattingRulesExact(GetIncomeRange(), fixSourceRange);

  // Variable Showdowns
  const monthRules = GetConditionnalFormattingRulesContain(GetMonthsRange(), variableMonthRange);
  const variableProbabilityRules = GetConditionnalFormattingRulesExact(GetProbabilityRange(), variableProbabilityRange);
  const variableSourceRules = GetConditionnalFormattingRulesExact(GetIncomeRange(), variableSourceRange);

  // Make Rules

  const rules = [
    ...ponctualProbabilityRules,
    ...ponctualSourceRules,
    ...fixProbabilityRules,
    ...fixSourceRules,
    ...monthRules,
    ...variableProbabilityRules,
    ...variableSourceRules,
  ]

  sheet.setConditionalFormatRules(rules);
}

///// EXECUTION /////___________________________________________________________________________

function RefreshShowdowns() {
  SetBacklogShowdowns();
  SetSprintConditionnalFomatRules();
  SetTimeManagementBars();
  SetTimeManagementConditionnalFormattingRules();
  SetScopeConditionnalFormatingRules();
  SetScopeBorders();
  RefreshAnnualBudgetsStyle();
  RefreshProductionBudgetsStyle();
  RefreshCustomBudgetStyle();
  return;
}

////////////////////////////
//          SETUP         //________________________________________________________________________________________________________________________________________
////////////////////////////

///// RANGES /////___________________________________________________________________________

function GetJobsRange() {
  const sheet = setupSheet;
  const jobsColumn = getColumnFromName(sheet, jobsName, setupTabsRow, setupTabsCol, setupTabsWidth);
  const range = sheet.getRange(jobsColumn + (setupTabsRow + 1) + ':' + jobsColumn + setupLength);
  return range;
}

function GetTeamRange() {
  const sheet = setupSheet;
  const teamColumn = getColumnFromName(sheet, teamName, setupTabsRow, setupTabsCol, setupTabsWidth);
  const range = sheet.getRange(teamColumn + (setupTabsRow + 1) + ':' + teamColumn + setupLength);
  return range;
}

function GetMonthsRange() {
  const sheet = setupSheet;
  const monthColumn = getColumnFromName(sheet, monthsName, setupTabsRow, setupTabsCol, setupTabsWidth);
  const range = sheet.getRange(monthColumn + (setupTabsRow + 1) + ':' + monthColumn + setupLength);
  return range;
}

const months = GetValues(GetMonthsRange())

function GetImportanceRange() {
  const sheet = setupSheet;
  const importanceColumn = getColumnFromName(sheet, importanceName, setupTabsRow, setupTabsCol, setupTabsWidth);
  let range = sheet.getRange(importanceColumn + (setupTabsRow + 1) + ':' + importanceColumn + setupLength);
  return range;
}

function GetStatusRange() {
  const sheet = setupSheet;
  const statusColumn = getColumnFromName(sheet, satusName, setupTabsRow, setupTabsCol, setupTabsWidth);
  const range = sheet.getRange(statusColumn + (setupTabsRow + 1) + ':' + statusColumn + setupLength);
  return range;
}

function GetProductionRange() {
  const sheet = setupSheet;
  const productionColumn = getColumnFromName(sheet, productionStageName, setupTabsRow, setupTabsCol, setupTabsWidth);
  const range = sheet.getRange(productionColumn + (setupTabsRow + 1) + ':' + productionColumn + setupLength);
  return range;
}

///// VALUES /////___________________________________________________________________________

function GetDarkModeCellA1() {
  const sheet = setupSheet;
  const parameterColumn = getColumnFromName(sheet, parameterName, setupTabsRow, setupTabsCol, setupTabsWidth);
  const parameterCol = getColumnFromA1(parameterColumn);
  const darkModeRow = getRowFromName(sheet, darkModeName, setupTabsRow, parameterCol, setupLength);

  return parameterColumn + (darkModeRow+1);
}

function GetDarkModeValue() {
  return setupSheet.getRange(GetDarkModeCellA1()).getValue();
}

function GetAutoBorderCellA1() {
  const sheet = setupSheet;
  const parameterColumn = getColumnFromName(sheet, parameterName, setupTabsRow, setupTabsCol, setupTabsWidth);
  const parameterCol = getColumnFromA1(parameterColumn);
  const autoBorderRow = getRowFromName(sheet, autoBorderName, setupTabsRow, parameterCol, setupLength) + 1;

  return parameterColumn + autoBorderRow;
}

function GetAutoBorderValue() {
  return setupSheet.getRange(GetAutoBorderCellA1()).getValue();
}

function GetAutoFillCellA1() {
  const sheet = setupSheet;
  const parameterColumn = getColumnFromName(sheet, parameterName, setupTabsRow, setupTabsCol, setupTabsWidth);
  const parameterCol = getColumnFromA1(parameterColumn);
  const autoFillRow = getRowFromName(sheet, autoFillName, setupTabsRow, parameterCol, setupLength) + 1;

  return parameterColumn + autoFillRow;
}

function GetAutoFillValue() {
  return setupSheet.getRange(GetAutoFillCellA1()).getValue();
}

////////////////////////////
//         BACKLOG        //_______________________________________________________________________________________________________________________________________
////////////////////////////

///// AUTOFILL /////___________________________________________________________________________

function AutoFillBacklogEmptyRow(row) {
  const sheet = backlogSheet;

  // Get Columns & Rows

  const firstColumn = getA1FromCol(backlogTabsCol);
  const lastColumn = getA1FromCol(backlogTabsWidth);
  const dayEstimatedColumn = getColumnFromName(sheet, nbDayEstName, backlogTabsRow, backlogTabsCol, backlogTabsWidth)

  // Get Range & Values
  const range = sheet.getRange(firstColumn + row + ':' + lastColumn + row);
  const values = range.getValues();
  const previousValues = sheet.getRange(firstColumn + (row - 1) + ':' + lastColumn + (row - 1)).getValues();
  const toDoValue = GetStatusRange().getValues()[0][0];

  // Set Rows
  if (values[0][0] == '') {
    range.setValues([[previousValues[0][0], previousValues[0][1], '', '', '=' + dayEstimatedColumn + row + '*1,3', '', '', toDoValue, '', '', '', '', '']]);
  }
}

function SwitchAutoEmptyRow() {
  const range = setupSheet.getRange(GetAutoFillValue());
  if (range.getValue() == true) {
    range.setValue(false);
    SpreadsheetApp.getActive().toast("AutoFill Deactivated");
  } else {
    range.setValue(true);
    SpreadsheetApp.getActive().toast("AutoFill Activated")
  };
}

///// BORDERS /////___________________________________________________________________________

function GetBacklogBordersSequences() {
  const sheet = backlogSheet;
  const dataRow = backlogTabsRow + 1;

  // get Columns
  const epicColumn = getColumnFromName(sheet, epicName, backlogTabsRow, backlogTabsCol, backlogTabsWidth)
  const storiesColumn = getColumnFromName(sheet, storiesName, backlogTabsRow, backlogTabsCol, backlogTabsWidth)

  // Get Ranges
  const epicRange = sheet.getRange(epicColumn + dataRow + ':' + epicColumn + backlogLength)
  const storiesRange = sheet.getRange(storiesColumn + dataRow + ':' + storiesColumn + backlogLength)

  // Get Values
  const epicValues = GetValues(epicRange)
  const storiesValues = GetValues(storiesRange)

  // Get Sequences
  const epicsSequence = GetSequence(epicValues);
  const storiesSequence = GetSequence(storiesValues);

  return [epicsSequence, storiesSequence];
}

function TraceBacklogBorders() {
  const sheet = backlogSheet;

  // get Columns

  const epicColumn = getColumnFromName(sheet, epicName, backlogTabsRow, backlogTabsCol, backlogTabsWidth);
  const storiesColumn = getColumnFromName(sheet, storiesName, backlogTabsRow, backlogTabsCol, backlogTabsWidth);
  const firstColumn = getA1FromCol(backlogTabsCol);
  const lastColumn = getA1FromCol(backlogTabsWidth);

  // Get Data Range
  const range = sheet.getRange(firstColumn + '2:' + lastColumn + backlogLength);

  // Reset Border
  ResetBorders(range);

  // Get Sequences
  const sequences = GetBacklogBordersSequences();


  // Get the ranges coordinate 

  // Initialize
  const startColumns = [epicColumn, storiesColumn];
  const endColumns = [epicColumn, lastColumn]
  ranges = []


  // Get the border Ranges
  for (let i in sequences) {
    let startRow = 2;
    for (let j in sequences[i]) {
      const borderLength = sequences[i][j];
      const endRow = startRow + borderLength - 1;
      const index = startColumns[i] + startRow + ':' + endColumns[i] + endRow;
      ranges.push(index);
      startRow = endRow + 1;
    }
  }

  // Set the border Ranges
  
  const rangeList = sheet.getRangeList(ranges);
  rangeList.setBorder(true, true, true, true, null, null, getHighlightcolor1(), SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  return;
}

function SwitchAutoBorder() {
  const range = setupSheet.getRange(GetAutoBorderCellA1());
  if (range.getValue() == true) {
    range.setValue(false);
    SpreadsheetApp.getActive().toast("AutoBorder Deactivated");
  } else {
    range.setValue(true);
    SpreadsheetApp.getActive().toast("AutoBorder Activated")
  };
}

////////////////////////////
//          SPRINT         //__________________________________________________________________________________________________________________________________________
////////////////////////////

///// SPRINT SELECTION /////___________________________________________________________________________

function ChangePivotTableFilter() {
  const sheet = sprintSheet;

  // Get Columns
  const backlogSprintColumn = getColumnFromName(backlogSheet, sprintName, backlogTabsRow, backlogTabsCol, backlogTabsWidth);
  const backlogSprintCol = getColumnFromA1(backlogSprintColumn);

  // Get Sprint Name
  const sprintSheetName = sheet.getRange(sprintTitleColumn + sprintValueRow).getValue();

  // Get Pivot Table
  const pivotTable = sheet.getPivotTables()[0];

  // Clean Filter
  const filters = pivotTable.getFilters();
  filters.map(f => f.remove());

  // Add New Filter
  const criteria = SpreadsheetApp.newFilterCriteria()
    .setVisibleValues([sprintSheetName])
    .build();
  pivotTable.addFilter(backlogSprintCol, criteria);
}

///// DATE MANAGEMENT /////___________________________________________________________________________

function ResetGanttDates() {
  const sheet = sprintSheet;
  
  // Get Row & Columns

  const firstRow = sprintTabsRow + 1
  const teamColumn = getColumnFromName(sheet, teamName, sprintTabsRow, sprintCol, sprintTabsWidth);
  const datesColumn = getColumnFromName(sheet, planningStartingDaysName, sprintTabsRow, sprintCol, sprintLength);
  const nbDaysColumn = getColumnFromName(sheet, nbDayEst30Name, sprintTabsRow, sprintCol, sprintTabsWidth)

  // Get Sheet Data
  
  const teamRange = sheet.getRange(teamColumn + firstRow + ':' + teamColumn + sheet.getLastRow());
  const teamValues = GetValues(teamRange)
  const setupTeamValues = GetValues(GetTeamRange());

  const datesRange = sheet.getRange(datesColumn + firstRow + ':' + datesColumn + sheet.getLastRow());
  const datesValues = GetValues(datesRange)

  const nbDaysRange = sheet.getRange(nbDaysColumn + firstRow + ':' + nbDaysColumn + sheet.getLastRow());
  const nbDaysValues = GetValues(nbDaysRange);

  let previousDate = new Date(sheet.getRange(sprintStartingDayColumn + sprintValueRow).getValue());
  let first0ffset = nbDaysValues[0]

  // Prepare Arrays

  const newDates = []; // Double Array containing the days for each members
  const teamDateCells = [] // Double Array containing the range coordonate corresponding to the team member
  const daysOffsets = [] // Array Containing the day offset for each members
  let previousTeamDates = [] // Array containing the previous date of each members

  // Calculate Dates

  let currentDate;

  // Initialize Arrays
  for (let member in setupTeamValues) {
    newDates.push([[previousDate]]);
    teamDateCells.push([]);
    daysOffsets.push(first0ffset)
    previousTeamDates.push(previousDate)
  }

  for (let index in teamValues) {

    // Get Variables
    const currentTeamMember = teamValues[index];
    const currentTeamMemberIndex = setupTeamValues.indexOf(currentTeamMember);
    const currentPreviousDate = new Date(previousTeamDates[currentTeamMemberIndex])

    // Get The Offset
    daysOffsets[currentTeamMemberIndex] = Math.round(nbDaysValues[index]) - 1;

    const sprintDateCell = datesColumn + ((parseInt(index) + parseInt(firstRow)));

    // Get The new date
    currentDate = new Date(currentPreviousDate);

    // Increment
    currentDate.setDate(currentDate.getDate() + 1);

    // Add offset
    currentDate.setDate(currentDate.getDate() + daysOffsets[currentTeamMemberIndex])

    // Check if week-end
    if (currentDate.getUTCDay() == 5) {
      currentDate.setDate(currentDate.getDate() + 2);
    }
    if (currentDate.getUTCDay() == 6) {
      currentDate.setDate(currentDate.getDate() + 2);
    }

    //Add the date and the cell to the corresponding Array
    newDates[currentTeamMemberIndex].push([currentDate])
    teamDateCells[currentTeamMemberIndex].push(sprintDateCell)

    previousTeamDates[currentTeamMemberIndex] = new Date(currentDate)
  };

  // Set The Dates

  for (let member in setupTeamValues) {
    for (let index in teamDateCells[member]) {
      sheet.getRange(teamDateCells[member][index]).setValue(newDates[member][index])
    }
  }
}

function CalculateSprintDuration() {
  const sheet = sprintSheet;

  // Get Rows & Columns

  const firstRow = sprintTabsRow + 1;
  const teamColumn = getColumnFromName(sheet, teamName, sprintTabsRow, sprintCol, sprintTabsWidth)
  const timeColumn = getColumnFromName(sheet, nbDayEst30Name, sprintTabsRow, sprintCol, sprintTabsWidth)

  // Get Datas

  const setupTeamValues = GetValues(GetTeamRange());
  const timeRange = sheet.getRange(timeColumn + firstRow + ':' + timeColumn + sprintLength)
  const timeValues = GetValues(timeRange);

  // Calculate Sums

  // Initialize Array
  const sums = []
  for (member in setupTeamValues) sums.push(0)

  // Calculate
  for (let index in timeValues) {
    const currentTeamCell = teamColumn + (parseInt(firstRow) + parseInt(index));
    const currentTeamValue = sheet.getRange(currentTeamCell).getValue();
    const currentTeamIndex = setupTeamValues.indexOf(currentTeamValue);
    const currentTimeValue = timeValues[index];

    sums[currentTeamIndex] += currentTimeValue;
  }

  // Return the maximum number of weeks
  return Math.max(...sums) / 5
}

function SetSprintDuration() {
  const sheet = sprintSheet;

  // Get Columns
  const datesColumn = getColumnFromName(sheet, planningStartingDaysName, sprintTabsRow, sprintCol, sprintLength)

  // Calculate Sprint Duration
  const sprintDuration = CalculateSprintDuration().toPrecision(2)

  // Set Value
  sheet.getRange(datesColumn + sprintValueRow).setValue(sprintDuration + ' Weeks');
}

function SetMonday(date) {
  // Get The date Cell
  const range = sprintSheet.getRange(sprintStartingDayColumn + sprintValueRow)

  // Set Date to Monday
  date = new Date(range.getValue())
  date.setHours(0);
  date.setMinutes(0);
  date.setSeconds(0);
  date.setMilliseconds(0);
  date.setDate(date.getDate() - date.getUTCDay());

  // Set the Monday Value
  range.setValue(date);
}

///// TOOLS /////___________________________________________________________________________

function SetSelectionToDoing() {
  const targetSheet = backlogSheet;

  // get Columns & Rows
  const backlogStatusColumn = getColumnFromName(targetSheet, satusName, backlogTabsRow, backlogTabsCol, backlogTabsWidth);
  const backlogTaskColumn = getColumnFromName(targetSheet, taskName, backlogTabsRow, backlogTabsCol, backlogTabsWidth)

  // Get Selection
  const cell = ss.getActiveCell();
  const task = cell.getValue();

  // Get Doing Value
  const doingValue = GetStatusRange().getValues()[1][0];

  // Localize Selected Task in Backlog
  const taskRange = targetSheet.getRange(backlogTaskColumn + backlogTabsRow + ':' + backlogTaskColumn + backlogLength);
  const taskList = cleanData(taskRange.getValues());
  const taskRow = taskList.indexOf(task) + 1;
  const backlogTaskStatusCell = targetSheet.getRange(backlogStatusColumn + taskRow);

  // Set Status Value
  backlogTaskStatusCell.setValue(doingValue)
}

function SetSelectionToDone() {
  const targetSheet = backlogSheet;

  // get Columns & Rows
  const backlogStatusColumn = getColumnFromName(targetSheet, satusName, backlogTabsRow, backlogTabsCol, backlogTabsWidth);
  const backlogTaskColumn = getColumnFromName(targetSheet, taskName, backlogTabsRow, backlogTabsCol, backlogTabsWidth)

  // Get Selection
  const cell = ss.getActiveCell();
  const task = cell.getValue();

  // Get Done Value
  const doneValue = GetStatusRange().getValues()[2][0];

  // Localize Selected Task in Backlog
  const taskRange = targetSheet.getRange(backlogTaskColumn + backlogTabsRow + ':' + backlogTaskColumn + backlogLength);
  const taskList = cleanData(taskRange.getValues());
  const taskRow = taskList.indexOf(task) + 1;
  const backlogTaskStatusCell = targetSheet.getRange(backlogStatusColumn + taskRow);

  // Set Status Value
  backlogTaskStatusCell.setValue(doneValue)
}

////////////////////////////
//         PLANNING       //_______________________________________________________________________________________________________________________________________
////////////////////////////

function CalculateMondaysWeeks(startingDate, nbWeeks) {

  // Calculate an array of all the monday of the duration

  // Set First Monday
  let monday = new Date(startingDate);
  monday.setDate(monday.getDate() - monday.getUTCDay())
  const mondays = [new Date(monday)];

  //Get Every Monday of the planning
  for (let k = 0; k < nbWeeks; k++) {
    mondays.push(monday)
    monday = new Date(monday.setDate(monday.getDate() + 7))
  }

  return mondays
}

function CreatePlanning() {
  const sheet = planningSheet;
  const firstColumn = getA1FromCol(planningCol);
  const secondColumn = getColumnOffset(firstColumn, 1);
  const seconCol = getColumnFromA1(secondColumn);
  const lastColumn = getA1FromCol(planningWidth);
  const maxColumn = getA1FromCol()

  // Cleaning the Sheet
  const planningMaxRange = sheet.getRange(firstColumn + planningfirstRow + ':' + maxColumn + planningLength);
  planningMaxRange.clearDataValidations();
  planningMaxRange.setValue('');
  planningMaxRange.breakApart();
  ResetBorders(planningMaxRange)

  // Variables Initialization

  // Row & Columns
  const startingRow = planningfirstRow;
  let previousRow = startingRow;
  let currentRow = 0;
  let currentColumn = seconCol;

  // Dates
  const startingDate = new Date(sheet.getRange(planningValueColumn + planningStartingDateRow).getValue());
  const nbWeeks = sheet.getRange(planningValueColumn + planningWeeksRow).getValue()
  const mondays = CalculateMondaysWeeks(startingDate, nbWeeks);
  const firstTuesday = new Date(mondays[0])
  firstTuesday.setDate(firstTuesday.getDate() + 3)

  // Trackers
  let weeksCount = 0;
  let quarter = 1;

  // Ranges
  const sprintRanges = [];
  const monthRanges = [];
  const borderRanges = [];

  for (let mondayIndex in mondays) {

    if (mondayIndex < (mondays.length - 1)) {
      // Setup the first Column of the Planning
      if (currentRow != previousRow) {
        if (currentRow == 0) {
          currentRow = startingRow;
        };
        previousRow = currentRow;

        // Get Ranges
        sheet.getRange(firstColumn + currentRow + ':' + firstColumn + (currentRow + 3)).setValues([['Quarter ' + quarter], [monthsName], ['Weeks'], ['Sprints']]);
        sheet.getRange(firstColumn + (currentRow + 3)).setFontColor(getHighlightcolor1());

        // Save Border Ranges
        borderRanges.push(firstColumn + (currentRow + 1) + ':' + firstColumn + (currentRow + 3))

        // Set Cells Dimensions
        sheet.setRowHeight(currentRow, planningQuarterCellsHeigth);
        sheet.setRowHeight(currentRow + 1, planningCellsHeigth);
        sheet.setRowHeight(currentRow + 2, planningCellsHeigth);
        sheet.setRowHeight(currentRow + 3, planningCellsHeigth * 2);

        quarter++
        if (quarter == 5) quarter = 1; // Reset quarters
      }

      // Get Current Tuesday
      const tuesday = new Date(mondays[mondayIndex])
      tuesday.setDate(tuesday.getDate() + 3)

      // Get Ranges
      monthCell = sheet.getRange(currentRow + 1, currentColumn);
      weekCell = sheet.getRange(currentRow + 2, currentColumn);

      // Set Values
      monthCell.setValue(Utilities.formatDate(new Date(tuesday), "GMT+1", "MMMMMMM yyyy"));
      weekCell.setValue(Utilities.formatDate(new Date(mondays[mondayIndex]), "GMT+1", "'Week' ww '('dd/MM')'"));

      // Save Borders
      const sprintColumn = getA1FromCol(currentColumn)
      borderRanges.push(sprintColumn + (currentRow + 2) + ':' + sprintColumn + (currentRow + 3))

      // Set Ranges
      const quarterRange = [currentRow, 1, 1, currentColumn]
      const monthRange = [currentRow + 1, 2, 1, currentColumn - 1]
      const sprintRange = [currentRow + 3, 2, 1, currentColumn - 1]

      // Set Column Width
      sheet.setColumnWidth(currentColumn, planningCellsWidth)

      // End the Row
      weeksCount++

      // End of the quarter
      if (weeksCount == 13 || mondayIndex == mondays.length - 2) {
        currentRow += 4;
        currentColumn = 1;
        weeksCount = 0;
        sprintRanges.push(sheet.getRange(sprintRange[0], sprintRange[1], sprintRange[2], sprintRange[3]))
        monthRanges.push(monthRange);
        sheet.getRange(quarterRange[0], quarterRange[1], quarterRange[2], quarterRange[3])
          .setFontColor(getHighlightcolor1())
          .setFontSize(16)
          .setVerticalAlignment('middle')
          .merge();
      }
      currentColumn++;
    }
  }

  SetPlanningShowdown(sprintRanges)

  //Months Merging & Conditionnal Format Rules
  let rules = []
  for (let i = 0; i < monthRanges.length; i++) {

    let initialColumn = 2;

    // Get Range
    const currentRange = sheet.getRange(monthRanges[i][0], monthRanges[i][1], monthRanges[i][2], monthRanges[i][3])

    // Get Values
    const monthValues = currentRange.getValues()

    // Get Month Sequence
    const sequences = GetSequence(monthValues[0])

    // Get Month Conditionnal Format Rule
    const newRules = GetConditionnalFormattingRulesContain(GetMonthsRange(), currentRange)
    rules = rules.concat(newRules)

    //Merge Cells & Save Border Ranges
    for (let j = 0; j < sequences.length; j++) {
      sheet.getRange(monthRanges[i][0], initialColumn, monthRanges[i][2], sequences[j])
        .merge()
      borderRanges.push(getR1C1Notation(monthRanges[i][0], initialColumn, 3, sequences[j]))
      borderRanges.push(getR1C1Notation(monthRanges[i][0], initialColumn, monthRanges[i][2], sequences[j]))
      initialColumn += sequences[j];
    }
  }

  // Set Borders And Rules

  //Save Borders Range into properties
  PropertiesService.getScriptProperties().setProperty('planningBordersCells', JSON.stringify(borderRanges));

  // Set Borders
  sheet.getRangeList(borderRanges)
    .setBorder(true, true, true, true, null, null, GetDarkModeColor(), SpreadsheetApp.BorderStyle.SOLID_THICK)
    .setVerticalAlignment('middle')

  // Set Rules
  sheet.setConditionalFormatRules(rules)
}

////////////////////////////
//          SCOPE         //__________________________________________________________________________________________________________________________________________
////////////////////////////

function SetScopeBorders() {
  const sheet = scopeSheet;

  // get Columns & Rows

  const scopeLengthRedux = scopeLength - 2;

  const firstRow = scopeTabsRow + 1;
  const firstColumn = scopeCol;
  let startRow;
  let endRow;
  let startColumn;
  const endColumn = sheet.getLastColumn();
  const tabsWidth = endColumn - firstColumn;
  const lastTabColumn = getA1FromCol(sheet.getLastColumn());
  const lastRow = getRowFromName(sheet, grandTotalName, scopeRow, scopeCol, scopeLength);

  const productionColumn = getColumnFromName(sheet, productionStageName, scopeTabsRow, scopeCol, scopeVerticalTabsWidth);
  const teamColumn = getColumnFromName(sheet, backlogTeamName, scopeTabsRow, scopeCol, scopeVerticalTabsWidth);
  const jobColumn = getColumnFromName(sheet, jobsName, scopeTabsRow, scopeCol, scopeVerticalTabsWidth);

  // Get Ranges

  const productionRange = sheet.getRange(productionColumn + firstRow + ':' + productionColumn + scopeLength);
  const teamRange = sheet.getRange(teamColumn + firstRow + ':' + teamColumn + scopeLength);
  const jobsRange = sheet.getRange(jobColumn + firstRow + ':' + jobColumn + scopeLength);

  const statusTotalRange = getR1C1Notation(firstRow, firstColumn + tabsWidth, firstRow + lastRow - 1, 1);
  const bottomTotalRange = getR1C1Notation(firstRow - 1, firstColumn, 1, firstColumn + tabsWidth - 1);
  const tabRange = getR1C1Notation(firstRow, firstColumn, lastRow + 1 - firstRow, firstColumn + tabsWidth - 1);

  const productionValues = productionRange.getValues();
  const teamValues = teamRange.getValues();
  const jobsValues = cleanData(jobsRange.getValues());

  // Get Setup Datas

  const setupJobValues = GetValues(GetJobsRange());
  const setupJobColors = GetColors(GetJobsRange());

  // Get Sequences

  const productionSequence = GetSequenceIgnoreEmpty(productionValues);
  const teamSequence = GetSequenceIgnoreEmpty(teamValues);

  // Reset Borders

  const sheetRange = sheet.getRange(scopeRow, scopeCol, scopeLengthRedux, scopeWidth);
  ResetBorders(sheetRange)

  // Get Border Ranges
  let borderRanges = [];
  const productionRanges = [];
  const productionTotalRanges = [];
  let totalsRanges = [];
  const teamRanges = [];
  const teamTotalRanges = [];

  if (sheet.getRange(scopeValueColumn + scopeBorderRow).getValue() == true) {

    // Get Production Borders

    startColumn = firstColumn;
    startRow = firstRow;
    endRow = firstRow;

    for (let k = 0; k < (productionSequence.length); k++) {
      endRow = startRow + productionSequence[k];

      let borderRange = getR1C1Notation(startRow, startColumn, endRow - startRow + 1, startColumn + tabsWidth - 1);

      productionRanges.push(borderRange);

      borderRange = getR1C1Notation(endRow, startColumn, 1, startColumn + tabsWidth - 1);
      productionTotalRanges.push(borderRange);

      borderRange = getR1C1Notation(endRow, startColumn + tabsWidth, 1, 1);
      totalsRanges.push(borderRange);

      startRow = endRow + 1;
      endRow = startRow;
    }

    // Get Team Borders

    startColumn = startColumn + 1;
    startRow = firstRow;
    endRow = firstRow;

    let temp = 0;

    for (let k = 0; k < teamSequence.length; k++) {
      endRow = startRow + teamSequence[k];
      let borderRange = getR1C1Notation(startRow + temp, startColumn, endRow - startRow + 1 - temp, startColumn + tabsWidth - 3);
      const borderValues = cleanData(sheet.getRange(borderRange).getValues());

      const totalTeamRange = getR1C1Notation(endRow, startColumn, 1, startColumn + tabsWidth - 3);
      const totalRange = getR1C1Notation(endRow, startColumn + tabsWidth - 1, 1, 1);

      if (borderValues[0] != '') {
        teamRanges.push(borderRange);
        teamTotalRanges.push(totalTeamRange);
        startRow = endRow + 1;
        temp = 0;
      } else {
        startRow = endRow + 1;
        temp = -1;
      }
      endRow = startRow;
    }

    // Set Borders

    // Combine Ranges
    totalsRanges = [
      bottomTotalRange,
      statusTotalRange,
      ...totalsRanges
    ];

    borderRanges = [
      ...productionRanges,
      ...productionTotalRanges,
      ...teamRanges,
      ...teamTotalRanges,
      statusTotalRange
    ];

    // Get RangeList
    const borderRangeList = sheet.getRangeList(borderRanges);
    const totalRangeList = sheet.getRangeList(totalsRanges);


    // Set Tabs Border
    borderRangeList.setBorder(true, true, true, true, null, null, GetDarkModeColor(), SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    totalRangeList.setBorder(true, true, true, true, null, null, GetDarkModeColor(), SpreadsheetApp.BorderStyle.SOLID_THICK);
    sheet.getRange(tabRange).setBorder(true, true, true, true, null, null, getHighlightcolor1(), SpreadsheetApp.BorderStyle.SOLID_THICK);

    // Set Jobs Borders
    startRow = firstRow;
    startColumn = firstColumn;

    for (let index in jobsValues) {
      if (jobsValues[index] != 0) {
        if (parseInt(index) < (jobsValues.length - 1)) {

          const currentRow = parseInt(firstRow) + parseInt(index);

          if (jobsValues[parseInt(index) + 1] != '') {
            const jobRange = sheet.getRange(jobColumn + currentRow + ':' + lastTabColumn + currentRow);
            const jobIndex = setupJobValues.indexOf(jobsValues[index]);
            const jobColor = setupJobColors[jobIndex];
            jobRange.setBorder(null, null, true, null, null, null, jobColor, SpreadsheetApp.BorderStyle.SOLID);
            jobRange.getValues(); // I don't get why but without that line the code isn't working. It makes the code way slower tho. 
          }
        }
      }
    }
  }
}

function RemoveScopeBorders() {
  scopeSheet.getRange(scopeValueColumn + scopeBorderRow).setValue(false);
  SetScopeBorders();
}

function CreateScopeBorders() {
  scopeSheet.getRange(scopeValueColumn + scopeBorderRow).setValue(true);
  SetScopeBorders();
}

function SwitchScopeBordersStyle() {
  const range = scopeSheet.getRange(scopeValueColumn + scopeBorderRow);
  if (range.getValue()) {
    range.setValue(false);
    SetScopeBorders();
  } else {
    range.setValue(true);
    SetScopeBorders();
  }
}

function SwitchScopeBackgroundStyle() {
  const range = scopeSheet.getRange(scopeValueColumn + scopeBackgroundStyleRow);
  if (range.getValue()) {
    range.setValue(false);
    SetScopeConditionnalFormatingRules();
  } else {
    range.setValue(true);
    SetScopeConditionnalFormatingRules();
  }
}

////////////////////////////
//     SECONDARY SCOPES   //__________________________________________________________________________________________________________________________________________
////////////////////////////


function CreateImportanceScope(targetSheet, column, row) {
  const sheet = importanceScopeSheet;
  const tabsRow = 2;
  const length = sheet.getLastRow()

  const targetImportanceColumn = getColumnOffset(column, 1);
  const targetDaysColumn = getColumnOffset(column, 2);
  const targetMonthColumn = getColumnOffset(column, 3);

  // Get Importance Scope Data

  // Get Team Values
  const teamColumn = 'A';
  const teamCol = 1;
  const teamRow = getRowFromName(sheet, backlogTeamName, tabsRow, teamCol, length);
  const teamValues = cleanData(sheet.getRange(teamColumn + (teamRow + 1) + ':' + teamColumn + length).getValues().filter(String));

  // Get Importance Values
  const importanceColumn = 'B';
  const lastColumn = getA1FromCol(sheet.getLastColumn());
  const importanceValues = cleanData(transposeArray(sheet.getRange(importanceColumn + tabsRow + ':' + lastColumn + tabsRow).getValues()).filter(String));

  // Prepare Tabs

  const tabsValues = [['Importance Scope', 'Importance', 'Min Nb of Days', 'Min Nb of Months']];

  // Get Formulas
  const formulas = GetImportanceScopeFormulas(importanceValues, teamValues);


  // Get Tabs Values
  for (let k = 0; k < formulas.length; k++) {
    if (k != 0) formulas[k] += '+' + targetDaysColumn + (row + k);
    const monthFormulas = ('=CONCAT(ceiling(' + targetDaysColumn + (row + k + 1) + '/21)," Months")');

    tabsValues.push(['Importance Scope', importanceValues[k], formulas[k], monthFormulas]);
  }

  // Set Values
  const targetRange = targetSheet.getRange(column + row + ':' + targetMonthColumn + (row + importanceValues.length));
  targetRange.setValues(tabsValues);

  // Merge First Column
  targetSheet.getRange(column + row + ':' + column + (row + importanceValues.length)).merge();
}

function MakeImportanceScopeFormula(importance, teamValues) {
  let formula = '=MAX(';
  for (let i = 0; i < teamValues.length; i++) {
    const teamFormulaPart = 'GETPIVOTDATA($A$1, ' + importanceScopeName + '!$A$1,"' + importanceName + '","' + importance + '","' + backlogTeamName + '","' + teamValues[i] + '"),';
    formula += teamFormulaPart;
  }
  formula += ')';
  return formula;
}

function GetImportanceScopeFormulas(importanceValues, teamValues) {
  const formulas = [];
  for (let j = 0; j < importanceValues.length; j++) {
    formulas.push(MakeImportanceScopeFormula(importanceValues[j], teamValues));
  }
  return formulas;
}

function CreateProductionScope(targetSheet, column, row) {
  const sheet = productionScopeSheet;
  const tabsRow = 2;
  const length = sheet.getLastRow();

  const targetImportanceColumn = getColumnOffset(column, 1);
  const targetDaysColumn = getColumnOffset(column, 2);
  const targetMonthColumn = getColumnOffset(column, 3);

  // Get Production Values
  const teamColumn = 'A';
  const teamCol = 1;
  const teamRow = getRowFromName(sheet, backlogTeamName, tabsRow, teamCol, length);
  const teamValues = cleanData(sheet.getRange(teamColumn + (teamRow + 1) + ':' + teamColumn + length).getValues().filter(String));

  // Get Production Values
  const productionColumn = 'B';
  const lastColumn = getA1FromCol(sheet.getLastColumn());
  const productionValues = cleanData(transposeArray(sheet.getRange(productionColumn + tabsRow + ':' + lastColumn + tabsRow).getValues()).filter(String));

  // Prepare Tabs

  const tabsValues = [['Production Scope', 'Production Stage', 'Min Nb of Days', 'Min Nb of Months']];

  // Get Formulas
  const formulas = GetProductionScopeFormulas(productionValues, teamValues);

  // Get Tabs Values
  for (let k = 0; k < formulas.length; k++) {
    //if (k != 0) formulas[k] +='+'+targetDaysColumn+(row+k);
    const monthFormulas = ('=CONCAT(ceiling(' + targetDaysColumn + (row + k + 1) + '/21)," Months")');

    tabsValues.push(['Production Scope', productionValues[k], formulas[k], monthFormulas]);
  }

  // Set Values
  const targetRange = targetSheet.getRange(column + row + ':' + targetMonthColumn + (row + productionValues.length));
  targetRange.setValues(tabsValues);

  // Merge First Column
  targetSheet.getRange(column + row + ':' + column + (row + productionValues.length)).merge();
}


function MakeProductionScopeFormula(productionStage, teamValues) {
  let formula = '=MAX(';
  for (let i = 0; i < teamValues.length; i++) {
    const teamFormulaPart = 'GETPIVOTDATA($A$1, ' + productionScopeName + '!$A$1,"' + productionStageName + '","' + productionStage + '","' + backlogTeamName + '","' + teamValues[i] + '"),';
    formula += teamFormulaPart;
  }
  formula += ')';
  return formula;
}

function GetProductionScopeFormulas(productionValues, teamValues) {
  const formulas = [];
  for (let j = 0; j < productionValues.length; j++) {
    formulas.push(MakeProductionScopeFormula(productionValues[j], teamValues));
  }
  return formulas;
}

////////////////////////////
//     TIME MANAGEMENT    //__________________________________________________________________________________________________________________________________________
////////////////////////////

function SetTimeManagementBars() {
  const sheet = timeManagementSheet;

  const length = sheet.getLastRow();
  const tabsRow = timeManagementTabsRow;
  const dataRow = tabsRow + 1;

  const col = timeManagementCol;
  const width = sheet.getLastColumn();

  // Get Setup Datas

  const setupImportanceRange = GetImportanceRange();
  const setupProductionRange = GetProductionRange();
  const setupTeamRange = GetTeamRange();
  const setupJobsRange = GetJobsRange();

  // Get Setup Lengths
  const jobsLength = GetValues(setupJobsRange).length;

  // Get Columns & Rows

  // Main Bar
  const mainBarTotalColumn = getColumnOffset(timeManagementMainBarColumn, -1);
  const mainBarValueColumn = getColumnOffset(timeManagementMainBarColumn, -2);

  // Progression Tabs
  const progressionColumn = getColumnFromName(sheet, importanceName, tabsRow, col, width);

  const progressionToDoColumn = getColumnOffset(progressionColumn, 1);
  const progressionTotalColumn = getColumnOffset(progressionColumn, 4);
  const progressionBarColumn = getColumnOffset(progressionColumn, 5);

  // Job Tab
  const jobColumn = getColumnFromName(sheet, jobsName, tabsRow, col, width);

  const jobToDoColumn = getColumnOffset(jobColumn, 1);
  const jobTotalColumn = getColumnOffset(jobColumn, 4);
  const jobBarColumn = getColumnOffset(jobColumn, 5);

  // Sprint Tab
  const sprintColumn = getColumnFromName(sheet, sprintName, tabsRow, col, width);

  const sprintTeamColumn = getColumnOffset(sprintColumn, 1);
  const sprintDaysColumn = getColumnOffset(sprintColumn, 2);
  const sprintBarColumn = getColumnOffset(sprintColumn, 3);

  // Epic Tab
  const epicColumn = getColumnFromName(sheet, epicName, tabsRow, col, width);

  const storiesColumn = getColumnOffset(epicColumn, 1);
  const storiesDayColumn = getColumnOffset(epicColumn, 2);
  const storiesTotalColumn = getColumnOffset(epicColumn, 3);
  const storiesBarColumn = getColumnOffset(epicColumn, 4);

  // Get Ranges

  // Sheet Ranges
  const progressionRange = sheet.getRange(progressionColumn + dataRow + ':' + progressionColumn + sheet.getLastRow());
  const progressionValues = cleanData(progressionRange.getValues());

  const jobRange = sheet.getRange(jobColumn + dataRow + ':' + jobColumn + sheet.getLastRow());
  const jobValues = cleanData(jobRange.getValues());

  const storiesRange = sheet.getRange(storiesColumn + dataRow + ':' + storiesColumn + sheet.getLastRow());
  const storiesValues = cleanData(storiesRange.getValues());

  // Get Sheet Datas
  const sprintTeamRange = sheet.getRange(sprintTeamColumn + dataRow + ':' + sprintTeamColumn + sheet.getLastRow());
  const sprintTeamValues = cleanData(sprintTeamRange.getValues());

  // Get Lengths
  const sprintTabLength = GetValues(sprintTeamRange).length;
  const storiesLength = GetValues(storiesRange).length;

  //Target Ranges
  const mainBarSparklineRange = sheet.getRange(timeManagementMainBarColumn + timeManagementMainBarRow);
  const progressionSparklineRange = sheet.getRange(progressionBarColumn + dataRow + ':' + progressionBarColumn + sheet.getLastRow());
  const progressionValuesRange = sheet.getRange(progressionColumn + dataRow + ':' + progressionBarColumn + sheet.getLastRow());

  const jobSparklineRange = sheet.getRange(jobBarColumn + dataRow + ':' + jobBarColumn + (dataRow + jobsLength - 3));
  const jobValueRange = sheet.getRange(jobColumn + dataRow + ':' + jobBarColumn + (dataRow + jobsLength - 3));

  const sprintSparklineRange = sheet.getRange(sprintBarColumn + dataRow + ':' + sprintBarColumn + (dataRow + sprintTabLength - 1));
  const sprintValueRange = sheet.getRange(sprintTeamColumn + dataRow + ':' + sprintBarColumn + (dataRow + sprintTabLength - 1));

  const storiesSparklineRange = sheet.getRange(storiesBarColumn + dataRow + ':' + storiesBarColumn + (dataRow + storiesLength - 1));

  // Prepare Color Matching 

  const setupImportanceValues = GetValues(setupImportanceRange)
  const setupImportanceColors = GetColors(setupImportanceRange);

  const setupProductionValues = GetValues(setupProductionRange);
  const setupProductionColors = GetColors(setupProductionRange);

  const setupTeamValues = GetValues(setupTeamRange);
  const setupTeamColors = GetColors(setupTeamRange);

  let setupJobValues = GetValues(setupJobsRange);
  const setupJobColors = GetColors(setupJobsRange);

  let keys = [
    ...setupImportanceValues,
    ...setupProductionValues,
    ...setupTeamValues
  ];


  let colors = [
    ...setupImportanceColors,
    ...setupProductionColors,
    ...setupTeamColors,
  ];

  const epicValues = sheet.getRange(epicColumn + dataRow + ':' + epicColumn + length).getValues();
  const storiesSequence = GetSequenceWithEmpty(epicValues);
  storiesSequence.pop();
  storiesSequence.push(storiesLength - storiesSequence.reduce((partialSum, a) => partialSum + a, 0)); // Get Last Value

  // Prepare Arrays

  const progressionBarSparklineValues = [];
  const progressionColorValues = [];

  const sprintSpaklineValues = [];
  const sprintColorValues = [];

  const jobSparklineValues = [];
  const jobColorValues = [];

  const storiesSequenceRanges = [];
  const storiesSparklineValues = [];

  // Fill Arrays

  // Main Bar
  const mainSpaklineValue = '=SPARKLINE(' + mainBarValueColumn + timeManagementMainBarRow + ';{"charttype","bar";"max",CEILING(' + mainBarTotalColumn + timeManagementMainBarRow + ');"color1","' + getHighlightcolor2() + '"})';

  mainBarSparklineRange.setValue(mainSpaklineValue);

  // Progression Bar
  firstCol = timeManagementCol;
  for (let index in progressionValues) {

    if (keys.includes(progressionValues[index])) {

      // Get Current Color Index
      const colorIndex = keys.indexOf(progressionValues[index]);

      // Make Color Row
      const rowColor = [colors[colorIndex], getHighlightcolor1(), middleGray, getHighlightcolor2(), colors[colorIndex], GetDarkModeTextColor()];

      // Get Formula
      const sparklineValue = '=SPARKLINE($' + progressionTotalColumn + (dataRow + parseInt(index)) + '-$' + progressionToDoColumn + (dataRow + parseInt(index)) + ';{"charttype","bar";"max",CEILING($' + progressionTotalColumn + (dataRow + parseInt(index)) + ');"color1","' + colors[colorIndex] + '"})';

      // Append Array
      progressionColorValues.push(rowColor);
      progressionBarSparklineValues.push([sparklineValue]);
    } else {
      // Append Array Empty
      progressionColorValues.push([null, null, null, null, null, null]);
      progressionBarSparklineValues.push([null]);
    }
  }

  // Job Bar
  firstCol = getColumnFromA1(jobColumn);
  for (let index in jobValues) {
    if (setupJobValues.includes(jobValues[index])) {

      // Get Color Index
      const colorIndex = setupJobValues.indexOf(jobValues[index]);

      // Make Color Row
      const rowColor = [setupJobColors[colorIndex], getHighlightcolor1(), middleGray, getHighlightcolor2(), setupJobColors[colorIndex], GetDarkModeTextColor()];

      // Get Formula
      const sparklineValue = '=SPARKLINE($' + jobTotalColumn + (dataRow + parseInt(index)) + '-$' + jobToDoColumn + (dataRow + parseInt(index)) + ';{"charttype","bar";"max",CEILING($' + jobTotalColumn + (dataRow + parseInt(index)) + ');"color1","' + setupJobColors[colorIndex] + '"})';

      // Append Array
      jobColorValues.push(rowColor);
      jobSparklineValues.push([sparklineValue]);
    }
  }

  //Sprint Bar
  firstCol = getColumnFromA1(sprintTeamColumn);
  for (let index in sprintTeamValues) {
    if (setupTeamValues.includes(sprintTeamValues[index])) {

      // Get Color Index
      const colorIndex = setupTeamValues.indexOf(sprintTeamValues[index]);

      // Make Color Row
      const rowColor = [setupTeamColors[colorIndex], GetDarkModeTextColor(), null];

      // Get Formula
      const sparklineValue = '=SPARKLINE(' + sprintDaysColumn + (dataRow + parseInt(index)) + ';{"charttype","bar";"max",CEILING(MAX(' + sprintDaysColumn + dataRow + ':' + sprintDaysColumn + '));"color1","' + setupTeamColors[colorIndex] + '"})';

      // Append Arrays
      sprintColorValues.push(rowColor);
      sprintSpaklineValues.push([sparklineValue]);
    }
  }

  //Stories Bar + Epic Days  

  //Unmerge Cells
  sheet.getRange(storiesTotalColumn + dataRow + ':' + storiesTotalColumn + sheet.getLastRow()).breakApart();

  let startRow = dataRow;
  for (let index in storiesSequence) {
    storiesSequenceRanges.push(storiesColumn + startRow + ':' + storiesBarColumn + (startRow + storiesSequence[index] - 1));
    const storiesMaxForula = '=SUM(' + storiesDayColumn + startRow + ':' + storiesDayColumn + (startRow + storiesSequence[index] - 1) + ')';
    const storiesTotalRanges = sheet.getRange(storiesTotalColumn + startRow + ':' + storiesTotalColumn + (startRow + storiesSequence[index] - 1));
    storiesTotalRanges.merge().setValue(storiesMaxForula);

    startRow = startRow + storiesSequence[index];
  }

  for (let index in storiesValues) {
    const sparklineValue = '=SPARKLINE(' + storiesDayColumn + (dataRow + parseInt(index)) + ';{"charttype","bar";"max",CEILING(MAX(' + storiesDayColumn + dataRow + ':' + storiesDayColumn + '));"color1","' + getHighlightcolor1() + '"})';
    storiesSparklineValues.push([sparklineValue]);
  }

  // Set Values

  progressionSparklineRange.setValues(progressionBarSparklineValues);
  progressionValuesRange.setFontColors(progressionColorValues);

  jobSparklineRange.setValues(jobSparklineValues);
  jobValueRange.setFontColors(jobColorValues);

  sprintSparklineRange.setValues(sprintSpaklineValues);
  sprintValueRange.setFontColors(sprintColorValues)

  storiesSparklineRange.setValues(storiesSparklineValues);

  SetTimeManagementConditionnalFormattingRules();
}

////////////////////////////
//     BUDGET SETUP       //__________________________________________________________________________________________________________________________________________
////////////////////////////

function GetProbabilityRange() {
  const columnA1 = getColumnFromName(budgetSetupSheet, probabilityName, budgetSetupRow, budgetSetupCol, budgetSetupWidth);
  const range = budgetSetupSheet.getRange(columnA1 + (budgetSetupRow + 1) + ':' + columnA1 + budgetSetupLength);
  return range;
}

function GetCostRange() {
  const columnA1 = getColumnFromName(budgetSetupSheet, costSourcesName, budgetSetupRow, budgetSetupCol, budgetSetupWidth);
  const range = budgetSetupSheet.getRange(columnA1 + (budgetSetupRow + 1) + ':' + columnA1 + budgetSetupLength);
  return range;
}

function GetIncomeRange() {
  const columnA1 = getColumnFromName(budgetSetupSheet, incomeSourcesName, budgetSetupRow, budgetSetupCol, budgetSetupWidth);
  const range = budgetSetupSheet.getRange(columnA1 + (budgetSetupRow + 1) + ':' + columnA1 + budgetSetupLength);
  return range;
}

function GetCostIncomeRange() {
  const columnA1 = getColumnFromName(budgetSetupSheet, parameterName, budgetSetupRow, budgetSetupCol, budgetSetupWidth);
  const range = budgetSetupSheet.getRange(columnA1 + (budgetSetupRow + 1) + ':' + columnA1 + (budgetSetupRow + 2));
  return range;
}

function CreateProbabilitySelection(targetSheet, column, row) {

  // Get All Probabilities
  const probabilities = GetValues(GetProbabilityRange());
  probabilities.shift(); // Get rid of the Certain Probability

  // Set the probabilities

  const tabsValues = [];

  for (let index in probabilities) {
    tabsValues.push([probabilityName, probabilities[index]]);
  }

  // Set the Values
  const targetRange = targetSheet.getRange(column + row + ':' + getColumnOffset(column, 1) + (row + probabilities.length - 1));
  targetRange.setValues(tabsValues);

  // Set the validation Rule

  const validationRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();

  const dataValidationRange = targetSheet.getRange(getColumnOffset(column, 2) + row + ':' + getColumnOffset(column, 2) + (row + probabilities.length - 1));
  dataValidationRange.setDataValidation(validationRule);

  // Merge the first Column
  targetSheet.getRange(column + row + ':' + column + (row + probabilities.length - 1)).merge();
}

////////////////////////////
//          COSTS         //__________________________________________________________________________________________________________________________________________
////////////////////////////

function GetSalaryRange() {
  const salaryColumn = getColumnFromName(costsSheet, salariesName, costsTabsRow - 1, costsCol, costsWidth);
  return costsSheet.getRange(salaryColumn + (costsTabsRow + 1) + ':' + getColumnOffset(salaryColumn, costsSalaryWidth - 1) + costsSheet.getLastRow());
}

function GetPonctualCostsRange() {
  const ponctualColumn = getColumnFromName(costsSheet, ponctualCostsName, costsTabsRow - 1, costsCol, costsWidth);

  return costsSheet.getRange(ponctualColumn + (costsTabsRow + 1) + ':' + getColumnOffset(ponctualColumn, costsPonctualWidth - 1) + costsSheet.getLastRow());
}

function GetFixCostsRange() {
  const fixColumn = getColumnFromName(costsSheet, fixCostsName, costsTabsRow - 1, costsCol, costsWidth);
  return costsSheet.getRange(fixColumn + (costsTabsRow + 1) + ':' + getColumnOffset(fixColumn, fixCostsWidth - 1) + costsSheet.getLastRow());
}

function GetVariableCostsRange() {
  const variableColumn = getColumnFromName(costsSheet, variableCostsName, costsTabsRow - 1, costsCol, costsWidth);
  return costsSheet.getRange(variableColumn + (costsTabsRow + 1) + ':' + getColumnOffset(variableColumn, costsSheet.getLastColumn() - getColumnFromA1(variableColumn)) + costsSheet.getLastRow());
}

////////////////////////////
//         INCOMES        //__________________________________________________________________________________________________________________________________________
////////////////////////////

function GetPonctualIncomesRange() {
  const ponctualColumn = getColumnFromName(incomesSheet, ponctualIncomesName, incomesTabsRow - 1, incomesCol, incomesWidth);
  return incomesSheet.getRange(ponctualColumn + (incomesTabsRow + 1) + ':' + getColumnOffset(ponctualColumn, incomesPonctualWidth - 1) + incomesSheet.getLastRow());
}

function GetFixIncomesRange() {
  const fixColumn = getColumnFromName(incomesSheet, fixIncomesName, incomesTabsRow - 1, incomesCol, incomesWidth);
  return incomesSheet.getRange(fixColumn + (incomesTabsRow + 1) + ':' + getColumnOffset(fixColumn, fixIncomesWidth - 1) + incomesSheet.getLastRow());
}

function GetVariableIncomesRange() {
  const variableColumn = getColumnFromName(incomesSheet, variableIncomesName, incomesTabsRow - 1, incomesCol, incomesWidth);
  return incomesSheet.getRange(variableColumn + (incomesTabsRow + 1) + ':' + getColumnOffset(variableColumn, incomesSheet.getLastColumn() - getColumnFromA1(variableColumn)) + incomesSheet.getLastRow());
}

////////////////////////////
//      BUDGET TIMELINE   //__________________________________________________________________________________________________________________________________________
////////////////////////////

function CalculateMonths(startingDate, nbMonths) {
  const date = new Date(startingDate);

  return Array.from({ length: nbMonths }, (_, k) => {
    const monthDate = new Date(date);
    monthDate.setUTCMonth(date.getUTCMonth() + k);
    return monthDate;
  });
}

function GetMonthsCoor(months) {
  return months.map(GetMonthCoor);
}

function GetMonthCoor(date) {
  const year = date.getFullYear();
  const month = date.getMonth();

  return year + ':' + month;
}

function GetMonthsIndex(startingDate, endingDate, months) {

  // Get the coordinates
  const startingDateCoor = GetMonthCoor(startingDate);
  const endingDateCoor = GetMonthCoor(endingDate);
  const monthsCoor = GetMonthsCoor(months);

  // Initialize Variable
  let startingDateIndex;
  let endingDateIndex;

  // Calculate Starting Date Index
  if (startingDate < months[0]) {
    startingDateIndex = 0; // If starting date before the start of the timeline
  } else {
    if (startingDate > months[months.length - 1]) {
      startingDateIndex = -1; // If starting date after the end of the timeline
    }
    else {
      startingDateIndex = monthsCoor.indexOf(startingDateCoor); // If starting date inside the timeline
    }
  }

  // Calculate Ending Date Index
  if (endingDate > months[months.length - 1]) {
    endingDateIndex = months.length - 1; // If ending date after the end of the timeline
  } else {
    if (endingDate < months[0]) {
      endingDateIndex = -1; // If ending date before the start of the timeline
    }
    else {
      endingDateIndex = monthsCoor.indexOf(endingDateCoor); // If ending date inside the timeline
    }
  }


  return [startingDateIndex, endingDateIndex];
}

function CreatebudgetTimeline() {
  const sheet = budgetTimelineSheet;
  const row = budgetTimelineRow;
  const dataRow = row + 1;
  const monthRow = row;
  const length = budgetTimelineLength;

  // Get Param Datas

  const startingDateRow = getRowFromName(sheet, budgetTimelineStartingDateName, row, getColumnFromA1(budgetTimelineTitleColumn), length);
  const nbMonthsRow = getRowFromName(sheet, budgetTimelineDurationName, row, getColumnFromA1(budgetTimelineTitleColumn), length);

  const nbMonths = sheet.getRange(budgetTimelineParam1Column + nbMonthsRow).getValue();
  const startingDate = new Date(sheet.getRange(budgetTimelineParam1Column + startingDateRow).getValue());

  const months = CalculateMonths(startingDate, nbMonths);

  // Get Columns & Rows

  const width = Math.max(sheet.getLastColumn(), budgetTimelineMinWidth);
  //  const lastColumn = getA1FromCol(width);
  const firstColumn = timelineFirstColumn;
  const firstCol = getColumnFromA1(firstColumn);

  const lastCol = firstCol + nbMonths;
  const lastColumn = getA1FromCol(lastCol);

  const secondColumn = getColumnOffset(firstColumn, 1);
  const seconCol = getColumnFromA1(secondColumn);

  const fundsRow = getRowFromName(sheet, fundsName, budgetTimelineRow, firstCol, length);
  const gainsRow = getRowFromName(sheet, gainsName, budgetTimelineRow, firstCol, length);
  const sparklineFundStartRow = fundsRow - budgetTimelineBarLength;
  const sparklineCurveStartRow = gainsRow - budgetTimelineCurveLength;
  const sparklineEndRow = gainsRow - 1;

  // Get Setup Data

  const setupTitlesColors = GetColors(GetCostIncomeRange());

  const setupCostsRange = GetCostRange();
  const setupIncomeRange = GetIncomeRange();

  const costsSources = setupCostsRange.getValues().filter(String);
  const incomesSources = setupIncomeRange.getValues().filter(String);

  const costColors = GetColors(setupCostsRange);
  const incomesColors = GetColors(setupIncomeRange);

  // First Column Values
  const firstColumnValues = [
    [monthsName],
    [costsName],
    ...costsSources,
    [incomesName],
    ...incomesSources,
    ...Array(budgetTimelineCurveLength).fill(['']),
    [gainsName],
    ...Array(budgetTimelineBarLength).fill(['']),
    [fundsName],
  ];

  const firstColumnValues_Clean = cleanData(firstColumnValues);

  // First Column Colors
  const colorIndex = [
    [''],
    setupTitlesColors[0],
    ...costColors,
    setupTitlesColors[1],
    ...incomesColors,
    ...Array(budgetTimelineCurveLength).fill([GetDarkModeColor()]),
    [getHighlightcolor2()],
    ...Array(budgetTimelineBarLength).fill([GetDarkModeColor()]),
    [getHighlightcolor1()],
  ];

  // Get Ranges

  const monthsRange = sheet.getRange(secondColumn + monthRow + ':' + getColumnOffset(secondColumn, nbMonths - 1) + monthRow);
  const firstColumnRange = sheet.getRange(firstColumn + monthRow + ':' + firstColumn + (monthRow + firstColumnValues.length - 1));
  const dataRange = sheet.getRange(secondColumn + dataRow + ':' + getColumnOffset(secondColumn, nbMonths - 1) + (dataRow + firstColumnValues.length));
  const planningMaxRange = sheet.getRange(firstColumn + monthRow + ':' + getA1FromCol(sheet.getMaxColumns()) + sheet.getMaxRows());

  // Parameter Datas
  const gradientRow = getRowFromName(sheet, budgetTimelineFundGradientName, row, getColumnFromA1(budgetTimelineTitleColumn), length);
  const fundRowRange = sheet.getRange(secondColumn + fundsRow + ':' + lastColumn + fundsRow);
  const gainRowRange = sheet.getRange(secondColumn + gainsRow + ':' + lastColumn + gainsRow);
  const idealFundCell = sheet.getRange(budgetTimelineParam2Column + gradientRow);
  const criticalFundCell = sheet.getRange(budgetTimelineParam2Column + (gradientRow + budgetTimelineGradientLength - 1));

  const fundMaxColor = idealFundCell.getBackground();
  const fundMaxValue = idealFundCell.getValue();
  const fundMinColor = criticalFundCell.getBackground();
  const fundMinValue = criticalFundCell.getValue();

  // Conditionnal Format Rules

  // Get Month Rules
  const monthRule = GetConditionnalFormattingRulesContain(GetMonthsRange(), monthsRange);

  // Get Cost and Income Rules
  const costRules = GetConditionnalFormattingRulesContain(GetCostRange(), firstColumnRange);
  const incomesRules = GetConditionnalFormattingRulesContain(GetIncomeRange(), firstColumnRange);
  const costIncomeRules = GetConditionnalFormattingRulesContain(GetCostIncomeRange(), firstColumnRange);

  // Get Gains and Funds Formatting Rules
  const gainFormatRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains(gainsName)
    .setBackground(getHighlightcolor2())
    .setFontColor('white')
    .setBold(true)
    .setRanges([firstColumnRange])
    .build();

  const fundFormatRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains(fundsName)
    .setBackground(getHighlightcolor1())
    .setFontColor('white')
    .setBold(true)
    .setRanges([firstColumnRange])
    .build();

  // Get Fund and Gain Gradient Formatting Rules
  const fundRowFormatRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue(fundMaxColor, SpreadsheetApp.InterpolationType.NUMBER, fundMaxValue)
    .setGradientMinpointWithValue(fundMinColor, SpreadsheetApp.InterpolationType.NUMBER, fundMinValue)
    .setRanges([fundRowRange])
    .build();

  const gainRowFormatRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue(fundMaxColor, SpreadsheetApp.InterpolationType.NUMBER, fundMinValue)
    .setGradientMinpointWithValue(fundMinColor, SpreadsheetApp.InterpolationType.NUMBER, -fundMinValue)
    .setRanges([gainRowRange])
    .build();

  // Get Data Conditionnal Format Rules
  const dataRules = [];
  const exceptions = [costsName, incomesName, gainsName, fundsName];

  for (let index in firstColumnValues_Clean) {
    const formula = '=$' + firstColumn + dataRow + '="' + firstColumnValues_Clean[index] + '"';

    if (exceptions.includes(firstColumnValues_Clean[index])) {
      const dataFormatRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(formula)
        .setBackground(lightenDarkenColor(colorIndex[index], 0.2))
        .setRanges([dataRange])
        .build();
      dataRules.push(dataFormatRule);
    }
    else {
      const dataFormatRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(formula)
        .setBackground(lightenDarkenColor(colorIndex[index], 0.9))
        .setRanges([dataRange])
        .build();
      dataRules.push(dataFormatRule);
    };
  };

  // Create Conditionnal Rules
  const rules = [
    fundRowFormatRule,
    gainRowFormatRule,
    ...costRules,
    ...incomesRules,
    ...costIncomeRules,
    ...monthRule,
    gainFormatRule,
    fundFormatRule,
    ...dataRules,
  ];

  ///// CREATE TIMELINE ///// _________________________________________________________________________________________________

  // Clear Datas
  
  planningMaxRange.clearDataValidations();
  planningMaxRange.clearContent();
  planningMaxRange.breakApart();

  // Set Axis
  monthsRange.setValues([months]);
  firstColumnRange.setValues(firstColumnValues);


  // Set Row Heights & Column Widths
  sheet.setRowHeights(dataRow, firstColumnValues.length, 50);
  sheet.setColumnWidths(seconCol, nbMonths, 120);

  sheet.setConditionalFormatRules(rules);

  // Merge the Cells
  for (let k = 0; k < (nbMonths + 1); k++) {
    const barColumn = getColumnOffset(firstColumn, k);
    budgetTimelineSheet.getRange(barColumn + sparklineFundStartRow + ':' + barColumn + (sparklineFundStartRow + budgetTimelineBarLength - 1)).merge();
  }
  budgetTimelineSheet.getRange(secondColumn + sparklineCurveStartRow + ':' + lastColumn + sparklineEndRow).merge();
  budgetTimelineSheet.getRange(firstColumn + sparklineCurveStartRow + ':' + firstColumn + sparklineEndRow).merge();

  ///// SET VALUES ///// _________________________________________________________________________________________________

  CalulateBudgetTimelineValues();
  SetbudgetTimelineDarkMode();
  SetbudgetTimelineHighlights();
}

function CalulateBudgetTimelineValues() {
  const sheet = budgetTimelineSheet;
  const row = budgetTimelineRow;
  const dataRow = row + 1;
  const monthRow = row;
  const length = budgetTimelineLength;

  // Get Columns & Rows

  const width = Math.max(sheet.getLastColumn(), budgetTimelineMinWidth);
  const lastColumn = getA1FromCol(width);
  const firstColumn = getColumnFromName(sheet, monthsName, monthRow, 1, width);
  const firstCol = getColumnFromA1(firstColumn);
  const secondColumn = getColumnOffset(firstColumn, 1);

  const startingDateRow = getRowFromName(sheet, budgetTimelineStartingDateName, row, getColumnFromA1(budgetTimelineTitleColumn), length);
  const nbMonthsRow = getRowFromName(sheet, budgetTimelineDurationName, row, getColumnFromA1(budgetTimelineTitleColumn), length);

  const fundsRow = getRowFromName(sheet, fundsName, budgetTimelineRow, firstCol, length);
  const gainsRow = getRowFromName(sheet, gainsName, budgetTimelineRow, firstCol, length);
  const costsRow = getRowFromName(sheet, costsName, budgetTimelineRow, firstCol, length);
  const incomesRow = getRowFromName(sheet, incomesName, budgetTimelineRow, firstCol, length);
  const initialFundRow = getRowFromName(budgetTimelineSheet, budgetTimelineInitialFundName, row, getColumnFromA1(budgetTimelineTitleColumn), length);

  const sparklineFundStartRow = fundsRow - budgetTimelineBarLength;
  const sparklineCurveStartRow = gainsRow - budgetTimelineCurveLength;

  // Sheet Data

  const nbMonths = sheet.getRange(budgetTimelineParam1Column + nbMonthsRow).getValue();
  const startingDate = new Date(sheet.getRange(budgetTimelineParam1Column + startingDateRow).getValue());

  const months = CalculateMonths(startingDate, nbMonths);

  const probabilityWithlist = GetTimelineProbabilityWithlist();

  // Get Setup Data

  const setupCostsRange = GetCostRange()
  const setupIncomeRange = GetIncomeRange()

  const costsSources = setupCostsRange.getValues().filter(String);
  const incomesSources = setupIncomeRange.getValues().filter(String);


  // First Column Values
  const firstColumnValues = [
    [costsName],
    ...costsSources,
    [incomesName],
    ...incomesSources,
    ...Array(budgetTimelineCurveLength).fill(['']),
    [gainsName],
    ...Array(budgetTimelineBarLength).fill(['']),
    [fundsName],
  ];

  const firstColumnValues_Clean = cleanData(firstColumnValues);

  // GET RANGES 

  // Get Ranges
  const dataRange = sheet.getRange(secondColumn + dataRow + ':' + getColumnOffset(secondColumn, nbMonths - 1) + (dataRow + firstColumnValues.length));

  //Preparing the Array 
  let budgetTimelineValues = Array.from({ length: dataRange.getValues().length }, () => {
    return Array(dataRange.getValues()[0].length).fill(0);
  });

  // Salaries
  budgetTimelineValues = AddSalaryToTimelineData(budgetTimelineValues, firstColumnValues_Clean, months);

  // Ponctual Costs
  budgetTimelineValues = AddPonctualCostsToTimelineData(budgetTimelineValues, firstColumnValues_Clean, months, probabilityWithlist);

  // Fix Costs
  budgetTimelineValues = AddFixCostsToTimelineData(budgetTimelineValues, firstColumnValues_Clean, months, probabilityWithlist);

  // Variable Costs
  budgetTimelineValues = AddVariableCostsToTimelineData(budgetTimelineValues, firstColumnValues_Clean, months, probabilityWithlist);

  // Ponctual Incomes
  budgetTimelineValues = AddPonctualIncomesToTimelineData(budgetTimelineValues, firstColumnValues_Clean, months, probabilityWithlist);

  // Fix Incomes
  budgetTimelineValues = AddFixIncomesToTimelineData(budgetTimelineValues, firstColumnValues_Clean, months, probabilityWithlist);

  // Variable Incomes
  budgetTimelineValues = AddVariableIncomesToTimelineData(budgetTimelineValues, firstColumnValues_Clean, months, probabilityWithlist);

  // Replace 0 with ''
  budgetTimelineValues = budgetTimelineValues.map(row => row.map(cell => cell === 0 ? '' : cell));

  // Prepare Formulas
  const costsFormulas = [];
  const incomesFormulas = [];
  const gainsFormulas = [];
  const fundsForumlas = [];
  const sparklineFormulas = [];

  const initialFundsFormula = '=SPARKLINE(' + budgetTimelineParam1Column + initialFundRow + ',{"charttype","column";"ymax",CEILING(MAX(' + firstColumn + fundsRow + ':' + lastColumn + fundsRow + '));"ymin",FLOOR(MIN(' + firstColumn + fundsRow + ':' + lastColumn + fundsRow + '));"color","' + getHighlightcolor1() + '";"negcolor","' + 'red' + '"})';

  for (let index in months) {
    const currentColumn = getColumnOffset(secondColumn, parseInt(index));

    const costsFormula = '=SUM(' + currentColumn + (costsRow + 1) + ':' + currentColumn + (costsRow + costsSources.length) + ')';

    const incomesFormula = '=SUM(' + currentColumn + (incomesRow + 1) + ':' + currentColumn + (incomesRow + incomesSources.length) + ')';

    const gainsFormula = '=' + currentColumn + incomesRow + '-' + currentColumn + costsRow + '';

    const sparklineFormula = '=SPARKLINE(' + currentColumn + fundsRow + ',{"charttype","column";"ymax",CEILING(MAX(' + firstColumn + fundsRow + ':' + lastColumn + fundsRow + '));"ymin",FLOOR(MIN(' + firstColumn + fundsRow + ':' + lastColumn + fundsRow + '));"color","' + getHighlightcolor1() + '";"negcolor","' + 'red' + '"})';

    const fundsFormula = '=' + budgetTimelineParam1Column + initialFundRow + '+ SUM(' + secondColumn + gainsRow + ':' + currentColumn + gainsRow + ')';

    costsFormulas.push(costsFormula);
    incomesFormulas.push(incomesFormula);
    gainsFormulas.push(gainsFormula);
    sparklineFormulas.push(sparklineFormula);
    fundsForumlas.push(fundsFormula);
  }

  const sparklineCurveFormula = '=SPARKLINE(' + secondColumn + gainsRow + ':' + lastColumn + gainsRow + ',{"color","' + getHighlightcolor2() + '";"ymax",MAX(ABS(CEILING(MIN(' + secondColumn + gainsRow + ':' + lastColumn + gainsRow + ')));CEILING(MAX(' + secondColumn + gainsRow + ':' + lastColumn + gainsRow + ')));"ymin",-MAX(ABS(CEILING(MIN(' + secondColumn + gainsRow + ':' + lastColumn + gainsRow + ')));CEILING(MAX(' + secondColumn + gainsRow + ':' + lastColumn + gainsRow + ')))})';

  const sparklineThresholdFormula = '=SPARKLINE(' + budgetTimelineParam1Column + initialFundRow + ',{"charttype","column";"ymax",10000000;"ymin",-10000000;"color","red"})';

  // Put Sparklines in Datas
  budgetTimelineValues[costsRow - dataRow] = costsFormulas;
  budgetTimelineValues[incomesRow - dataRow] = incomesFormulas;
  budgetTimelineValues[gainsRow - dataRow] = gainsFormulas;
  budgetTimelineValues[fundsRow - dataRow] = fundsForumlas;
  budgetTimelineValues[sparklineFundStartRow - dataRow] = sparklineFormulas;
  budgetTimelineValues[sparklineCurveStartRow - dataRow][0] = sparklineCurveFormula;

  // Set Datas
  dataRange.setValues(budgetTimelineValues);
  sheet.getRange(firstColumn + (gainsRow - budgetTimelineCurveLength)).setValue(sparklineThresholdFormula);
  sheet.getRange(firstColumn + (gainsRow + 1)).setValue(initialFundsFormula);
}

function GetTimelineProbabilityWithlist() {
  // Get Row
  const probabilityRow = getRowFromName(budgetTimelineSheet, probabilityName, budgetTimelineRow, getColumnFromA1(budgetTimelineTitleColumn), budgetTimelineLength);

  // Get Setup Datas
  const probabilityList = GetValues(GetProbabilityRange());
  const probabilityLength = probabilityList.length - 1;

  // Get Sheet Datas
  const timelineProbaRange = budgetTimelineSheet.getRange(budgetTimelineParam2Column + probabilityRow + ':' + budgetTimelineParam2Column + (probabilityRow + probabilityLength - 1));
  const timelineProbaValues = timelineProbaRange.getValues();

  // Handle the Certain probability
  probabilityWithlist = [probabilityList[0]];
  probabilityList.shift();

  // Check if probability selected
  for (let index in timelineProbaValues) {
    if (timelineProbaValues[index][0] == true) probabilityWithlist.push(probabilityList[index]);
  }

  // Return the withlist
  return probabilityWithlist.filter(String);
}

function RefreshTimelineMenu() {
  const sheet = budgetTimelineSheet;
  const firstRow = 9;
  let currentRow = firstRow;

  // Delete previous menu
  const menuRange = sheet.getRange(budgetTimelineTitleColumn + firstRow + ':' + budgetTimelineParam3Column + sheet.getLastRow())
  menuRange.clearContent();
  menuRange.clearDataValidations();
  menuRange.setBorder(true, true, true, true, true, true, GetDarkModeColor(), SpreadsheetApp.BorderStyle.SOLID);
  menuRange.breakApart();

  // Create Probability Selector
  CreateProbabilitySelection(sheet, budgetTimelineTitleColumn, firstRow);

  // Create Importance Scope
  const probabilities = GetValues(GetProbabilityRange());
  probabilities.shift()

  currentRow += probabilities.length + 1;

  CreateImportanceScope(sheet, budgetTimelineTitleColumn, currentRow);
  SetbudgetTimelineHighlights();
}

function AddSalaryToTimelineData(data, dataFirstColumn, months) {
  const sheet = costsSheet;
  const row = costsTabsRow;
  const col = costsCol;
  const width = costsSalaryWidth;

  // Get Salaries Datas
  const salaryDataValues = GetSalaryRange().getValues();

  // Get Indexes 
  const salaryCol = getColumnFromA1(getColumnFromName(sheet, salariesName, row - 1, col, costsWidth));
  const salaryTeamIndex = getColumnFromA1(getColumnFromName(sheet, salariesTeamName, row, salaryCol, width)) - salaryCol;
  const salaryContractIndex = getColumnFromA1(getColumnFromName(sheet, permanentContractName, row, salaryCol, width)) - salaryCol;
  const salarayStartingDateIndex = getColumnFromA1(getColumnFromName(sheet, startingDateName, row, salaryCol, width)) - salaryCol;
  const salaryEndingDateIndex = getColumnFromA1(getColumnFromName(sheet, endingDateName, row, salaryCol, width)) - salaryCol;
  const salaryCostIndex = getColumnFromA1(getColumnFromName(sheet, salariesMonthlyCostName, row, salaryCol, width)) - salaryCol;
  const timelineRowIndex = dataFirstColumn.indexOf(salariesName);

  // For each Cost
  for (const salaryData of salaryDataValues) {

    const costname = salaryData[salaryTeamIndex];

    // Check if Cost exist
    if (costname !== '') {

      // Get Cost datas
      const costValue = parseFloat(salaryData[salaryCostIndex]);
      const permanentContract = salaryData[salaryContractIndex];
      const costStartingDate = salaryData[salarayStartingDateIndex];
      const costEndingDate = salaryData[salaryEndingDateIndex];

      // Get Months Index
      let monthsIndex;

      // Check Contract
      if (permanentContract) {
        monthsIndex = [0, months.length - 1];
      } else {
        monthsIndex = GetMonthsIndex(costStartingDate, costEndingDate, months);
      }

      // for each month in Timeline
      for (let i = monthsIndex[0]; i <= monthsIndex[1]; i++) {
        // Add Cost to Timeline
        data[timelineRowIndex][i] += costValue;
      }
    }
  }

  // Return Timeline Data
  return data;
}

function AddPonctualCostsToTimelineData(data, dataFirstColumn, months, probabilityWithlist) {
  const sheet = costsSheet;
  const row = costsTabsRow;
  const col = costsCol;
  const width = costsPonctualWidth;

  // Get Datas
  const monthsCoor = GetMonthsCoor(months);
  const ponctualCostsDataValues = GetPonctualCostsRange().getValues();

  // Get Indexes
  const ponctualCostsCol = getColumnFromA1(getColumnFromName(sheet, ponctualCostsName, row - 1, col, costsWidth));
  const ponctualCostNameIndex = getColumnFromA1(getColumnFromName(sheet, costNameName, row, ponctualCostsCol, width)) - ponctualCostsCol;
  const ponctualCostIndex = getColumnFromA1(getColumnFromName(sheet, costValueName, row, ponctualCostsCol, width)) - ponctualCostsCol;
  const ponctualCostsDateIndex = getColumnFromA1(getColumnFromName(sheet, dateName, row, ponctualCostsCol, width)) - ponctualCostsCol;
  const ponctualCostProbabilityIndex = getColumnFromA1(getColumnFromName(sheet, probabilityName, row, ponctualCostsCol, width)) - ponctualCostsCol;
  const ponctualCostSourceIndex = getColumnFromA1(getColumnFromName(sheet, sourcesName, row, ponctualCostsCol, width)) - ponctualCostsCol;

  // For each Cost
  for (let cost of ponctualCostsDataValues) {
    const costname = cost[ponctualCostNameIndex];

    // Check if Cost exist
    if (costname !== '') {
      const costProbability = cost[ponctualCostProbabilityIndex];

      // Check Probability Withlist
      if (probabilityWithlist.includes(costProbability)) {

        // Get Cost Datas
        const costValue = parseFloat(cost[ponctualCostIndex]);
        const costDate = cost[ponctualCostsDateIndex];
        const costSource = cost[ponctualCostSourceIndex];

        // Get Month Coordinate
        const costDateCoor = GetMonthCoor(costDate);

        // Get Timeline Index
        const timelineMonthIndex = monthsCoor.indexOf(costDateCoor);
        const timelineRowIndex = dataFirstColumn.indexOf(costSource);

        // Add Cost to Timeline
        data[timelineRowIndex][timelineMonthIndex] += costValue;
      }
    }
  }

  // Return Timeline Data
  return data;
}

function AddFixCostsToTimelineData(data, dataFirstColumn, months, probabilityWithlist) {
  const sheet = costsSheet;
  const row = costsTabsRow;
  const col = costsCol;
  const width = fixCostsWidth;

  // Get cost Datas
  const fixCostsDataValues = GetFixCostsRange().getValues();

  // Get Indexes
  const fixCostsCol = getColumnFromA1(getColumnFromName(sheet, fixCostsName, row - 1, col, costsWidth));
  const fixCostNameIndex = getColumnFromA1(getColumnFromName(sheet, costNameName, row, fixCostsCol, width)) - fixCostsCol;
  const fixCostIndex = getColumnFromA1(getColumnFromName(sheet, costValueName, row, fixCostsCol, width)) - fixCostsCol;
  const fixCostStartingDateIndex = getColumnFromA1(getColumnFromName(sheet, startingDateName, row, fixCostsCol, width)) - fixCostsCol;
  const fixCostEndingDateIndex = getColumnFromA1(getColumnFromName(sheet, endingDateName, row, fixCostsCol, width)) - fixCostsCol;
  const fixCostProbabilityIndex = getColumnFromA1(getColumnFromName(sheet, probabilityName, row, fixCostsCol, width)) - fixCostsCol;
  const fixCostSourceIndex = getColumnFromA1(getColumnFromName(sheet, sourcesName, row, fixCostsCol, width)) - fixCostsCol;

  // For each Cost
  for (let cost of fixCostsDataValues) {
    const costname = cost[fixCostNameIndex];

    // Check if Cost exist
    if (costname !== '') {
      const costProbability = cost[fixCostProbabilityIndex];

      // Check Probabilty Withlist
      if (probabilityWithlist.includes(costProbability)) {

        // Get Cost datas
        const costValue = parseFloat(cost[fixCostIndex]);
        const costSource = cost[fixCostSourceIndex];
        const costStartingDate = cost[fixCostStartingDateIndex];
        const costEndingDate = cost[fixCostEndingDateIndex];

        // Get Month Index
        const monthsIndex = GetMonthsIndex(costStartingDate, costEndingDate, months);
        
        // Get Timeline Y Index
        const timelineRowIndex = dataFirstColumn.indexOf(costSource);

        // For each months in Timeline
        for (let timelineMonthIndex = monthsIndex[0]; timelineMonthIndex <= monthsIndex[1]; timelineMonthIndex++) {

          // Add Cost to Timeline
          data[timelineRowIndex][timelineMonthIndex] += costValue;
        }
      }
    }
  }

  // Return Timeline Datas
  return data;
}

function AddVariableCostsToTimelineData(data, dataFirstColumn, months, probabilityWithlist) {
  const sheet = costsSheet;
  const row = costsTabsRow;
  const col = costsCol;
  const length = costsLength;

  // Get Cost Datas
  const monthsCoor = GetMonthsCoor(months);
  const variableCostsDataValues = transposeArray(GetVariableCostsRange().getValues());

  // Get Indexes
  const variableCostsCol = getColumnFromA1(getColumnFromName(sheet, variableCostsName, row - 1, col, costsWidth));
  const variableCostsNameIndex = getRowFromName(sheet, costNameName, row, variableCostsCol, length) - (row + 1);
  const variableCostsProbabilityIndex = getRowFromName(sheet, probabilityName, row, variableCostsCol, length) - (row + 1);
  const variableCostsSourceIndex = getRowFromName(sheet, sourcesName, row, variableCostsCol, length) - (row + 1);
  const variableCostStartIndex = Math.max(variableCostsProbabilityIndex, variableCostsSourceIndex) + 1;

  // Get months of the Cost
  const variableCostsMonths = variableCostsDataValues.shift()

  // For each Cost
  for (let cost of variableCostsDataValues) {
    const costName = cost[variableCostsNameIndex];

    // Check if Cost exist
    if (costName !== '') {
      const costProbability = cost[variableCostsProbabilityIndex];

      // Check Probability Withlist
      if (probabilityWithlist.includes(costProbability)) {

        // Get Cost datas
        const costSource = cost[variableCostsSourceIndex];

        // Get Timeline Y Index
        const timelineRowIndex = dataFirstColumn.indexOf(costSource);

        // For each Cost's month
        for (let index = variableCostStartIndex; index < variableCostsMonths.length; index++) {

          // Get Cost value
          const costValue = cost[index];
          if (costValue !== '') {
            
            // Get  Cost Date
            const costDate = variableCostsMonths[index];

            // Get Timeline X Index
            const timelineMonthIndex = monthsCoor.indexOf(GetMonthCoor(costDate));

            // Add Cost to Timeline
            data[timelineRowIndex][timelineMonthIndex] += costValue;
          }
        }
      }
    }
  }

  // Return Timeline Datas
  return data;
}

function AddPonctualIncomesToTimelineData(data, dataFirstColumn, months, probabilityWithlist) {
  const sheet = incomesSheet;
  const row = incomesTabsRow;
  const col = incomesCol;
  const width = incomesPonctualWidth;

  // Get Income Datas
  const monthsCoor = GetMonthsCoor(months);
  const ponctualIncomesDataValues = GetPonctualIncomesRange().getValues();

  // Get Indexes
  const ponctualIncomesCol = getColumnFromA1(getColumnFromName(sheet, ponctualIncomesName, row - 1, col, costsWidth));
  const ponctualIncomesNameIndex = getColumnFromA1(getColumnFromName(sheet, incomeNameName, row, ponctualIncomesCol, width)) - ponctualIncomesCol;
  const ponctualIncomesIndex = getColumnFromA1(getColumnFromName(sheet, incomeValueName, row, ponctualIncomesCol, width)) - ponctualIncomesCol;
  const ponctualIncomesDateIndex = getColumnFromA1(getColumnFromName(sheet, dateName, row, ponctualIncomesCol, width)) - ponctualIncomesCol;
  const ponctualIncomesProbabilityIndex = getColumnFromA1(getColumnFromName(sheet, probabilityName, row, ponctualIncomesCol, width)) - ponctualIncomesCol;
  const ponctualIncomesSourceIndex = getColumnFromA1(getColumnFromName(sheet, sourcesName, row, ponctualIncomesCol, width)) - ponctualIncomesCol;

  // For each Income
  for (let income of ponctualIncomesDataValues) {
    const incomename = income[ponctualIncomesNameIndex];

    // Check if Income Exist
    if (incomename !== '') {
      const incomeProbability = income[ponctualIncomesProbabilityIndex];

      // Check probability withlist
      if (probabilityWithlist.includes(incomeProbability)) {

        // Get Income datas
        const incomeValue = parseFloat(income[ponctualIncomesIndex]);
        const incomeDate = income[ponctualIncomesDateIndex];
        const incomeSource = income[ponctualIncomesSourceIndex];

        // Get Income Coordinate
        const incomeDateCoor = GetMonthCoor(incomeDate);

        // Get Timeline Index
        const timelineMonthIndex = monthsCoor.indexOf(incomeDateCoor);
        const timelineRowIndex = dataFirstColumn.indexOf(incomeSource);

        // Add Income to Timeline
        data[timelineRowIndex][timelineMonthIndex] += incomeValue;

      }
    }
  }

  // Return Timeline Datas
  return data;
}

function AddFixIncomesToTimelineData(data, dataFirstColumn, months, probabilityWithlist) {
  const sheet = incomesSheet;
  const row = incomesTabsRow;
  const col = incomesCol;
  const width = fixIncomesWidth;

  // get Income Datas
  const fixIncomesDataValues = GetFixIncomesRange().getValues();

  const fixIncomesCol = getColumnFromA1(getColumnFromName(sheet, fixIncomesName, row - 1, col, incomesWidth));
  const fixIncomesNameIndex = getColumnFromA1(getColumnFromName(sheet, incomeNameName, row, fixIncomesCol, width)) - fixIncomesCol;
  const fixIncomesIndex = getColumnFromA1(getColumnFromName(sheet, incomeValueName, row, fixIncomesCol, width)) - fixIncomesCol;
  const fixIncomesStartingDateIndex = getColumnFromA1(getColumnFromName(sheet, startingDateName, row, fixIncomesCol, width)) - fixIncomesCol;
  const fixIncomesEndingDateIndex = getColumnFromA1(getColumnFromName(sheet, endingDateName, row, fixIncomesCol, width)) - fixIncomesCol;
  const fixIncomesProbabilityIndex = getColumnFromA1(getColumnFromName(sheet, probabilityName, row, fixIncomesCol, width)) - fixIncomesCol;
  const fixIncomesSourceIndex = getColumnFromA1(getColumnFromName(sheet, sourcesName, row, fixIncomesCol, width)) - fixIncomesCol;

  // For each Income
  for (let income of fixIncomesDataValues) {
    const incomename = income[fixIncomesNameIndex];

    // Check if Income exist
    if (incomename !== '') {
      const incomeProbability = income[fixIncomesProbabilityIndex];

      // Check probability withlist
      if (probabilityWithlist.includes(incomeProbability)) {

        // Get Income datas
        const incomeValue = parseFloat(income[fixIncomesIndex]);
        const incomeSource = income[fixIncomesSourceIndex];
        const incomeStartingDate = income[fixIncomesStartingDateIndex];
        const incomeEndingDate = income[fixIncomesEndingDateIndex];

        // Get Month Index
        const monthsIndex = GetMonthsIndex(incomeStartingDate, incomeEndingDate, months);

        // Get Timeline Y Index
        const timelineRowIndex = dataFirstColumn.indexOf(incomeSource);

        for (let timelineMonthIndex = monthsIndex[0]; timelineMonthIndex <= monthsIndex[1]; timelineMonthIndex++) {
          // Add cost to Timeline 
          data[timelineRowIndex][timelineMonthIndex] += incomeValue;
        }
      }
    }
  }

  // Return Timeline Datas
  return data;
}

function AddVariableIncomesToTimelineData(data, dataFirstColumn, months, probabilityWithlist) {
  const sheet = incomesSheet;
  const row = incomesTabsRow;
  const col = incomesCol;
  const length = incomesLength;

  // Get Incomes Datas
  const monthsCoor = GetMonthsCoor(months);
  const variableIncomesDataValues = transposeArray(GetVariableIncomesRange().getValues());

  // Get Indexes
  const variableIncomesCol = getColumnFromA1(getColumnFromName(sheet, variableIncomesName, row - 1, col, incomesWidth));
  const variableIncomesNameIndex = getRowFromName(sheet, incomeNameName, row, variableIncomesCol, length) - (row + 1);
  const variableIncomesProbabilityIndex = getRowFromName(sheet, probabilityName, row, variableIncomesCol, length) - (row + 1);
  const variableIncomesSourceIndex = getRowFromName(sheet, sourcesName, row, variableIncomesCol, length) - (row + 1);
  const variableIncomesStartIndex = Math.max(variableIncomesProbabilityIndex, variableIncomesSourceIndex) + 1;

  // Get Income Months
  const variableIncomesMonths = variableIncomesDataValues.shift();

  // For Each Income
  for (let income of variableIncomesDataValues) {
    const incomeName = income[variableIncomesNameIndex];

    // Check if Income exist
    if (incomeName !== '') {
      const incomeProbability = income[variableIncomesProbabilityIndex];

      // Check Probabilities Withlist
      if (probabilityWithlist.includes(incomeProbability)) {

        // Get Income datas
        const incomeSource = income[variableIncomesSourceIndex];

        // Get Timeline Y Index
        const timelineRowIndex = dataFirstColumn.indexOf(incomeSource);

        // For each cost's month
        for (let index = variableIncomesStartIndex; index < variableIncomesMonths.length; index++) {

          // Get Income value
          const incomeValue = income[index];
          if (incomeValue !== '') {

            // Get Cost Date
            const incomeDate = variableIncomesMonths[index];

            // Get Timeline X Index
            const timelineMonthIndex = monthsCoor.indexOf(GetMonthCoor(incomeDate));

            // Add Income to Timeline
            data[timelineRowIndex][timelineMonthIndex] += incomeValue;
          }
        }
      }
    }
  }

  // Return Timeline Datas
  return data;
}

////////////////////////////
//         BUDGETS         //__________________________________________________________________________________________________________________________________________
////////////////////////////

function RefreshBudgetMenus(){
  RefreshTimelineMenu();
  RefreshAnnualBudgetMenu();
  CreateProductionBudgetMenu();
  CreateCustomBudgetMenu();
}

function GetBudgetValues(startDate, endDate, withlist) {

  //// COSTS /////
  const salaryBudgetRows = CalculateSalaryBudgetRows(startDate, endDate);

  const ponctualCostsBudgetRows = CalculatePonctualBudgetRows(startDate, endDate, withlist);

  const fixCostsBudgetRows = CalculateFixCostsBudgetRows(startDate, endDate, withlist);

  const variableCostsBudgetRows = CalculateVariableCostsBudgetRows(startDate, endDate, withlist);

  const costsBudgetRows = [
    ...salaryBudgetRows,
    ...ponctualCostsBudgetRows,
    ...fixCostsBudgetRows,
    ...variableCostsBudgetRows
  ];

  // Create Costs Module
  const costsValues = CreateModules(costsBudgetRows, true);


  //// INCOMES /////
  const ponctualIncomesBudgetRows = CalculatePonctualIncomesBudgetRows(startDate, endDate, withlist);

  const fixIncomesBudgetRows = CalculateFixIncomesBudgetRows(startDate, endDate, withlist);

  const variableIncomesBudgetRows = CalculateVariableIncomesBudgetRows(startDate, endDate, withlist);

  const incomesBudgetRows = [
    ...ponctualIncomesBudgetRows,
    ...fixIncomesBudgetRows,
    ...variableIncomesBudgetRows
  ];

  // Create Incomes Module
  const incomesValues = CreateModules(incomesBudgetRows, false);

  // Create Budget from Modules
  const budgetValues = WrapBudgetValues(costsValues, incomesValues);

  // Return Budget
  return [budgetValues, [costsValues[1], incomesValues[1]]];
}

function GetNbMonthFromDates(startDate, endDate) {
  var months;
  months = (endDate.getFullYear() - startDate.getFullYear()) * 12;
  months -= startDate.getMonth();
  months += endDate.getMonth();
  return months <= 0 ? 0 : months;
}

function CalculateSalaryBudgetRows(startDate, endDate) {
  const sheet = costsSheet;
  const row = costsTabsRow;
  const col = costsCol;
  const width = costsSalaryWidth;

  // Inititalize Rows
  const salaryRows = [];

  // Get Cost Datas
  const salaryDatas = GetSalaryRange().getValues();

  // Get Indexes
  const salaryCol = getColumnFromA1(getColumnFromName(sheet, salariesName, row - 1, col, costsWidth));
  const salaryTeamIndex = getColumnFromA1(getColumnFromName(sheet, salariesTeamName, row, salaryCol, width)) - salaryCol;
  const salaryContractIndex = getColumnFromA1(getColumnFromName(sheet, permanentContractName, row, salaryCol, width)) - salaryCol;
  const salarayStartingDateIndex = getColumnFromA1(getColumnFromName(sheet, startingDateName, row, salaryCol, width)) - salaryCol;
  const salaryEndingDateIndex = getColumnFromA1(getColumnFromName(sheet, endingDateName, row, salaryCol, width)) - salaryCol;
  const salaryCostIndex = getColumnFromA1(getColumnFromName(sheet, salariesMonthlyCostName, row, salaryCol, width)) - salaryCol;
  let nbMonths;

  // For each Cost
  for (const salaryData of salaryDatas) {
    const costname = salaryData[salaryTeamIndex];

    // Check if Cost exist
    if (costname !== '') {

      // Get Cost datas
      const costValue = parseFloat(salaryData[salaryCostIndex]);
      const permanentContract = salaryData[salaryContractIndex];
      let costStartingDate = salaryData[salarayStartingDateIndex];
      let costEndingDate = salaryData[salaryEndingDateIndex];

      // Check contract
      if (permanentContract) {
        nbMonths = GetNbMonthFromDates(startDate, endDate);
      } else {
        // Check if Cost Date into budget
        if (costStartingDate <= endDate && costEndingDate >= startDate) {

          if (costStartingDate <= startDate) costStartingDate = startDate; // If Starting date before Budget
          if (costEndingDate >= endDate) costEndingDate = endDate; // If Ending date after Budget

          nbMonths = GetNbMonthFromDates(costStartingDate, costEndingDate);
        }
        else {
          nbMonths = 0;
        }
      };
      // Add Cost Row
      if (nbMonths * costValue != 0) salaryRows.push([salariesName, [costname, '', '', nbMonths * costValue]]);
    }
  }

  // Return Costs Rows
  return salaryRows;
}

function CalculatePonctualBudgetRows(startDate, endDate, probabilityWithlist) {
  const sheet = costsSheet;
  const row = costsTabsRow;
  const col = costsCol;
  const width = costsPonctualWidth;

  // Get Costs Datas
  const ponctualCostsDataValues = GetPonctualCostsRange().getValues();

  // Initialize Rows
  ponctualCostsBudgetRows = [];

  // Get Indexes
  const ponctualCostsCol = getColumnFromA1(getColumnFromName(sheet, ponctualCostsName, row - 1, col, costsWidth));
  const ponctualCostNameIndex = getColumnFromA1(getColumnFromName(sheet, costNameName, row, ponctualCostsCol, width)) - ponctualCostsCol;
  const ponctualCostIndex = getColumnFromA1(getColumnFromName(sheet, costValueName, row, ponctualCostsCol, width)) - ponctualCostsCol;
  const ponctualCostsDateIndex = getColumnFromA1(getColumnFromName(sheet, dateName, row, ponctualCostsCol, width)) - ponctualCostsCol;
  const ponctualCostProbabilityIndex = getColumnFromA1(getColumnFromName(sheet, probabilityName, row, ponctualCostsCol, width)) - ponctualCostsCol;
  const ponctualCostSourceIndex = getColumnFromA1(getColumnFromName(sheet, sourcesName, row, ponctualCostsCol, width)) - ponctualCostsCol;

  // For each Cost
  for (let cost of ponctualCostsDataValues) {
    const costname = cost[ponctualCostNameIndex];

    // Check if Cost exist
    if (costname !== '') {
      const costProbability = cost[ponctualCostProbabilityIndex];

      // Check Probability Withlist
      if (probabilityWithlist.includes(costProbability)) {

        // Get Cost data
        const costValue = parseFloat(cost[ponctualCostIndex]);
        const costDate = cost[ponctualCostsDateIndex];
        const costSource = cost[ponctualCostSourceIndex];

        // Check if cost into Budget
        if (costDate >= startDate && costDate <= endDate) {

          // Add Cost row
          ponctualCostsBudgetRows.push([costSource, [costname, '', '', costValue]]);
        }
      }
    }
  }

  // Return Costs Rows
  return ponctualCostsBudgetRows
}

function CalculateFixCostsBudgetRows(startDate, endDate, probabilityWithlist) {
  const sheet = costsSheet;
  const row = costsTabsRow;
  const col = costsCol;
  const width = fixCostsWidth;

  // Initialize Rows
  const fixCostsModuleRows = [];

  // Get Cost Datas
  const fixCostsDataValues = GetFixCostsRange().getValues();

  // Get Indexes
  const fixCostsCol = getColumnFromA1(getColumnFromName(sheet, fixCostsName, row - 1, col, costsWidth));
  const fixCostNameIndex = getColumnFromA1(getColumnFromName(sheet, costNameName, row, fixCostsCol, width)) - fixCostsCol;
  const fixCostIndex = getColumnFromA1(getColumnFromName(sheet, costValueName, row, fixCostsCol, width)) - fixCostsCol;
  const fixCostStartingDateIndex = getColumnFromA1(getColumnFromName(sheet, startingDateName, row, fixCostsCol, width)) - fixCostsCol;
  const fixCostEndingDateIndex = getColumnFromA1(getColumnFromName(sheet, endingDateName, row, fixCostsCol, width)) - fixCostsCol;
  const fixCostProbabilityIndex = getColumnFromA1(getColumnFromName(sheet, probabilityName, row, fixCostsCol, width)) - fixCostsCol;
  const fixCostSourceIndex = getColumnFromA1(getColumnFromName(sheet, sourcesName, row, fixCostsCol, width)) - fixCostsCol;

  // for each Cost
  for (let cost of fixCostsDataValues) {
    const costname = cost[fixCostNameIndex];

    // Check if Cost exist
    if (costname !== '') {
      const costProbability = cost[fixCostProbabilityIndex];

      // Check Probability Withlist
      if (probabilityWithlist.includes(costProbability)) {

        // Get Cost data
        const costValue = parseFloat(cost[fixCostIndex]);
        let costStartingDate = cost[fixCostStartingDateIndex];
        let costEndingDate = cost[fixCostEndingDateIndex];
        const costSource = cost[fixCostSourceIndex];

        // Check if Cost into Budget
        if (costStartingDate <= endDate && costEndingDate >= startDate) {

          if (costStartingDate <= startDate) costStartingDate = startDate; // If Cost start before Budget
          if (costEndingDate >= endDate) costEndingDate = endDate; // If Cost end after Budget

          // Get Nb Months
          const nbMonths = GetNbMonthFromDates(costStartingDate, costEndingDate);

          // Add Cost row
          fixCostsModuleRows.push([costSource, [costname, '', '', nbMonths * costValue]]);
        }
      }
    }
  }

  // Return Costs Rows
  return fixCostsModuleRows;
}

function CalculateVariableCostsBudgetRows(startDate, endDate, probabilityWithlist) {
  const sheet = costsSheet;
  const row = costsTabsRow;
  const col = costsCol;
  const length = costsLength;

  // Initialize Rows
  const variableCostsBudgetRows = [];

  // Get Cost Datas
  const variableCostsDataValues = transposeArray(GetVariableCostsRange().getValues());

  // Get Indexes
  const variableCostsCol = getColumnFromA1(getColumnFromName(sheet, variableCostsName, row - 1, col, costsWidth));
  const variableCostsNameIndex = getRowFromName(sheet, costNameName, row, variableCostsCol, length) - (row + 1);
  const variableCostsProbabilityIndex = getRowFromName(sheet, probabilityName, row, variableCostsCol, length) - (row + 1);
  const variableCostsSourceIndex = getRowFromName(sheet, sourcesName, row, variableCostsCol, length) - (row + 1);
  const variableCostStartIndex = Math.max(variableCostsProbabilityIndex, variableCostsSourceIndex, variableCostsSourceIndex) + 1;

  // Get Cost months
  const variableCostsMonths = variableCostsDataValues.shift()

  // For each Cost
  for (let cost of variableCostsDataValues) {
    const costName = cost[variableCostsNameIndex];

    // Check if Cost exist
    if (costName !== '') {
      const costProbability = cost[variableCostsProbabilityIndex];

      // Check Probability Withlist
      if (probabilityWithlist.includes(costProbability)) {

        // Initialize Cost
        let costAcc = 0;

        // Get Cost datas
        const costSource = cost[variableCostsSourceIndex];

        // For each Cost's month
        for (let index = variableCostStartIndex; index < variableCostsMonths.length; index++) {

          // Get Cost value
          const costValue = cost[index];
          if (costValue !== '') {
            const costDate = variableCostsMonths[index];

            // Check if Cost into Budget
            if (costDate >= startDate && costDate <= endDate) {

              // Increment Cost
              costAcc += costValue;
            }
          }
        }
        // Add Cost Row
        if (costAcc != 0) variableCostsBudgetRows.push([costSource, [costName, '', '', costAcc]]);
      }
    }
  }

  // Return Costs Rows
  return variableCostsBudgetRows;
}

function CalculatePonctualIncomesBudgetRows(startDate, endDate, probabilityWithlist) {
  const sheet = incomesSheet;
  const row = incomesTabsRow;
  const col = incomesCol;
  const width = incomesPonctualWidth;
  const ponctualIncomesDataValues = GetPonctualIncomesRange().getValues();

  // Initialize Rows
  ponctualIncomesBudgetRows = [];

  // Get Indexes
  const ponctualIncomesCol = getColumnFromA1(getColumnFromName(sheet, ponctualIncomesName, row - 1, col, incomesWidth));
  const ponctualIncomeNameIndex = getColumnFromA1(getColumnFromName(sheet, incomeNameName, row, ponctualIncomesCol, width)) - ponctualIncomesCol;
  const ponctualIncomeIndex = getColumnFromA1(getColumnFromName(sheet, incomeValueName, row, ponctualIncomesCol, width)) - ponctualIncomesCol;
  const ponctualIncomesDateIndex = getColumnFromA1(getColumnFromName(sheet, dateName, row, ponctualIncomesCol, width)) - ponctualIncomesCol;
  const ponctualIncomeProbabilityIndex = getColumnFromA1(getColumnFromName(sheet, probabilityName, row, ponctualIncomesCol, width)) - ponctualIncomesCol;
  const ponctualIncomeSourceIndex = getColumnFromA1(getColumnFromName(sheet, sourcesName, row, ponctualIncomesCol, width)) - ponctualIncomesCol;

  // For each Income
  for (let income of ponctualIncomesDataValues) {
    const incomename = income[ponctualIncomeNameIndex];

    // Check if Income exist
    if (incomename !== '') {
      const incomeProbability = income[ponctualIncomeProbabilityIndex];

      // Check Probability Withlist
      if (probabilityWithlist.includes(incomeProbability)) {

        // Get Income data
        const incomeValue = parseFloat(income[ponctualIncomeIndex]);
        const incomeDate = income[ponctualIncomesDateIndex];
        const incomeSource = income[ponctualIncomeSourceIndex];

        // Check if Income into Budget
        if (incomeDate >= startDate && incomeDate <= endDate) {

          // Add Income Row
          ponctualIncomesBudgetRows.push([incomeSource, [incomename, '', '', incomeValue]]);
        }
      }
    }
  }

  // Return Income Rows
  return ponctualIncomesBudgetRows;
}

function CalculateFixIncomesBudgetRows(startDate, endDate, probabilityWithlist) {
  const sheet = incomesSheet;
  const row = incomesTabsRow;
  const col = incomesCol;
  const width = fixIncomesWidth;

  // Initialize Row
  const fixIncomesModuleRows = [];

  // Get Income Datas
  const fixIncomesDataValues = GetFixIncomesRange().getValues();

  // Get Indexes
  const fixIncomesCol = getColumnFromA1(getColumnFromName(sheet, fixIncomesName, row - 1, col, incomesWidth));
  const fixIncomeNameIndex = getColumnFromA1(getColumnFromName(sheet, incomeNameName, row, fixIncomesCol, width)) - fixIncomesCol;
  const fixIncomeIndex = getColumnFromA1(getColumnFromName(sheet, incomeValueName, row, fixIncomesCol, width)) - fixIncomesCol;
  const fixIncomeStartingDateIndex = getColumnFromA1(getColumnFromName(sheet, startingDateName, row, fixIncomesCol, width)) - fixIncomesCol;
  const fixIncomeEndingDateIndex = getColumnFromA1(getColumnFromName(sheet, endingDateName, row, fixIncomesCol, width)) - fixIncomesCol;
  const fixIncomeProbabilityIndex = getColumnFromA1(getColumnFromName(sheet, probabilityName, row, fixIncomesCol, width)) - fixIncomesCol;
  const fixIncomeSourceIndex = getColumnFromA1(getColumnFromName(sheet, sourcesName, row, fixIncomesCol, width)) - fixIncomesCol;

  // For each Income
  for (let income of fixIncomesDataValues) {
    const incomename = income[fixIncomeNameIndex];

    // Check if Income exist
    if (incomename !== '') {
      const incomeProbability = income[fixIncomeProbabilityIndex];

      // Check Probability Withlist
      if (probabilityWithlist.includes(incomeProbability)) {

        // Get Income data
        const incomeValue = parseFloat(income[fixIncomeIndex]);
        let incomeStartingDate = income[fixIncomeStartingDateIndex];
        let incomeEndingDate = income[fixIncomeEndingDateIndex];
        const incomeSource = income[fixIncomeSourceIndex];

        // Check if Income into Budget
        if (incomeStartingDate <= endDate && incomeEndingDate >= startDate) {

          if (incomeStartingDate <= startDate) incomeStartingDate = startDate; // If Income start before Budget
          if (incomeEndingDate >= endDate) incomeEndingDate = endDate; // If Income end after Budget

          // Get Nb Months
          const nbMonths = GetNbMonthFromDates(incomeStartingDate, incomeEndingDate);

          // Add Income Rows
          fixIncomesModuleRows.push([incomeSource, [incomename, '', '', nbMonths * incomeValue]]);
        }
      }
    }
  }

  // Return Incomes Rows
  return fixIncomesModuleRows;
}

function CalculateVariableIncomesBudgetRows(startDate, endDate, probabilityWithlist) {
  const sheet = incomesSheet;
  const row = incomesTabsRow;
  const col = incomesCol;
  const length = incomesLength;

  // Initialize Rows
  const variableIncomesBudgetRows = []

  // Get Income Datas
  const variableIncomesDataValues = transposeArray(GetVariableIncomesRange().getValues());

  // Get Indexes
  const variableIncomesCol = getColumnFromA1(getColumnFromName(sheet, variableIncomesName, row - 1, col, incomesWidth));
  const variableIncomesNameIndex = getRowFromName(sheet, incomeNameName, row, variableIncomesCol, length) - (row + 1);
  const variableIncomesProbabilityIndex = getRowFromName(sheet, probabilityName, row, variableIncomesCol, length) - (row + 1);
  const variableIncomesSourceIndex = getRowFromName(sheet, sourcesName, row, variableIncomesCol, length) - (row + 1);
  const variableIncomeStartIndex = Math.max(variableIncomesProbabilityIndex, variableIncomesSourceIndex, variableIncomesSourceIndex) + 1;

  // Get Incomes Months
  const variableIncomesMonths = variableIncomesDataValues.shift();

  // For each Income
  for (let income of variableIncomesDataValues) {
    const incomeName = income[variableIncomesNameIndex];

    // Check if Income exist
    if (incomeName !== '') {
      const incomeProbability = income[variableIncomesProbabilityIndex];

      // Check Probability Withlist
      if (probabilityWithlist.includes(incomeProbability)) {

        // Initialize Income
        let incomeAcc = 0;

        // Get Income datas
        const incomeSource = income[variableIncomesSourceIndex];

        // For each Income's month
        for (let index = variableIncomeStartIndex; index < variableIncomesMonths.length; index++) {

          // Get Income value
          const incomeValue = income[index];
          if (incomeValue !== '') {
            const incomeDate = variableIncomesMonths[index];

            // Check if Income into Budget
            if (incomeDate >= startDate && incomeDate <= endDate) {

              // Increment Income
              incomeAcc += incomeValue;
            }
          }
        }
        // Add Income Row
        if (incomeAcc != 0) variableIncomesBudgetRows.push([incomeSource, [incomeName, '', '', incomeAcc]]);
      }
    }
  }

  // Return Incomes Rows
  return variableIncomesBudgetRows;
}

function CreateModules(budgetRows, isCost) {
  let name;

  // Get Name
  if (isCost) name = costsName;
  else name = incomesName;

  // Initialize Arrays
  const sources = [];
  const sourcesTotals = [];
  const sources_Sorted = [];
  const modules = [];
  let modulesResult = [];

  // For each Row
  for (let index in budgetRows) {

    // Row data
    const source = budgetRows[index][0];
    const budgetRow = budgetRows[index][1];

    // Check if source module already made
    if (!sources.includes(source)) {

      // Add Source in Source list
      sources.push(source);

      // Initialize Source total
      sourcesTotals.push(0);

      // Initialize Module
      modules.push([
        [source, '', '', ''],
        [namesName, '', '', name]
      ]);
    }

    // Get Source index
    const sourceIndex = sources.indexOf(source);

    // Add Row into Module
    modules[sourceIndex].push(budgetRow);

    // Increment Total
    sourcesTotals[sourceIndex] += budgetRow[3];
  }

  // For each Module
  for (let index in modules) {

    // Finalize Module
    modules[index].push([totalName, '', '', sourcesTotals[index]]);

    // Make the Source list
    sources_Sorted.push(modules[index][0][0]);

    // Make Module Result
    modulesResult = modulesResult.concat(modules[index]);
  }
  modulesResult.push(['', '', '', '']);

  // Calculate Global Total
  const total = sourcesTotals.reduce((accumulator, currentValue) => accumulator + currentValue, 0);

  // Make Total Row
  const totalRow = ['Total ' + name, total, total, total];

  // Return Datas
  return [modulesResult, [totalRow, sourcesTotals, sources_Sorted]];
}

function WrapBudgetValues(costsValues, incomesValues) {

  // Get Budget length
  const maxRows = Math.max(costsValues[0].length, incomesValues[0].length);

  // Initialize Budget Values
  const budgetValues = [];

  // For each Line
  for (let i = 0; i < maxRows; i++) {
    let row1 = [];
    let row2 = [];

    // If line < Cost length
    if (i < costsValues[0].length) row1 = costsValues[0][i];
    else row1 = ['', '', '', '']

    // If line < Income length
    if (i < incomesValues[0].length) row2 = incomesValues[0][i];
    else row2 = ['', '', '', ''];

    // Make row
    const mergedRow = row1.concat(row2);

    // Add Row to Budget
    budgetValues.push(mergedRow);
  }

  // Get Top and Bottom Row
  const titleRow = [costsName, '', '', '', incomesName, '', '', ''];
  const totalBudgetRow = costsValues[1][0].concat(incomesValues[1][0]);
  const balanceBudgetRow = [balanceName, '', '', '', totalBudgetRow[7] - totalBudgetRow[3], '', '', ''];

  // Add Top and Bottom Row
  budgetValues.unshift(titleRow);
  budgetValues.push(totalBudgetRow);
  budgetValues.push(['', '', '', '', '', '', '', '']);
  budgetValues.push(balanceBudgetRow);

  // Return Budget values
  return budgetValues;
}
/* 
function SetBudgetStyle(sheet, row, col) {
  const length = sheet.getLastRow();
  const costsCol = col;
  const incomesCol = col + 4;

  // Get Columns & Rows

  const costsColumn = getA1FromCol(costsCol);
  const incomesColumn = getA1FromCol(incomesCol);

  const costsEndColumn = getColumnOffset(costsColumn, 3);
  const incomesEndColumn = getColumnOffset(incomesColumn, 3);

  const charCol = col + budgetWidth + 1;
  const chartsColumn = getA1FromCol(charCol);
  const costsChartRow = row + 1;
  const incomesChartRow = row + budgetChartLength + 3;

  const titleRow = getRowFromName(sheet, costsName, row, costsCol, length);
  const totalsRow = getRowFromName(sheet, totalCostsName, row, costsCol, length);
  const balanceRow = getRowFromName(sheet, balanceName, row, costsCol, length);

  const chartBorderRanges = []
  chartBorderRanges.push(chartsColumn + costsChartRow + ':' + getColumnOffset(chartsColumn, budgetChartWidth) + (costsChartRow + budgetChartLength));
  chartBorderRanges.push(chartsColumn + incomesChartRow + ':' + getColumnOffset(chartsColumn, budgetChartWidth) + (incomesChartRow + budgetChartLength));

  // Get Range

  const budgetRange = costsColumn + titleRow + ':' + incomesEndColumn + balanceRow;
  const range = sheet.getRange(budgetRange);

  // Get Setup Datas

  // Setup Ranges
  const setupCostsIncomesRange = GetCostIncomeRange();
  const setupCostsRange = GetCostRange();
  const setupIncomesRange = GetIncomeRange();

  // Setup Values
  const titles = GetValues(setupCostsIncomesRange);
  const costsSources = GetValues(setupCostsRange);
  const incomesSources = GetValues(setupIncomesRange);

  // Setup Colors
  const titleColors = GetColors(setupCostsIncomesRange);
  const costsColors = GetColors(setupCostsRange);
  const incomesColors = GetColors(setupIncomesRange);

  // Setup TextStyle
  const titleTextStyles = GetTextStyles(setupCostsIncomesRange);
  const costsTextStyles = GetTextStyles(setupCostsRange);
  const incomesTextStyles = GetTextStyles(setupIncomesRange);

  // Build Color Matching
  const keys = [
    ...titles,
    ...costsSources,
    ...incomesSources
  ];

  const colors = [
    ...titleColors,
    ...costsColors,
    ...incomesColors
  ];

  const textStyles = [
    ...titleTextStyles,
    ...costsTextStyles,
    ...incomesTextStyles
  ];

  // Get Textstyles
  const dataTextStyle = SpreadsheetApp.newTextStyle()
    .setFontSize(10)
    .setForegroundColor(GetDarkModeTextColor())
    .setItalic(true)
    .build();

  const totalTextStyle = SpreadsheetApp.newTextStyle()
    .setFontSize(12)
    .setForegroundColor(GetDarkModeTextColor())
    .setBold(true)
    .build();

  const balanceTextStyle = SpreadsheetApp.newTextStyle()
    .setFontSize(15)
    .setForegroundColor(GetDarkModeTextColor())
    .setBold(true)
    .build();

  const balanceValueTextStyle = SpreadsheetApp.newTextStyle()
    .setFontSize(15)
    .setForegroundColor(getHighlightcolor1())
    .setBold(true)
    .build();

  const backgroundColor = GetDarkModeColor();

  // Get Budget Values
  const costsValues = sheet.getRange(costsColumn + titleRow + ':' + costsColumn + balanceRow).getValues();
  const incomesValues = sheet.getRange(incomesColumn + titleRow + ':' + incomesColumn + balanceRow).getValues();

  // Prepare Variable
  const backgroundColors = [];
  const budgetTextStyles = [];
  const borderRanges = [];
  const mergeRanges = [];

  let costIndex;
  let incomeIndex;

  let costColor;
  let costDataColor;
  let costNameColor;
  let costTotalColor;
  let costBorder;
  let costTextStyle;

  let incomeColor;
  let incomeDataColor;
  let incomeNameColor;
  let incomeTotalColor;
  let incomeBorder;
  let incomeTextStyle;

  for (let currentRow = 0; currentRow < (balanceRow - titleRow + 1); currentRow++) {

    let leftRow = [backgroundColor, backgroundColor, backgroundColor, backgroundColor];
    let rightRow = [backgroundColor, backgroundColor, backgroundColor, backgroundColor];

    let leftFontRow = [dataTextStyle, dataTextStyle, dataTextStyle, dataTextStyle];
    let rightFontRow = [dataTextStyle, dataTextStyle, dataTextStyle, dataTextStyle];

    const costValue = costsValues[currentRow][0];

    ///// COST SIDE /////

    if (costValue != '') {

      // Is New Cost or Title
      if (keys.includes(costValue)) {

        costIndex = keys.indexOf(costValue);
        costColor = colors[costIndex];
        costDataColor = lightenDarkenColor(costColor, 0.9);
        costNameColor = lightenDarkenColor(costColor, 0.75);
        costTotalColor = lightenDarkenColor(costColor, 0.25);
        costTextStyle = textStyles[costIndex];

        // Budget Title Formating
        if (titles.includes(costValue)) budgetCostTextStyle = SpreadsheetApp.newTextStyle()
          .setForegroundColor(costTextStyle.getForegroundColor())
          .setBold(true)
          .setFontSize(15)
          .build();

        // Costs Titles Formating
        else budgetCostTextStyle = SpreadsheetApp.newTextStyle()
          .setForegroundColor(costTextStyle.getForegroundColor())
          .setBold(true)
          .setFontSize(12)
          .build();

        // Set Left Rows
        leftRow = [costColor, costColor, costColor, costColor];
        leftFontRow = [budgetCostTextStyle, budgetCostTextStyle, budgetCostTextStyle, budgetCostTextStyle];

        // Save Border & Merging
        costBorder = costsColumn + (currentRow + row + 1) + ':';
        mergeRanges.push(costsColumn + (currentRow + row + 1) + ':' + costsEndColumn + (currentRow + row + 1));

      }

      // First Row of Current Cost
      else if (costValue == namesName) {
        leftRow = [costNameColor, costNameColor, costNameColor, costNameColor];
        leftFontRow = [totalTextStyle, totalTextStyle, totalTextStyle, totalTextStyle];
      }

      // Total of Current Cost
      else if (costValue == totalName) {
        leftRow = [costTotalColor, costTotalColor, costTotalColor, costTotalColor];
        costBorder += costsEndColumn + (currentRow + row);
        borderRanges.push(costBorder);
        leftFontRow = [totalTextStyle, totalTextStyle, totalTextStyle, totalTextStyle];
      }

      // Total of Costs
      else if (costValue == totalCostsName || costValue == costsName) {
        leftRow = [titleColors[0], titleColors[0], titleColors[0], titleColors[0]];
        borderRanges.push(costsColumn + (currentRow + row) + ':' + costsEndColumn + (currentRow + row));
        leftFontRow = [totalTextStyle, totalTextStyle, totalTextStyle, totalTextStyle];
      }

      // Balance
      else if (costValue == balanceName) {
        leftRow = [getHighlightcolor1(), getHighlightcolor1(), getHighlightcolor1(), getHighlightcolor1()];
        borderRanges.push(costsColumn + (currentRow + row) + ':' + incomesEndColumn + (currentRow + row));
        leftFontRow = [balanceTextStyle, balanceTextStyle, balanceTextStyle, balanceTextStyle];
      }

      // Empty
      else {
        leftRow = [costDataColor, costDataColor, costDataColor, costDataColor];
        leftFontRow = [dataTextStyle, dataTextStyle, dataTextStyle, dataTextStyle];
      }
    }

    ///// INCOME SIDE /////

    const incomeValue = incomesValues[currentRow][0];

    if (incomeValue != '') {

      // New Incomes or Income Title
      if (keys.includes(incomeValue)) {
        incomeIndex = keys.indexOf(incomeValue);
        incomeColor = colors[incomeIndex];
        incomeDataColor = lightenDarkenColor(incomeColor, 0.9);
        incomeNameColor = lightenDarkenColor(incomeColor, 0.75);
        incomeTotalColor = lightenDarkenColor(incomeColor, 0.25);
        incomeTextStyle = textStyles[incomeIndex];

        // Budget Title Formatting
        if (titles.includes(incomeValue)) budgetIncomeTextStyle = SpreadsheetApp.newTextStyle()
          .setForegroundColor(incomeTextStyle.getForegroundColor())
          .setBold(true)
          .setFontSize(15)
          .build();

        // Income Title Formatting
        else budgetIncomeTextStyle = SpreadsheetApp.newTextStyle()
          .setForegroundColor(incomeTextStyle.getForegroundColor())
          .setBold(true)
          .setFontSize(12)
          .build();

        // Set Right Row
        rightRow = [incomeColor, incomeColor, incomeColor, incomeColor];
        rightFontRow = [budgetIncomeTextStyle, budgetIncomeTextStyle, budgetIncomeTextStyle, budgetIncomeTextStyle];

        // Save Border & Merging
        incomeBorder = incomesColumn + (currentRow + row + 1) + ':';
        mergeRanges.push(incomesColumn + (currentRow + row + 1) + ':' + incomesEndColumn + (currentRow + row + 1));
      }

      // Current Income First Row
      else if (incomeValue == namesName) {
        rightRow = [incomeNameColor, incomeNameColor, incomeNameColor, incomeNameColor];
        rightFontRow = [totalTextStyle, totalTextStyle, totalTextStyle, totalTextStyle];
      }

      // Current Income Total
      else if (incomeValue == totalName) {
        rightRow = [incomeTotalColor, incomeTotalColor, incomeTotalColor, incomeTotalColor];
        incomeBorder += incomesEndColumn + (currentRow + row + 1);
        borderRanges.push(incomeBorder);
        rightFontRow = [totalTextStyle, totalTextStyle, totalTextStyle, totalTextStyle];
      }

      // Incomes Total
      else if (incomeValue == totalIncomesName || incomeValue == incomesName) {
        rightRow = [titleColors[1], titleColors[1], titleColors[1], titleColors[1]];
        borderRanges.push(incomesColumn + (currentRow + row) + ':' + incomesEndColumn + (currentRow + row));
        rightFontRow = [totalTextStyle, totalTextStyle, totalTextStyle, totalTextStyle];
      }

      // Balance
      else if (costValue == balanceName) {
        rightRow = [backgroundColor, backgroundColor, backgroundColor, backgroundColor];
        rightFontRow = [balanceValueTextStyle, balanceValueTextStyle, balanceValueTextStyle, balanceValueTextStyle];

      }

      // Empty
      else {
        rightRow = [incomeDataColor, incomeDataColor, incomeDataColor, incomeDataColor];
        rightFontRow = [dataTextStyle, dataTextStyle, dataTextStyle, dataTextStyle];
      };
    }

    // Combine both side and add to Array

    backgroundColors.push(leftRow.concat(rightRow));
    budgetTextStyles.push(leftFontRow.concat(rightFontRow));
  }

  // Tab Datas Borders
  borderRanges.push(costsColumn + (titleRow) + ':' + costsEndColumn + (balanceRow - 1));
  borderRanges.push(incomesColumn + (titleRow) + ':' + incomesEndColumn + (balanceRow - 1));

  // Get Merging Ranges

  // Totals
  mergeRanges.push(getColumnOffset(costsColumn, 1) + (totalsRow) + ':' + costsEndColumn + (totalsRow));
  mergeRanges.push(getColumnOffset(incomesColumn, 1) + (totalsRow) + ':' + incomesEndColumn + (totalsRow));

  // Balance
  mergeRanges.push(costsColumn + (balanceRow) + ':' + costsEndColumn + (balanceRow));
  mergeRanges.push(incomesColumn + (balanceRow) + ':' + incomesEndColumn + (balanceRow));

  // Borders

  // Data Borders
  const borderRangeList = sheet.getRangeList(borderRanges);
  borderRangeList.setBorder(true, true, true, true, null, null, GetDarkModeTextColor(), SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Tabs Borders
  const highlightBorderRange = [
    costsColumn + (balanceRow) + ':' + incomesEndColumn + (balanceRow),
    costsColumn + titleRow + ':' + incomesEndColumn + (balanceRow),
    ...chartBorderRanges
  ];
  sheet.getRangeList(highlightBorderRange).setBorder(true, true, true, true, null, null, getHighlightcolor1(), SpreadsheetApp.BorderStyle.SOLID_THICK);

  // Budget cells width
  sheet.setColumnWidth(costsCol, 150);
  sheet.setColumnWidths((costsCol + 1), 2, 50);
  sheet.setColumnWidth((costsCol + 3), 100);

  sheet.setColumnWidth(incomesCol, 150);
  sheet.setColumnWidths((incomesCol + 1), 2, 50);
  sheet.setColumnWidth((incomesCol + 3), 100);

  // Budget/ Chart Separator Width
  sheet.setColumnWidth(charCol - 1, 25);

  // Chart Width
  sheet.setColumnWidths(charCol, budgetChartWidth, 50);

  // Next Budget Separator Width
  sheet.setColumnWidth(charCol + budgetChartWidth + 1, 150);

  // Merge Cells
  sheet.getRange(budgetRange).breakApart()
  for (let index in mergeRanges) {
    sheet.getRange(mergeRanges[index]).merge();
  };

  range.setBackgrounds(backgroundColors);
  range.setTextStyles(budgetTextStyles);

}
 */
function SetBudgetStyle(sheet, row, col, budgetColor) {
  const length = sheet.getLastRow();
  const costsCol = col;
  const incomesCol = col + 4;

  // Get Columns & Rows
  const costsColumn = getA1FromCol(costsCol);
  const incomesColumn = getA1FromCol(incomesCol);
  const costsEndColumn = getColumnOffset(costsColumn, 3);
  const incomesEndColumn = getColumnOffset(incomesColumn, 3);
  const charCol = col + budgetWidth + 1;
  const chartsColumn = getA1FromCol(charCol);
  const costsChartRow = row + 1;
  const incomesChartRow = row + budgetChartLength + 3;

  const titleRow = getRowFromName(sheet, costsName, row, costsCol, length);
  const totalsRow = getRowFromName(sheet, totalCostsName, row, costsCol, length);
  const balanceRow = getRowFromName(sheet, balanceName, row, costsCol, length);

  const chartBorderRanges = [
    `${chartsColumn}${costsChartRow}:${getColumnOffset(chartsColumn, budgetChartWidth)}${costsChartRow + budgetChartLength}`,
    `${chartsColumn}${incomesChartRow}:${getColumnOffset(chartsColumn, budgetChartWidth)}${incomesChartRow + budgetChartLength}`
  ];

  // Get Range
  const budgetRange = `${costsColumn}${titleRow}:${incomesEndColumn}${balanceRow}`;
  const range = sheet.getRange(budgetRange);

  // Get Setup Datas
  const setupCostsIncomesRange = GetCostIncomeRange();
  const setupCostsRange = GetCostRange();
  const setupIncomesRange = GetIncomeRange();

  // Get Setup Values
  const titles = GetValues(setupCostsIncomesRange);
  const costsSources = GetValues(setupCostsRange);
  const incomesSources = GetValues(setupIncomesRange);

  // Get Setup Colors
  const titleColors = GetColors(setupCostsIncomesRange);
  const costsColors = GetColors(setupCostsRange);
  const incomesColors = GetColors(setupIncomesRange);

  // Get Setup TextStyle
  const titleTextStyles = GetTextStyles(setupCostsIncomesRange);
  const costsTextStyles = GetTextStyles(setupCostsRange);
  const incomesTextStyles = GetTextStyles(setupIncomesRange);

  // Build Color Matching
  const keys = [...titles, ...costsSources, ...incomesSources];
  const colors = [...titleColors, ...costsColors, ...incomesColors];
  const textStyles = [...titleTextStyles, ...costsTextStyles, ...incomesTextStyles];

  // Get Textstyles
  const dataTextStyle = SpreadsheetApp.newTextStyle()
    .setFontSize(10)
    .setForegroundColor(GetDarkModeTextColor())
    .setItalic(true)
    .build();

  const totalTextStyle = SpreadsheetApp.newTextStyle()
    .setFontSize(12)
    .setForegroundColor(GetDarkModeTextColor())
    .setBold(true)
    .build();

  const balanceTextStyle = SpreadsheetApp.newTextStyle()
    .setFontSize(15)
    .setForegroundColor(GetDarkModeTextColor())
    .setBold(true)
    .build();

  const balanceValueTextStyle = SpreadsheetApp.newTextStyle()
    .setFontSize(15)
    .setForegroundColor(getHighlightcolor1())
    .setBold(true)
    .build();

  const backgroundColor = GetDarkModeColor();

  // Get Budget Values
  const costsValues = sheet.getRange(costsColumn + titleRow + ':' + costsColumn + balanceRow).getValues();
  const incomesValues = sheet.getRange(incomesColumn + titleRow + ':' + incomesColumn + balanceRow).getValues();

  // Prepare Variable
  const backgroundColors = [];
  const budgetTextStyles = [];
  const borderRanges = [];
  const mergeRanges = [];

  let costIndex;
  let incomeIndex;

  let costColor;
  let costDataColor;
  let costNameColor;
  let costTotalColor;
  let costBorder;
  let costTextStyle;

  let incomeColor;
  let incomeDataColor;
  let incomeNameColor;
  let incomeTotalColor;
  let incomeBorder;
  let incomeTextStyle;

  for (let currentRow = 0; currentRow < (balanceRow - titleRow + 1); currentRow++) {
    let leftRow = [backgroundColor, backgroundColor, backgroundColor, backgroundColor];
    let rightRow = [backgroundColor, backgroundColor, backgroundColor, backgroundColor];

    let leftFontRow = [dataTextStyle, dataTextStyle, dataTextStyle, dataTextStyle];
    let rightFontRow = [dataTextStyle, dataTextStyle, dataTextStyle, dataTextStyle];

    const costValue = costsValues[currentRow][0];

    ///// COST SIDE /////
    if (costValue != '') {
      // Is New Cost or Title
      if (keys.includes(costValue)) {
        costIndex = keys.indexOf(costValue);
        costColor = colors[costIndex];
        costDataColor = lightenDarkenColor(costColor, 0.9);
        costNameColor = lightenDarkenColor(costColor, 0.75);
        costTotalColor = lightenDarkenColor(costColor, 0.25);
        costTextStyle = textStyles[costIndex];

        // Budget Title Formating
        if (titles.includes(costValue)) budgetCostTextStyle = SpreadsheetApp.newTextStyle()
          .setForegroundColor(costTextStyle.getForegroundColor())
          .setBold(true)
          .setFontSize(15)
          .build();

        // Costs Titles Formating
        else budgetCostTextStyle = SpreadsheetApp.newTextStyle()
          .setForegroundColor(costTextStyle.getForegroundColor())
          .setBold(true)
          .setFontSize(12)
          .build();

        // Set Left Rows
        leftRow = [costColor, costColor, costColor, costColor];
        leftFontRow = [budgetCostTextStyle, budgetCostTextStyle, budgetCostTextStyle, budgetCostTextStyle];

        // Save Border & Merging
        costBorder = costsColumn + (currentRow + row + 1) + ':';
        mergeRanges.push(costsColumn + (currentRow + row + 1) + ':' + costsEndColumn + (currentRow + row + 1));
      }

      // First Row of Current Cost
      else if (costValue == namesName) {
        leftRow = [costNameColor, costNameColor, costNameColor, costNameColor];
        leftFontRow = [totalTextStyle, totalTextStyle, totalTextStyle, totalTextStyle];
      }

      // Total of Current Cost
      else if (costValue == totalName) {
        leftRow = [costTotalColor, costTotalColor, costTotalColor, costTotalColor];
        costBorder += costsEndColumn + (currentRow + row);
        borderRanges.push(costBorder);
        leftFontRow = [totalTextStyle, totalTextStyle, totalTextStyle, totalTextStyle];
      }

      // Total of Costs
      else if (costValue == totalCostsName || costValue == costsName) {
        leftRow = [titleColors[0], titleColors[0], titleColors[0], titleColors[0]];
        borderRanges.push(costsColumn + (currentRow + row) + ':' + costsEndColumn + (currentRow + row));
        leftFontRow = [totalTextStyle, totalTextStyle, totalTextStyle, totalTextStyle];
      }

      // Balance
      else if (costValue == balanceName) {
        leftRow = [getHighlightcolor1(), getHighlightcolor1(), getHighlightcolor1(), getHighlightcolor1()];
        borderRanges.push(costsColumn + (currentRow + row) + ':' + incomesEndColumn + (currentRow + row));
        leftFontRow = [balanceTextStyle, balanceTextStyle, balanceTextStyle, balanceTextStyle];
      }

      // Empty
      else {
        leftRow = [costDataColor, costDataColor, costDataColor, costDataColor];
        leftFontRow = [dataTextStyle, dataTextStyle, dataTextStyle, dataTextStyle];
      }
    }

    ///// INCOME SIDE /////
    const incomeValue = incomesValues[currentRow][0];
    if (incomeValue != '') {
      // New Incomes or Income Title
      if (keys.includes(incomeValue)) {
        incomeIndex = keys.indexOf(incomeValue);
        incomeColor = colors[incomeIndex];
        incomeDataColor = lightenDarkenColor(incomeColor, 0.9);
        incomeNameColor = lightenDarkenColor(incomeColor, 0.75);
        incomeTotalColor = lightenDarkenColor(incomeColor, 0.25);
        incomeTextStyle = textStyles[incomeIndex];

        // Budget Title Formatting
        if (titles.includes(incomeValue)) budgetIncomeTextStyle = SpreadsheetApp.newTextStyle()
          .setForegroundColor(incomeTextStyle.getForegroundColor())
          .setBold(true)
          .setFontSize(15)
          .build();

        // Income Title Formatting
        else budgetIncomeTextStyle = SpreadsheetApp.newTextStyle()
          .setForegroundColor(incomeTextStyle.getForegroundColor())
          .setBold(true)
          .setFontSize(12)
          .build();

        // Set Right Row
        rightRow = [incomeColor, incomeColor, incomeColor, incomeColor];
        rightFontRow = [budgetIncomeTextStyle, budgetIncomeTextStyle, budgetIncomeTextStyle, budgetIncomeTextStyle];

        // Save Border & Merging
        incomeBorder = incomesColumn + (currentRow + row + 1) + ':';
        mergeRanges.push(incomesColumn + (currentRow + row + 1) + ':' + incomesEndColumn + (currentRow + row + 1));
      }

      // Current Income First Row
      else if (incomeValue == namesName) {
        rightRow = [incomeNameColor, incomeNameColor, incomeNameColor, incomeNameColor];
        rightFontRow = [totalTextStyle, totalTextStyle, totalTextStyle, totalTextStyle];
      }

      // Current Income Total
      else if (incomeValue == totalName) {
        rightRow = [incomeTotalColor, incomeTotalColor, incomeTotalColor, incomeTotalColor];
        incomeBorder += incomesEndColumn + (currentRow + row + 1);
        borderRanges.push(incomeBorder);
        rightFontRow = [totalTextStyle, totalTextStyle, totalTextStyle, totalTextStyle];
      }

      // Incomes Total
      else if (incomeValue == totalIncomesName || incomeValue == incomesName) {
        rightRow = [titleColors[1], titleColors[1], titleColors[1], titleColors[1]];
        borderRanges.push(incomesColumn + (currentRow + row) + ':' + incomesEndColumn + (currentRow + row));
        rightFontRow = [totalTextStyle, totalTextStyle, totalTextStyle, totalTextStyle];
      }

      // Balance
      else if (costValue == balanceName) {
        rightRow = [backgroundColor, backgroundColor, backgroundColor, backgroundColor];
        rightFontRow = [balanceValueTextStyle, balanceValueTextStyle, balanceValueTextStyle, balanceValueTextStyle];
      }

      // Empty
      else {
        rightRow = [incomeDataColor, incomeDataColor, incomeDataColor, incomeDataColor];
        rightFontRow = [dataTextStyle, dataTextStyle, dataTextStyle, dataTextStyle];
      };
    }

    // Combine both side and add to Array
    backgroundColors.push(leftRow.concat(rightRow));
    budgetTextStyles.push(leftFontRow.concat(rightFontRow));
  }

  // Tab Datas Borders
  borderRanges.push(costsColumn + (titleRow) + ':' + costsEndColumn + (balanceRow - 1));
  borderRanges.push(incomesColumn + (titleRow) + ':' + incomesEndColumn + (balanceRow - 1));

  // Get Merging Ranges
  // Totals
  mergeRanges.push(getColumnOffset(costsColumn, 1) + (totalsRow) + ':' + costsEndColumn + (totalsRow));
  mergeRanges.push(getColumnOffset(incomesColumn, 1) + (totalsRow) + ':' + incomesEndColumn + (totalsRow));

  // Balance
  mergeRanges.push(costsColumn + (balanceRow) + ':' + costsEndColumn + (balanceRow));
  mergeRanges.push(incomesColumn + (balanceRow) + ':' + incomesEndColumn + (balanceRow));

  // Borders
  // Data Borders
  const borderRangeList = sheet.getRangeList(borderRanges);
  borderRangeList.setBorder(true, true, true, true, null, null, GetDarkModeTextColor(), SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Tabs Borders
  const highlightBorderRange = [
    costsColumn + (balanceRow) + ':' + incomesEndColumn + (balanceRow),
    costsColumn + titleRow + ':' + incomesEndColumn + (balanceRow),
    ...chartBorderRanges
  ];
  sheet.getRangeList(highlightBorderRange).setBorder(true, true, true, true, null, null, budgetColor, SpreadsheetApp.BorderStyle.SOLID_THICK);

  // Budget cells width
  sheet.setColumnWidth(costsCol, 150);
  sheet.setColumnWidths((costsCol + 1), 2, 50);
  sheet.setColumnWidth((costsCol + 3), 100);

  sheet.setColumnWidth(incomesCol, 150);
  sheet.setColumnWidths((incomesCol + 1), 2, 50);
  sheet.setColumnWidth((incomesCol + 3), 100);

  // Budget/ Chart Separator Width
  sheet.setColumnWidth(charCol - 1, 25);

  // Chart Width
  sheet.setColumnWidths(charCol, budgetChartWidth, 50);

  // Next Budget Separator Width
  sheet.setColumnWidth(charCol + budgetChartWidth + 1, 150);

  // Merge Cells
  sheet.getRange(budgetRange).breakApart()
  for (let index in mergeRanges) {
    sheet.getRange(mergeRanges[index]).merge();
  };

  range.setBackgrounds(backgroundColors);
  range.setTextStyles(budgetTextStyles);
}

function CreateBudgetChart(sheet, row, col, budgetTotals) {

  // Get Columns & Row
  const chartsColumn = getA1FromCol(col + budgetWidth + 1);
  const costsChartRow = row + 1;
  const incomesChartRow = row + budgetChartLength + 3;

  // Setup Ranges
  const setupCostsIncomesRange = GetCostIncomeRange();
  const setupCostsRange = GetCostRange();
  const setupIncomesRange = GetIncomeRange();

  // Setup Values
  const titles = GetValues(setupCostsIncomesRange);
  const costsSources = GetValues(setupCostsRange);
  const incomesSources = GetValues(setupIncomesRange);

  // Setup Colors
  const titleColors = GetColors(setupCostsIncomesRange);
  const costsColors = GetColors(setupCostsRange);
  const incomesColors = GetColors(setupIncomesRange);

  // Get Chart Ranges
  const costsChartRange = sheet.getRange(chartsColumn + costsChartRow);
  const costsChartMergeRange = sheet.getRange(chartsColumn + costsChartRow + ':' + getColumnOffset(chartsColumn, budgetChartWidth) + (costsChartRow + budgetChartLength));

  const incomesChartRange = sheet.getRange(chartsColumn + incomesChartRow);
  const incomesChartMergeRange = sheet.getRange(chartsColumn + incomesChartRow + ':' + getColumnOffset(chartsColumn, budgetChartWidth) + (incomesChartRow + budgetChartLength));

  // Get Cost Sources
  const costsSources_Sorted = budgetTotals[1];
  const costColors_Sorted = [];

  const incomesSources_Sorted = budgetTotals[3];
  const incomeColors_Sorted = [];

  // Get Chart Text Style
  const costTitleTextStyle = Charts.newTextStyle()
    .setFontSize(20)
    .setColor(titleColors[0])
    .build();

  const incomeTitleTextStyle = Charts.newTextStyle()
    .setFontSize(20)
    .setColor(titleColors[1])
    .build();

  // Get Totals
  const costsTotals = budgetTotals[0];
  const incomesTotals = budgetTotals[2];


  // Building the Costs DataBase

  // Initialize Cost Database Column
  const costsDatas = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, sourcesName)
    .addColumn(Charts.ColumnType.NUMBER, totalsName);

  // For each Cost sources
  for (let index in costsSources_Sorted) {
    
    // Get Cost Index
    const setupIndex = costsSources.indexOf(costsSources_Sorted[index]);

    // Add Color to sorted List
    costColors_Sorted.push(costsColors[setupIndex]);

    // Add Row to Cost Database
    costsDatas.addRow([costsSources_Sorted[index], costsTotals[index]]);
  }

  // Build Cost Database
  costsDatas.build()

  // Building the Incomes DataBase

  // Initialize Income Database Column
  const incomesDatas = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, sourcesName)
    .addColumn(Charts.ColumnType.NUMBER, totalsName);

  // For each Income Source
  for (let index in incomesSources_Sorted) {

    // Get Income Index
    const setupIndex = incomesSources.indexOf(incomesSources_Sorted[index]);

    // Add Color to sorted List
    incomeColors_Sorted.push(incomesColors[setupIndex]);

    // Add Row to Income Database
    incomesDatas.addRow([incomesSources_Sorted[index], incomesTotals[index]]);
  }

  // Build Income Database
  costsDatas.build();

  // Building the Costs Chart
  const costsChart = Charts.newPieChart()
    .setDataTable(costsDatas)
    .setColors(costColors_Sorted)
    .setBackgroundColor(GetDarkModeColor())
    .setTitle(costsName)
    .setTitleTextStyle(costTitleTextStyle)
    .setDimensions(350, 400)
    .setOption(
      'legend', {
      position: 'top',
      textStyle: {
        color: GetDarkModeTextColor(),
        fontSize: 15,
        bold: true
      },
      maxLines: 5
    })
    .setOption('pieHole', 0.5)
    .setOption('tooltip', {
      ignoreBounds: true,
    })
    .build();

  // HTML Wizardry
  var htmlOutput = HtmlService.createHtmlOutput().setTitle('My Chart');
  let costsChartImageData = Utilities.base64Encode(costsChart.getAs('image/png').getBytes());

  var costsChartURL = "data:image/png;base64," + encodeURI(costsChartImageData);
  htmlOutput.append("Render chart server side: <br/>");
  htmlOutput.append("<img border=\"1\" src=\"" + costsChartURL + "\">");

  // Make Cost Cell Image
  const costsChartCellImage = SpreadsheetApp.newCellImage()
    .setSourceUrl(costsChartURL)
    .build();

  // Building the Incomes Chart
  const incomesChart = Charts.newPieChart()
    .setDataTable(incomesDatas)
    .setColors(incomeColors_Sorted)
    .setBackgroundColor(GetDarkModeColor())
    .setTitle(incomesName)
    .setTitleTextStyle(incomeTitleTextStyle)
    .setDimensions(350, 400)
    .setOption(
      'legend', {
      position: 'top',
      textStyle: {
        color: GetDarkModeTextColor(),
        fontSize: 15,
        bold: true
      },
      maxLines: 5
    })
    .setOption('pieHole', 0.5)
    .setOption('tooltip', {
      ignoreBounds: true,
    })
    .build();

  // HTML Wizardry
  var htmlOutput = HtmlService.createHtmlOutput().setTitle('My Chart');
  let incomesChartImageData = Utilities.base64Encode(incomesChart.getAs('image/png').getBytes());

  var incomesChartURL = "data:image/png;base64," + encodeURI(incomesChartImageData);
  htmlOutput.append("Render chart server side: <br/>");
  htmlOutput.append("<img border=\"1\" src=\"" + incomesChartURL + "\">");

  // Make the Income Cell Image
  const incomesChartCellImage = SpreadsheetApp.newCellImage()
    .setSourceUrl(incomesChartURL)
    .build();

  // Set Image
  costsChartRange.setValue(costsChartCellImage);
  incomesChartRange.setValue(incomesChartCellImage);

  // Merge Image cells
  costsChartMergeRange.merge();
  incomesChartMergeRange.merge();
}

////////////////////////////
//      ANNUAL BUDGETS    //__________________________________________________________________________________________________________________________________________
////////////////////////////

const GetAnnualBudgetLength = function () {
  return Math.max(annualBudgetSheet.getLastRow(), annualBudgetLength);
}

function GetAnnualBudgetProbabilityWithlist() {
  const sheet = annualBudgetSheet;
  const row = annualBudgetRow;
  const col = getColumnFromA1(annualBudgetParamTitleColumn);
  const length = GetAnnualBudgetLength();

  // Setup Data
  const setupProbabilityRange = GetProbabilityRange();
  const probabilityList = GetValues(setupProbabilityRange);
  const probabilityLength = probabilityList.length - 1;

  // Get Withlist Values
  const probabilityRow = getRowFromName(sheet, probabilityName, row, col, length);
  const timelineProbaRange = sheet.getRange(annualBudgetParam2Column + probabilityRow + ':' + annualBudgetParam2Column + (probabilityRow + probabilityLength - 1));
  const timelineProbaValues = cleanData(timelineProbaRange.getValues());

  probabilityWithlist = [probabilityList[0]];
  probabilityList.shift(); // Get Rid of the Certain Proba

  for (let index in timelineProbaValues) {
    if (timelineProbaValues[index] == true) probabilityWithlist.push(probabilityList[index]);
  }
  return probabilityWithlist;
}

function CreateAnnualBudgets() {
  const sheet = annualBudgetSheet;
  const col = annualBudgetFirstCol;
  const row = annualBudgetRow;
  const length = GetAnnualBudgetLength();

  // Get Columns & Rows

  const paramTitleColumn = annualBudgetParamTitleColumn;
  const paramTitleCol = getColumnFromA1(paramTitleColumn);

  const startingYearRow = getRowFromName(sheet, startingDateName, row, paramTitleCol, length);
  const endingYearRow = getRowFromName(sheet, endingDateName, row, paramTitleCol, length);

  // Get Dates

  const startingYear = sheet.getRange(annualBudgetParam1Column + startingYearRow).getValue();
  const endingYear = sheet.getRange(annualBudgetParam1Column + endingYearRow).getValue();

  // Clean the sheet

  maxRange = sheet.getRange(row, col, sheet.getMaxRows(), sheet.getMaxColumns());
  maxRange.clearContent();
  maxRange.setBorder(true, true, true, true, true, true, GetDarkModeColor(), SpreadsheetApp.BorderStyle.SOLID);
  maxRange.setBackground(GetDarkModeColor());
  maxRange.breakApart();

  // Make the Budgets

  let currentCol = col;

  for (let year = new Date(startingYear); year.getFullYear() <= endingYear.getFullYear(); year.setFullYear(year.getFullYear() + 1)) {
    CreateAnnualBudget(year, sheet, currentCol, row);
    currentCol += 17;
  }

}

function CreateAnnualBudget(year, sheet, col, row) {

  // Initialize Dates
  const yearStart = new Date(year);
  const yearEnd = new Date(year);

  yearStart.setMonth(0);
  yearStart.setDate(1);
  yearEnd.setMonth(11);
  yearEnd.setDate(31);

  // Get Budget Datas
  let budgetValues = GetBudgetValues(yearStart, yearEnd, GetAnnualBudgetProbabilityWithlist());
  let budgetTotals = budgetValues[1];
  budgetTotals = [budgetTotals[0][1], budgetTotals[0][2], budgetTotals[1][1], budgetTotals[1][2]];
  budgetValues = budgetValues[0];

  // Add Top Row
  const yearRow = [year.getFullYear(), '', '', '', '', '', '', ''];
  budgetValues.unshift(yearRow);

  // Get Budget Dimension
  const budgetWidth = budgetValues[0].length;
  const budgetLength = budgetValues.length;

  // Get Range
  const budgetRange = sheet.getRange(row, col, budgetLength, budgetWidth);
  const budgetYearRange = sheet.getRange(row, col, 1, budgetWidth);

  // Set Values
  budgetRange.setValues(budgetValues);

  // Set Title Style
  budgetYearRange
    .setBorder(true, true, true, true, null, null, getHighlightcolor1(), SpreadsheetApp.BorderStyle.SOLID_THICK)
    .setFontColor(getHighlightcolor2())
    .setFontSize(15)
    .setHorizontalAlignment('center')
    .merge();

  // Finalize Budget
  CreateBudgetChart(sheet, row, col, budgetTotals);
  SetBudgetStyle(sheet, row, col, getHighlightcolor1());
}

function RefreshAnnualBudgetMenu() {
  const sheet = annualBudgetSheet;
  const firstRow = 5;

  // Delete previous menu
  const menuRange = sheet.getRange(annualBudgetParamTitleColumn + firstRow + ':' + annualBudgetParam2Column + sheet.getLastRow())
  menuRange.clearContent();
  menuRange.clearDataValidations();
  menuRange.setBorder(true, true, true, true, true, true, GetDarkModeColor(), SpreadsheetApp.BorderStyle.SOLID);
  menuRange.breakApart();

  // Create Probability Selector
  CreateProbabilitySelection(sheet, annualBudgetParamTitleColumn, firstRow);
  SetAnnualBudgetHighlights();
}

function RefreshAnnualBudgetsStyle() {
  const sheet = annualBudgetSheet;
  const col = annualBudgetFirstCol;
  const row = annualBudgetRow;
  const length = GetAnnualBudgetLength();

  // Get Columns & Rows
  const paramTitleColumn = annualBudgetParamTitleColumn;
  const paramTitleCol = getColumnFromA1(paramTitleColumn);

  const startingYearRow = getRowFromName(sheet, startingDateName, row, paramTitleCol, length);
  const endingYearRow = getRowFromName(sheet, endingDateName, row, paramTitleCol, length);

  // Get Dates
  const startingYear = sheet.getRange(annualBudgetParam1Column + startingYearRow).getValue();
  const endingYear = sheet.getRange(annualBudgetParam1Column + endingYearRow).getValue();

  // Refresh Styles
  let currentCol = col;

  for (let year = new Date(startingYear); year.getFullYear() <= endingYear.getFullYear(); year.setFullYear(year.getFullYear() + 1)) {
    SetBudgetStyle(sheet, row, currentCol, getHighlightcolor1());
    currentCol += 17;
  }

}

////////////////////////////
//   PRODUCTION BUDGETS   //__________________________________________________________________________________________________________________________________________
////////////////////////////

function CreateProductionBudgetMenu() {
  const sheet = productionBudgetSheet;
  const row = productionBudgetRow;

  let currentRow = row;

  // Delete previous menu
  const menuRange = sheet.getRange(productionBudgetParamTitleColumn + row + ':' + productionBudgetParam3Column + sheet.getMaxRows())
  menuRange.clearContent();
  menuRange.clearDataValidations();
  menuRange.setBorder(true, true, true, true, true, true, GetDarkModeColor(), SpreadsheetApp.BorderStyle.SOLID);
  menuRange.breakApart();

  // Create Probability Selector
  CreateProbabilitySelection(sheet, productionBudgetParamTitleColumn, currentRow);

  // Create Scope
  const probabilities = GetValues(GetProbabilityRange());
  probabilities.shift();

  // Create Production Scope
  currentRow += probabilities.length + 1
  CreateProductionScope(sheet, productionBudgetParamTitleColumn, currentRow)

  // Create Date menu
  const productionScopeRow = getRowFromName(sheet, productionStageName, row, getColumnFromA1(productionBudgetParam1Column), sheet.getLastRow() - row + 1)
  const productionStages = cleanData(sheet.getRange(productionBudgetParam1Column + (productionScopeRow + 1) + ':' + productionBudgetParam1Column + (sheet.getLastRow())).getValues().filter(String))

  currentRow += productionStages.length + 2;

  // For each Production Stage
  for (let productionStage of productionStages) {

    // First Column
    const datesValues = [
      [productionStage],
      [startingDateName],
      [endingDateName]
    ];

    // Get Ranges

    const dateTitleRange = sheet.getRange(productionBudgetParamTitleColumn + currentRow + ':' + productionBudgetParamTitleColumn + (currentRow + 2));

    const productionStageTitleRange = sheet.getRange(productionBudgetParamTitleColumn + currentRow + ':' + productionBudgetParam3Column + currentRow);

    const startingDateTitleRange = sheet.getRange(productionBudgetParamTitleColumn + (currentRow + 1) + ':' + productionBudgetParam1Column + (currentRow + 1));
    const endingDateTitleRange = sheet.getRange(productionBudgetParamTitleColumn + (currentRow + 2) + ':' + productionBudgetParam1Column + (currentRow + 2));

    const startingDateValueRange = sheet.getRange(productionBudgetParam2Column + (currentRow + 1) + ':' + productionBudgetParam3Column + (currentRow + 1));
    const endingDateValueRange = sheet.getRange(productionBudgetParam2Column + (currentRow + 2) + ':' + productionBudgetParam3Column + (currentRow + 2));

    // Set First Column
    dateTitleRange.setValues(datesValues);

    // Set Date Values
    startingDateValueRange.setValue(Utilities.formatDate(new Date(), "GMT+1", "MMMMMMM yyyy"));
    endingDateValueRange.setValue(Utilities.formatDate(new Date(), "GMT+1", "MMMMMMM yyyy"));

    // Merge Cells
    productionStageTitleRange.merge();
    startingDateTitleRange.merge();
    endingDateTitleRange.merge();
    startingDateValueRange.merge();
    endingDateValueRange.merge();

    // Increment Row
    currentRow += 4;
  }
  SetProductionBudgetHighlights();
}

function CreateProductionBudgets() {
  const sheet = productionBudgetSheet;
  const row = productionBudgetRow;
  const col = productionBudgetFirstCol;

  // Clean the sheet
  maxRange = sheet.getRange(row, col, sheet.getMaxRows(), sheet.getMaxColumns());
  maxRange.clearContent();
  maxRange.setBorder(true, true, true, true, true, true, GetDarkModeColor(), SpreadsheetApp.BorderStyle.SOLID);
  maxRange.setBackground(GetDarkModeColor());
  maxRange.breakApart()

  // Get Production Stages
  const productionScopeRow = getRowFromName(sheet, productionStageName, row, getColumnFromA1(productionBudgetParam1Column), sheet.getLastRow() - row + 1);
  const productionStages = cleanData(sheet.getRange(productionBudgetParam1Column + (productionScopeRow + 1) + ':' + productionBudgetParam1Column + (sheet.getLastRow())).getValues().filter(String));

  // Create Budgets

  let currentCol = col;

  // for each Production Stage
  for (let productionStage of productionStages) {

    // Get Dates
    const stageRow = getRowFromName(sheet, productionStage, row, getColumnFromA1(productionBudgetParamTitleColumn), sheet.getLastRow());
    const startingDate = sheet.getRange(productionBudgetParam2Column + (stageRow + 1)).getValue();
    const endingDate = sheet.getRange(productionBudgetParam2Column + (stageRow + 2)).getValue();

    // Create Budget
    CreateProductionBudget(productionStage, sheet, row, currentCol, startingDate, endingDate)
    currentCol += 17;
  }

}

function CreateProductionBudget(productionStage, sheet, row, col, startingDate, endingDate) {

  // Get Setup Datas
  const productionStages = GetValues(GetProductionRange());
  const productionStagesColors = GetColors(GetProductionRange());

  // Get Production Stage Color
  const productionColor = productionStagesColors[productionStages.indexOf(productionStage)]

  // Get Budget Data
  let budgetValues = GetBudgetValues(startingDate, endingDate, GetProductionBudgetProbabilityWithlist());
  let budgetTotals = budgetValues[1];
  budgetTotals = [budgetTotals[0][1], budgetTotals[0][2], budgetTotals[1][1], budgetTotals[1][2]];
  budgetValues = budgetValues[0];

  // Add Top Row
  const productionRow = [productionStage, '', '', '', '', '', '', ''];
  budgetValues.unshift(productionRow);

  // Get Budget Dimensions
  const budgetWidth = budgetValues[0].length;
  const budgetLength = budgetValues.length;

  // Get Ranges
  const budgetRange = sheet.getRange(row, col, budgetLength, budgetWidth);
  const budgetYearRange = sheet.getRange(row, col, 1, budgetWidth);

  // Set Values
  budgetRange.setValues(budgetValues);

  // Set Title Style
  budgetYearRange
    .setBorder(true, true, true, true, null, null, productionColor, SpreadsheetApp.BorderStyle.SOLID_THICK)
    .setFontColor(getHighlightcolor1())
    .setFontSize(15)
    .setHorizontalAlignment('center')
    .merge();

  // Finalize Budget
  CreateBudgetChart(sheet, row, col, budgetTotals);
  SetBudgetStyle(sheet, row, col,productionColor);
}

function GetProductionBudgetProbabilityWithlist() {
  const sheet = productionBudgetSheet;
  const row = productionBudgetRow;
  const col = getColumnFromA1(productionBudgetParamTitleColumn);
  const length = productionBudgetLength;

  // Setup Data
  const setupProbabilityRange = GetProbabilityRange();
  const probabilityList = GetValues(setupProbabilityRange);
  const probabilityLength = probabilityList.length - 1;

  // Get Withlist Values
  const probabilityRow = getRowFromName(sheet, probabilityName, row, col, length);
  const timelineProbaRange = sheet.getRange(productionBudgetParam2Column + probabilityRow + ':' + productionBudgetParam2Column + (probabilityRow + probabilityLength - 1));
  const timelineProbaValues = cleanData(timelineProbaRange.getValues());

  probabilityWithlist = [probabilityList[0]];
  probabilityList.shift(); // Get Rid of the Certain Proba

  // For each Probability
  for (let index in timelineProbaValues) {

    // Check if the Probability is in Withlist
    if (timelineProbaValues[index] == true) probabilityWithlist.push(probabilityList[index]);
  }

  return probabilityWithlist;
}

function RefreshProductionBudgetsStyle() {
  const sheet = productionBudgetSheet;
  const row = productionBudgetRow;
  const col = productionBudgetFirstCol;

  // Get Row
  const productionScopeRow = getRowFromName(sheet, productionStageName, row, getColumnFromA1(productionBudgetParam1Column), sheet.getLastRow() - row + 1);

  // Get Production Stages Datas
  const productionStages = cleanData(sheet.getRange(productionBudgetParam1Column + (productionScopeRow + 1) + ':' + productionBudgetParam1Column + (sheet.getLastRow())).getValues().filter(String));
  const productionStagesColors = GetColors(GetProductionRange());

  // Refresh Styles
  let currentCol = col;

  for (let productionStage of productionStages) {
    const productionColor = productionStagesColors[productionStages.indexOf(productionStage)]
    SetBudgetStyle(sheet, row, currentCol,productionColor);
    currentCol += 17;
  }
}

////////////////////////////
//     CUSTOM BUDGETS     //__________________________________________________________________________________________________________________________________________
////////////////////////////

function CreateCustomBudgetMenu() {
  const sheet = customBudgetSheet;
  const row = customBudgeRow;
  const probabilityRow = row +3;

  // Delete previous menu

  const menuRange = sheet.getRange(customBudgetParamTitleColumn + probabilityRow+ ':' + customBudgetParam3Column + sheet.getMaxRows())
  menuRange.clearContent();
  menuRange.clearDataValidations();
  menuRange.setBorder(true, true, true, true, true, true, GetDarkModeColor(), SpreadsheetApp.BorderStyle.SOLID);
  menuRange.breakApart();

  // Create Probability Selector
  CreateProbabilitySelection(sheet, customBudgetParamTitleColumn, probabilityRow);
  SetCustomBudgetHighlights()
}

function GetCustomBudgetProbabilityWithlist() {
  const sheet = customBudgetSheet;
  const row = customBudgeRow;
  const col = getColumnFromA1(customBudgetParamTitleColumn);
  const length = customBudgetLength;

  // Setup Data
  const setupProbabilityRange = GetProbabilityRange();
  const probabilityList = GetValues(setupProbabilityRange);
  const probabilityLength = probabilityList.length - 1;

  // Get Withlist Values
  const probabilityRow = getRowFromName(sheet, probabilityName, row, col, length);
  const timelineProbaRange = sheet.getRange(customBudgetParam2Column + probabilityRow + ':' + customBudgetParam2Column + (probabilityRow + probabilityLength - 1));
  const timelineProbaValues = cleanData(timelineProbaRange.getValues());

  probabilityWithlist = [probabilityList[0]];
  probabilityList.shift(); // Get Rid of the Certain Proba

  // Make the withlist
  for (let index in timelineProbaValues) {
    if (timelineProbaValues[index] == true) probabilityWithlist.push(probabilityList[index]);
  }
  return probabilityWithlist;
}

function CreateCustomBudget() {
  const sheet = customBudgetSheet;
  const row = customBudgeRow;
  const col = customBudgetFirstCol;
  const firstColumn = getA1FromCol(col)

  // Clean the sheet
  maxRange = sheet.getRange(getA1FromCol(col)+row+':'+ getA1FromCol(sheet.getMaxColumns())+ sheet.getMaxRows());
  maxRange.clearContent();
  maxRange.setBorder(true, true, true, true, true, true, GetDarkModeColor(), SpreadsheetApp.BorderStyle.SOLID);
  maxRange.setBackground(GetDarkModeColor());
  maxRange.breakApart()

  // Get Dates
  const startingDate = sheet.getRange(customBudgetParam1Column + row).getValue();
  const endingDate = sheet.getRange(customBudgetParam1Column + (row + 1)).getValue();

  // Get Budget Datas
  let budgetValues = GetBudgetValues(startingDate, endingDate, GetCustomBudgetProbabilityWithlist());
  let budgetTotals = budgetValues[1];
  budgetTotals = [budgetTotals[0][1], budgetTotals[0][2], budgetTotals[1][1], budgetTotals[1][2]];
  budgetValues = budgetValues[0];

  // Get Budget Dimensions
  const budgetWidth = budgetValues[0].length;
  const budgetLength = budgetValues.length;

  // Get Range
  const budgetRange = sheet.getRange(row, col, budgetLength, budgetWidth);

  // Set Values
  budgetRange.setValues(budgetValues);
  
  // Finalize Budget
  CreateBudgetChart(sheet, row, col, budgetTotals);
  SetBudgetStyle(sheet, row, col,getHighlightcolor1());

  // Correct Merging
  sheet.getRange(firstColumn+row+':'+getColumnOffset(firstColumn,3)+row).merge();
  sheet.getRange(getColumnOffset(firstColumn,4)+row+':'+getColumnOffset(firstColumn,7)+row).merge();
  sheet.getRange(firstColumn+row+':'+getColumnOffset(firstColumn,7)+row).setBorder(true,true,true,true,null,null,getHighlightcolor1(),SpreadsheetApp.BorderStyle.SOLID_THICK)
}

function RefreshCustomBudgetStyle() {
  const sheet = customBudgetSheet;
  const row = customBudgeRow;
  const col = customBudgetFirstCol;
  const firstColumn = getA1FromCol(col)
  
  SetBudgetStyle(sheet, row, col,getHighlightcolor1());
}






























