const crossCompare = () => {
  const ui = SpreadsheetApp.getUi();
  const spreadSheet = SpreadsheetApp.getActive();
  const sheet = SpreadsheetApp.getActiveSheet();

  const objects = [
    {
      title: "Give first column letter",
      text: "Give the colmun letter of the first column and click Yes.\r\n This column will be used in the comparison.",
      range: null,
      rangeAsText: "",
    },
    {
      title: "Give second column letter",
      text: "Give the column letter of the second column and click Yes.\r\n  This column will be used in the comparison",
      range: null,
      rangeAsText: "",
    },
  ];

  for (let i = 0; i < objects.length; i++) {
    const obj = objects[i];
    const response = ui.prompt(obj.title, obj.text, ui.ButtonSet.YES_NO);
    if (response.getSelectedButton() == ui.Button.NO) return;
    const responseText = response.getResponseText();
    if (typeof responseText !== "string") return;
    const trimmedText = responseText.trim().toUpperCase();
    if (!COLUMNLETTERS.find((val) => val === trimmedText)) return;
    const lastRow = sheet.getLastRow();
    obj.rangeAsText = `${trimmedText}1:${trimmedText}${lastRow}`;
    obj.range = sheet.getRange(obj.rangeAsText);
  }

  const useTheseToCreateComparison = ui.alert(
    `Use these?`,
    `Use these ranges for comparison?
    ${objects[0].rangeAsText}\r\n
    ${objects[1].rangeAsText}\r\n`,
    ui.ButtonSet.YES_NO
  );
  if (useTheseToCreateComparison == ui.Button.NO) return;

  const range0Arr = objects[0].range.getValues();
  const range1Arr = objects[1].range.getValues();
  const range0Obj = {};
  const range1Obj = {};

  const unionArr = [[`Found in both`]];
  const onlyCol0Arr = [[`Only in first column`]];
  const onlyCol1Arr = [[`Only in second column`]];
  const duplicate0Obj = {};
  const duplicate1Obj = {};

  for (let row = 1; row < range0Arr.length; row++) {
    const value = range0Arr[row][0];
    if (value != "") {
      if (range0Obj[value]) {
        if (duplicate0Obj[value]) {
          duplicate0Obj[value] = duplicate0Obj[value] + 1;
        } else {
          duplicate0Obj[value] = 1;
        }
      } else {
        range0Obj[value] = true;
      }
    }
  }
  for (let row = 1; row < range1Arr.length; row++) {
    const value = range1Arr[row][0];
    if (value != "") {
      if (range1Obj[value]) {
        if (duplicate1Obj[value]) {
          duplicate1Obj[value] = duplicate1Obj[value] + 1;
        } else {
          duplicate1Obj[value] = 1;
        }
      } else {
        range1Obj[value] = true;
      }
      if (range0Obj[value]) {
        unionArr.push([value]);
      } else {
        onlyCol1Arr.push([value]);
      }
    }
  }
  for (let row = 1; row < range0Arr.length; row++) {
    const value = range0Arr[row][0];
    if (value != "") {
      if (!range1Obj[value]) {
        onlyCol0Arr.push([value]);
      }
    }
  }

  let comparisonSheet = spreadSheet.getSheetByName(`ComparisonSheet`);
  if (!comparisonSheet) {
    comparisonSheet = spreadSheet.insertSheet(
      `ComparisonSheet`,
      sheet.getIndex() + 1
    );
  }

  comparisonSheet.clear();
  [unionArr, onlyCol0Arr, onlyCol1Arr].forEach((arr, index) => {
    comparisonSheet
      .getRange(1, index + 1, arr.length, arr[0].length)
      .setValues(arr);
  });

  const headingText = [`first`, `second`];

  [duplicate0Obj, duplicate1Obj].forEach((duplicateObj, index) => {
    const resultArr = [
      [`Duplicates in ${headingText[index]}`, `Number of duplicates`],
    ];
    Object.keys(duplicateObj).forEach((key) => {
      resultArr.push([key, duplicateObj[key]]);
    });
    comparisonSheet
      .getRange(1, (index + 1) * 3 + 2, resultArr.length, resultArr[0].length)
      .setValues(resultArr);
  });
  comparisonSheet.activate();
  const compSheetLastCol = comparisonSheet.getLastColumn();
  comparisonSheet
    .getRange(1, 1, 1, compSheetLastCol)
    .setFontSize("14")
    .setFontWeight("bold");
  comparisonSheet.autoResizeColumns(1, compSheetLastCol);
  showPopup({
    title: "Comparison created successfully!",
    messages: ["Comparison created successfully!"],
    file: null,
    folder: null,
    modal: false,
    height: 100,
    width: 200,
  });
};

const COLUMNLETTERS = [
  "", // Offset by one because array referencing value in array is base0
  "A",
  "B",
  "C",
  "D",
  "E",
  "F",
  "G",
  "H",
  "I",
  "J",
  "K",
  "L",
  "M",
  "N",
  "O",
  "P",
  "Q",
  "R",
  "S",
  "T",
  "U",
  "V",
  "W",
  "X",
  "Y",
  "Z",
  "AA",
  "AB",
  "AC",
  "AD",
  "AE",
  "AF",
  "AG",
  "AH",
  "AI",
  "AJ",
  "AK",
  "AL",
  "AM",
  "AN",
  "AO",
  "AP",
  "AQ",
  "AR",
  "AS",
  "AT",
  "AU",
  "AV",
  "AW",
  "AX",
  "AY",
  "AZ",
  "BA",
  "BB",
  "BC",
  "BD",
  "BE",
  "BF",
  "BG",
  "BH",
  "BI",
  "BJ",
  "BK",
  "BL",
  "BM",
  "BN",
  "BO",
  "BP",
  "BQ",
  "BR",
  "BS",
  "BT",
  "BU",
  "BV",
  "BW",
  "BX",
  "BY",
  "BZ",
  "CA",
  "CB",
  "CC",
  "CD",
  "CE",
  "CF",
  "CG",
  "CH",
  "CI",
  "CJ",
  "CK",
  "CL",
  "CM",
  "CN",
  "CO",
  "CP",
  "CQ",
  "CR",
  "CS",
  "CT",
  "CU",
  "CV",
  "CW",
  "CX",
  "CY",
  "CZ",
  "DA",
  "DB",
  "DC",
  "DD",
  "DE",
  "DF",
  "DG",
  "DH",
  "DI",
  "DJ",
  "DK",
  "DL",
  "DM",
  "DN",
  "DO",
  "DP",
  "DQ",
  "DR",
  "DS",
  "DT",
  "DU",
  "DV",
  "DW",
  "DX",
  "DY",
  "DZ",
  "EA",
  "EB",
  "EC",
  "ED",
  "EE",
  "EF",
  "EG",
  "EH",
  "EI",
  "EJ",
  "EK",
  "EL",
  "EM",
  "EN",
  "EO",
  "EP",
  "EQ",
  "ER",
  "ES",
  "ET",
  "EU",
  "EV",
  "EW",
  "EX",
  "EY",
  "EZ",
  "FA",
  "FB",
  "FC",
  "FD",
  "FE",
  "FF",
  "FG",
  "FH",
  "FI",
  "FJ",
  "FK",
  "FL",
  "FM",
  "FN",
  "FO",
  "FP",
  "FQ",
  "FR",
  "FS",
  "FT",
  "FU",
  "FV",
  "FW",
  "FX",
  "FY",
  "FZ",
  "GA",
  "GB",
  "GC",
  "GD",
  "GE",
  "GF",
  "GG",
  "GH",
  "GI",
  "GJ",
  "GK",
  "GL",
  "GM",
  "GN",
  "GO",
  "GP",
  "GQ",
  "GR",
  "GS",
  "GT",
  "GU",
  "GV",
  "GW",
  "GX",
  "GY",
  "GZ",
  "HA",
  "HB",
  "HC",
  "HD",
  "HE",
  "HF",
  "HG",
  "HH",
  "HI",
  "HJ",
  "HK",
  "HL",
  "HM",
  "HN",
  "HO",
  "HP",
  "HQ",
  "HR",
  "HS",
  "HT",
  "HU",
  "HV",
  "HW",
  "HX",
  "HY",
  "HZ",
  "IA",
  "IB",
  "IC",
  "ID",
  "IE",
  "IF",
  "IG",
  "IH",
  "II",
  "IJ",
  "IK",
  "IL",
  "IM",
  "IN",
  "IO",
  "IP",
  "IQ",
  "IR",
  "IS",
  "IT",
  "IU",
  "IV",
  "IW",
  "IX",
  "IY",
  "IZ",
  "JA",
  "JB",
  "JC",
  "JD",
  "JE",
  "JF",
  "JG",
  "JH",
  "JI",
  "JJ",
  "JK",
  "JL",
  "JM",
  "JN",
  "JO",
  "JP",
  "JQ",
  "JR",
  "JS",
  "JT",
  "JU",
  "JV",
  "JW",
  "JX",
  "JY",
  "JZ",
  "KA",
  "KB",
  "KC",
  "KD",
  "KE",
  "KF",
  "KG",
  "KH",
  "KI",
  "KJ",
  "KK",
  "KL",
  "KM",
  "KN",
  "KO",
  "KP",
  "KQ",
  "KR",
  "KS",
  "KT",
  "KU",
  "KV",
  "KW",
  "KX",
  "KY",
  "KZ",
  "LA",
  "LB",
  "LC",
  "LD",
  "LE",
  "LF",
  "LG",
  "LH",
  "LI",
  "LJ",
  "LK",
  "LL",
  "LM",
  "LN",
  "LO",
  "LP",
  "LQ",
  "LR",
  "LS",
  "LT",
  "LU",
  "LV",
  "LW",
  "LX",
  "LY",
  "LZ",
  "MA",
  "MB",
  "MC",
  "MD",
  "ME",
  "MF",
  "MG",
  "MH",
  "MI",
  "MJ",
  "MK",
  "ML",
  "MM",
  "MN",
  "MO",
  "MP",
  "MQ",
  "MR",
  "MS",
  "MT",
  "MU",
  "MV",
  "MW",
  "MX",
  "MY",
  "MZ",
  "NA",
  "NB",
  "NC",
  "ND",
  "NE",
  "NF",
  "NG",
  "NH",
  "NI",
  "NJ",
  "NK",
  "NL",
  "NM",
  "NN",
  "NO",
  "NP",
  "NQ",
  "NR",
  "NS",
  "NT",
  "NU",
  "NV",
  "NW",
  "NX",
  "NY",
  "NZ",
  "OA",
  "OB",
  "OC",
  "OD",
  "OE",
  "OF",
  "OG",
  "OH",
  "OI",
  "OJ",
  "OK",
  "OL",
  "OM",
  "ON",
  "OO",
  "OP",
  "OQ",
  "OR",
  "OS",
  "OT",
  "OU",
  "OV",
  "OW",
  "OX",
  "OY",
  "OZ",
  "PA",
  "PB",
  "PC",
  "PD",
  "PE",
  "PF",
  "PG",
  "PH",
  "PI",
  "PJ",
  "PK",
  "PL",
  "PM",
  "PN",
  "PO",
  "PP",
  "PQ",
  "PR",
  "PS",
  "PT",
  "PU",
  "PV",
  "PW",
  "PX",
  "PY",
  "PZ",
  "QA",
  "QB",
  "QC",
  "QD",
  "QE",
  "QF",
  "QG",
  "QH",
  "QI",
  "QJ",
  "QK",
  "QL",
  "QM",
  "QN",
  "QO",
  "QP",
  "QQ",
  "QR",
  "QS",
  "QT",
  "QU",
  "QV",
  "QW",
  "QX",
  "QY",
  "QZ",
  "RA",
  "RB",
  "RC",
  "RD",
  "RE",
  "RF",
  "RG",
  "RH",
  "RI",
  "RJ",
  "RK",
  "RL",
  "RM",
  "RN",
  "RO",
  "RP",
  "RQ",
  "RR",
  "RS",
  "RT",
  "RU",
  "RV",
  "RW",
  "RX",
  "RY",
  "RZ",
  "SA",
  "SB",
  "SC",
  "SD",
  "SE",
  "SF",
  "SG",
  "SH",
  "SI",
  "SJ",
  "SK",
  "SL",
  "SM",
  "SN",
  "SO",
  "SP",
  "SQ",
  "SR",
  "SS",
  "ST",
  "SU",
  "SV",
  "SW",
  "SX",
  "SY",
  "SZ",
  "TA",
  "TB",
  "TC",
  "TD",
  "TE",
  "TF",
  "TG",
  "TH",
  "TI",
  "TJ",
  "TK",
  "TL",
  "TM",
  "TN",
  "TO",
  "TP",
  "TQ",
  "TR",
  "TS",
  "TT",
  "TU",
  "TV",
  "TW",
  "TX",
  "TY",
  "TZ",
  "UA",
  "UB",
  "UC",
  "UD",
  "UE",
  "UF",
  "UG",
  "UH",
  "UI",
  "UJ",
  "UK",
  "UL",
  "UM",
  "UN",
  "UO",
  "UP",
  "UQ",
  "UR",
  "US",
  "UT",
  "UU",
  "UV",
  "UW",
  "UX",
  "UY",
  "UZ",
  "VA",
  "VB",
  "VC",
  "VD",
  "VE",
  "VF",
  "VG",
  "VH",
  "VI",
  "VJ",
  "VK",
  "VL",
  "VM",
  "VN",
  "VO",
  "VP",
  "VQ",
  "VR",
  "VS",
  "VT",
  "VU",
  "VV",
  "VW",
  "VX",
  "VY",
  "VZ",
  "WA",
  "WB",
  "WC",
  "WD",
  "WE",
  "WF",
  "WG",
  "WH",
  "WI",
  "WJ",
  "WK",
  "WL",
  "WM",
  "WN",
  "WO",
  "WP",
  "WQ",
  "WR",
  "WS",
  "WT",
  "WU",
  "WV",
  "WW",
  "WX",
  "WY",
  "WZ",
  "XA",
  "XB",
  "XC",
  "XD",
  "XE",
  "XF",
  "XG",
  "XH",
  "XI",
  "XJ",
  "XK",
  "XL",
  "XM",
  "XN",
  "XO",
  "XP",
  "XQ",
  "XR",
  "XS",
  "XT",
  "XU",
  "XV",
  "XW",
  "XX",
  "XY",
  "XZ",
  "YA",
  "YB",
  "YC",
  "YD",
  "YE",
  "YF",
  "YG",
  "YH",
  "YI",
  "YJ",
  "YK",
  "YL",
  "YM",
  "YN",
  "YO",
  "YP",
  "YQ",
  "YR",
  "YS",
  "YT",
  "YU",
  "YV",
  "YW",
  "YX",
  "YY",
  "YZ",
  "ZA",
  "ZB",
  "ZC",
  "ZD",
  "ZE",
  "ZF",
  "ZG",
  "ZH",
  "ZI",
  "ZJ",
  "ZK",
  "ZL",
  "ZM",
  "ZN",
  "ZO",
  "ZP",
  "ZQ",
  "ZR",
  "ZS",
  "ZT",
  "ZU",
  "ZV",
  "ZW",
  "ZX",
  "ZY",
  "ZZ",
];