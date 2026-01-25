/**
 * Template Data - Single source of truth for sheet structure.
 * Sheet names must match original Excel EXACTLY (including typos, leading/trailing spaces).
 */
const TEMPLATE_DATA = {
  workbooks: {
    ModelData: {
      name: "Model Data",
      description: "Basic Settings",
      color: "#2563eb",
      filename: "ModelData",
      uiGroups: [
        { key: "basic", label: "Basic", sheets: ["Company", "Business Unit"] },
        { key: "m1", label: "Module 1", sheets: ["Company Resource"] },
        { key: "m2", label: "Module 2", sheets: ["Activity Center", " Normal Capacity"] },
        { key: "m3", label: "Module 3", sheets: ["Activity", "Driver and Allocation Formula", "Machine(Activity Center Driver)"] },
        { key: "m4", label: "Module 4", sheets: ["Material", "ProductProject", "Product", "Customer", "Service Driver"] }
      ]
    },
    PeriodData: {
      name: "Period Data",
      description: "Monthly Data",
      color: "#059669",
      filename: "PeriodData",
      uiGroups: [
        { key: "basic", label: "Basic Info", sheets: ["Exchange Rate"] },
        { key: "m1", label: "Module 1", sheets: ["Resource", "Resource Driver(Actvity Center)", "Resource Driver(Value Object)", "Resource Driver(Machine)", "Resource Driver(M. A. C.)", "Resource Driver(S. A. C.)"] },
        { key: "m2", label: "Module 2", sheets: ["Activity Center Driver(N. Cap.)", "Activity Center Driver(A. Cap.)"] },
        { key: "m3", label: "Module 3", sheets: ["Activity Driver", "ProductProject Driver", "Manufacture Order"] },
        { key: "m4", label: "Module 4", sheets: ["Manufacture Material", "Purchased Material and WIP", "Expected Project Value", "Sales Revenue ", "Revenue(InternalTransaction)NA", "Service Driver_Period", "Std. Workhour", "Std. Material(BOM)"] }
      ]
    }
  },

  sheets: {
    // ========================================
    // MODEL DATA WORKBOOK - Hidden/System
    // ========================================
    "Remark": {
      workbook: "ModelData",
      hidden: true,
      exportAsBlank: true,
      headers: [],
      data: []
    },
    "Item<IT使用>": {
      workbook: "ModelData",
      hidden: true,
      sheetNameInExcel: "Item<IT使用>",
      exportUseTemplate: true,
      headers: ["No.", "中文", "英文", "工作表名稱", "Remark", "需求", "條件說明"],
      data: [
        ["1", "公司", "Company", "Company", "", "Required", ""],
        ["2", "事業單位", "Business Unit", "Business Unit", "", "Required", ""],
        ["3", "公司資源", "Company Resource", "Company Resource", "", "Required", ""],
        ["4", "作業中心", "Activity Center", "Activity Center", "", "Required", ""],
        ["5", "機台(作業中心動因)", "Machine(Activity Center Driver)", "Machine(Activity Center Driver)", "", "Required", "製造業 Required；其它產業 Optional"],
        ["6", "機台(替代動因)", "Machine(Substitute Driver)", "Machine(Substitute Driver)", "待確認", "Optional", "設計非主要機台"],
        ["7", "正常產能", "Normal Capacity", "Normal Capacity", "", "Optional", "正常產能法"],
        ["8", "作業", "Activity", "Activity", "", "Required", ""],
        ["9", "產品專案", "ProductProject", "ProductProject", "", "Optional", "設計產品專案為價值標的"],
        ["10", "原料", "Material", "Material", "", "Required", ""],
        ["11", "產品", "Product", "Product", "", "Required", ""],
        ["12", "顧客", "Customer", "Customer", "", "Required", ""],
        ["13", "標籤", "Label", "Label", "", "Required", ""],
        ["14", "清單設定", "List setting", "List setting", "", "Required", ""],
        ["15", "服務動因", "Service Driver", "Service Driver", "", "Required", ""],
        ["16", "動因公式", "Driver and Allocation Formula", "Driver and Allocation Formula", "", "Required", ""]
      ]
    },
    "TableMapping<IT使用>": {
      workbook: "ModelData",
      hidden: true,
      exportAsBlank: true,
      headers: [],
      data: []
    },
    "List_Item<IT使用>": {
      workbook: "ModelData",
      hidden: true,
      exportAsBlank: true,
      headers: [],
      data: []
    },

    // ========================================
    // MODEL DATA WORKBOOK - User-editable
    // ========================================
    "Company": {
      workbook: "ModelData",
      headers: ["Company", "Description", "Currency", "Resource Level", "Activity Center Level", "Activity Level"],
      data: [["", "", "", "", "", ""]],
      required: ["Company", "Description", "Currency", "Resource Level", "Activity Center Level", "Activity Level"]
    },
    "Business Unit": {
      workbook: "ModelData",
      headers: ["Business Unit", "Description", "Currency", "Region"],
      data: [["", "", "", ""]],
      required: ["Business Unit", "Currency"]
    },
    "Company Resource": {
      workbook: "ModelData",
      headers: ["Resource Code", "Resource - Level 1", "Resource - Level 2", "Resource - Level 3", "Resource - Level 4", "Resource - Level 5", "Description", "A.C. or Value Object Type", "Resource Driver", "Product Cost Type"],
      data: [["", "", "", "", "", "", "", "", "", ""]],
      required: ["Resource Code", "Resource - Level 1", "Resource - Level 2", "Description", "A.C. or Value Object Type", "Resource Driver"]
    },
    "Activity Center": {
      workbook: "ModelData",
      headers: ["Business Unit", "Activity Center Code 1", "Description 1", "Activity Center Code 2", "Description 2", "Activity Center Code 3", "Description 3", "Activity Center Code 4", "Description 4", "Activity Center Code 5", "Description 5", "Allocation", "ABC-Implemented", "Sales Revenue"],
      data: [["", "", "", "", "", "", "", "", "", "", "", "", "", ""]],
      required: ["Business Unit", "Activity Center Code 1"]
    },
    "List setting": {
      workbook: "ModelData",
      headers: ["List Name", "List Description", "List Item", "List Item Description"],
      data: [["", "", "", ""]],
      required: ["List Name", "List Item"]
    },
    "Label": {
      workbook: "ModelData",
      headers: ["Entity", "System Property", "Label", "List Name", "Default Value", "AllowFreeText", "AllowNull"],
      data: [["", "", "", "", "", "", ""]],
      required: ["Entity", "System Property"]
    },
    "Driver and Allocation Formula": {
      workbook: "ModelData",
      headers: ["Entity", "Driver Name or Allocation", "Description"],
      data: [["", "", ""]],
      required: ["Entity", "Driver Name or Allocation"]
    },
    "Machine(Activity Center Driver)": {
      workbook: "ModelData",
      headers: ["Activity Center Code", "Machine Code", "Machine Quantity", "Machine Name"],
      data: [["", "", "", ""]],
      required: ["Activity Center Code", "Machine Code"]
    },
    " Normal Capacity": {
      workbook: "ModelData",
      headers: ["Activity Code", "Activity Name", "Description"],
      data: [["", "", ""]],
      required: ["Activity Code"]
    },
    "Activity": {
      workbook: "ModelData",
      headers: ["Activity Code", "Activity - Level 1", "Activity - Level 2", "Activity - Level 3", "Activity - Level 4", "Activity Name", "Description", "Activity Driver", "Quality Attribute", "Customer Service Attribute", "Productivity Attribute", "Value-added Attribute", "Reason Group", "Value Object Type", "Product Cost Type"],
      data: [["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]],
      required: ["Activity Code", "Activity Name"]
    },
    "Material": {
      workbook: "ModelData",
      headers: ["Material Code", "Description"],
      data: [["", ""]],
      required: ["Material Code"]
    },
    "ProductProject": {
      workbook: "ModelData",
      headers: ["Activity Center", "Project Code", "Description", "Project Driver"],
      data: [["", "", "", ""]],
      required: ["Activity Center", "Project Code"]
    },
    "Product": {
      workbook: "ModelData",
      headers: ["Product Code", "Description"],
      data: [["", ""]],
      required: ["Product Code"]
    },
    "Customer": {
      workbook: "ModelData",
      headers: ["Customer Code", "Description"],
      data: [["", ""]],
      required: ["Customer Code"]
    },
    "Service Driver": {
      workbook: "ModelData",
      headers: ["Entity", "Code", "Service Driver"],
      data: [["", "", ""]],
      required: ["Entity", "Code"]
    },

    // ========================================
    // PERIOD DATA WORKBOOK - Hidden/System
    // ========================================
    "工作表2": {
      workbook: "PeriodData",
      hidden: true,
      exportAsBlank: true,
      headers: [],
      data: []
    },
    "Sheet2": {
      workbook: "PeriodData",
      hidden: true,
      exportAsBlank: true,
      headers: [],
      data: []
    },
    "Item": {
      workbook: "PeriodData",
      hidden: true,
      exportAsBlank: true,
      headers: [],
      data: []
    },
    "TableMapping": {
      workbook: "PeriodData",
      hidden: true,
      exportAsBlank: true,
      headers: [],
      data: []
    },

    // ========================================
    // PERIOD DATA WORKBOOK - User-editable
    // ========================================
    "Exchange Rate": {
      workbook: "PeriodData",
      headers: ["Business Unit Currency", "Company Currency", "Exchange Rate"],
      data: [["", "", ""], ["", "", ""], ["", "", ""]],
      required: ["Business Unit Currency", "Company Currency", "Exchange Rate"]
    },
    "Resource": {
      workbook: "PeriodData",
      headers: ["Business Unit", "Resource Code", "Activity Center Code", "Amount", "Value Object Type", "Value Object Code", "Machine Code", "Product Code"],
      data: [["", "", "", "", "", "", "", ""]],
      required: ["Business Unit", "Resource Code", "Amount"]
    },
    "Resource Driver(Actvity Center)": {
      workbook: "PeriodData",
      headers: ["Activity Center Code", "(Activity Center)", "Driver Code 1", "Driver Code 2"],
      data: [["", "", "", ""]],
      required: ["Activity Center Code"]
    },
    "Resource Driver(Value Object)": {
      workbook: "PeriodData",
      headers: ["Business Unit", "Value Object Type", "Value Object Code", "Driver Code 1", "Driver Code 2"],
      data: [["", "", "", "", ""]],
      required: ["Business Unit", "Value Object Type", "Value Object Code"]
    },
    "Resource Driver(Machine)": {
      workbook: "PeriodData",
      headers: ["Activity Center Code", "Machine Code", "Driver Code 1", "Driver Code 2"],
      data: [["", "", "", ""]],
      required: ["Activity Center Code", "Machine Code"]
    },
    "Resource Driver(M. A. C.)": {
      workbook: "PeriodData",
      headers: ["Activity Center Code", "Machine Code", "Driver 1", "Driver 2", "Driver 3"],
      data: [["", "", "", "", ""]],
      required: ["Activity Center Code", "Machine Code"]
    },
    "Resource Driver(S. A. C.)": {
      workbook: "PeriodData",
      headers: ["Activity Center Code", "Machine Code", "Driver 1", "Driver 2", "Driver 3", "Driver 4", "Driver 5", "Driver 6"],
      data: [["", "", "", "", "", "", "", ""]],
      required: ["Activity Center Code"]
    },
    "Activity Center Driver(N. Cap.)": {
      workbook: "PeriodData",
      headers: ["Activity Center Code", "Machine Code", "Activity Code", "Normal Capacity Hours"],
      data: [["", "", "", ""]],
      required: ["Activity Center Code", "Activity Code", "Normal Capacity Hours"]
    },
    "Activity Center Driver(A. Cap.)": {
      workbook: "PeriodData",
      headers: ["Activity Center Code", "Machine Code", "Supported Activity Center Code", "Activity Code", "Actual Capacity Hours", "Value Object Code", "Value Object Type", "Product Code"],
      data: [["", "", "", "", "", "", "", ""]],
      required: ["Activity Center Code", "Activity Code", "Actual Capacity Hours"]
    },
    "Activity Driver": {
      workbook: "PeriodData",
      headers: ["Activity Center Code", "Machine Code", "Activity Code", "Activity Driver", "Activity Driver Value", "Value Object Code", "Value Object Type", "Product Code"],
      data: [["", "", "", "", "", "", "", ""]],
      required: ["Activity Center Code", "Activity Code", "Activity Driver", "Activity Driver Value"]
    },
    "ProductProject Driver": {
      workbook: "PeriodData",
      headers: ["Product Code", "Project Driver", "Project Driver Value"],
      data: [["", "", ""]],
      required: ["Product Code", "Project Driver", "Project Driver Value"]
    },
    "Manufacture Order": {
      workbook: "PeriodData",
      headers: ["Business Unit", "MO", "Product Code", "Quantity", "Closed"],
      data: [["", "", "", "", ""]],
      required: ["Business Unit", "MO", "Product Code", "Quantity"]
    },
    "Manufacture Material": {
      workbook: "PeriodData",
      headers: ["Business Unit", "MO", "Material Code", "Quantity", "Amount", "Note"],
      data: [["", "", "", "", "", ""]],
      required: ["Business Unit", "MO", "Material Code", "Quantity"]
    },
    "Purchased Material and WIP": {
      workbook: "PeriodData",
      headers: ["Business Unit", "Material Code", "Quantity", "Amount", "End Inventory Qty", "Unit", "End Inventory Amount"],
      data: [["", "", "", "", "", "", ""]],
      required: ["Business Unit", "Material Code"]
    },
    "Expected Project Value": {
      workbook: "PeriodData",
      headers: ["Project Code", "Total Project Driver Value"],
      data: [["", ""]],
      required: ["Project Code", "Total Project Driver Value"]
    },
    "Sales Revenue ": {
      workbook: "PeriodData",
      headers: ["Order No", "Customer Code", "Product Code", "Quantity", "Amount", "Sales Activity Center Code", "Shipment Business Unit", "Currency"],
      data: [["", "", "", "", "", "", "", ""]],
      required: ["Order No", "Customer Code", "Product Code", "Quantity", "Amount"]
    },
    "Revenue(InternalTransaction)NA": {
      workbook: "PeriodData",
      headers: ["Business Unit", "Resource Code", "Activity Center Code", "Amount", "Supported Activity Center Code", "Product Code", "Quantity"],
      data: [["", "", "", "", "", "", ""]],
      required: ["Business Unit", "Resource Code", "Amount"]
    },
    "Service Driver_Period": {
      workbook: "PeriodData",
      sheetNameInExcel: "Service Driver",
      headers: ["Business Unit", "Customer Code", "Product Code"],
      data: [["", "", ""]],
      required: ["Business Unit"]
    },
    "Std. Workhour": {
      workbook: "PeriodData",
      headers: ["Business Unit", "Product", "Activity Code", "Std. Work Hour", "Std. Machine Hour"],
      data: [["", "", "", "", ""]],
      required: ["Business Unit", "Product", "Activity Code"]
    },
    "Std. Material(BOM)": {
      workbook: "PeriodData",
      headers: ["Business Unit", "Product", "Material", "Std. Material Quantity"],
      data: [["", "", "", ""]],
      required: ["Business Unit", "Product", "Material", "Std. Material Quantity"]
    }
  },

  sheetOrder: {
    ModelData: [
      "Company",
      "Business Unit",
      "Company Resource",
      "Activity Center",
      "List setting",
      "Label",
      "Driver and Allocation Formula",
      "Machine(Activity Center Driver)",
      " Normal Capacity",
      "Activity",
      "Material",
      "ProductProject",
      "Product",
      "Customer",
      "Service Driver"
    ],
    PeriodData: [
      "Exchange Rate",
      "Resource",
      "Resource Driver(Actvity Center)",
      "Resource Driver(Value Object)",
      "Resource Driver(Machine)",
      "Resource Driver(M. A. C.)",
      "Resource Driver(S. A. C.)",
      "Activity Center Driver(N. Cap.)",
      "Activity Center Driver(A. Cap.)",
      "Activity Driver",
      "ProductProject Driver",
      "Manufacture Order",
      "Manufacture Material",
      "Purchased Material and WIP",
      "Expected Project Value",
      "Sales Revenue ",
      "Revenue(InternalTransaction)NA",
      "Service Driver_Period",
      "Std. Workhour",
      "Std. Material(BOM)"
    ]
  }
};

function getSheetsForWorkbook(workbookKey) {
  return TEMPLATE_DATA.sheetOrder[workbookKey] || [];
}

function getSheetConfig(sheetName) {
  return TEMPLATE_DATA.sheets[sheetName];
}

function norm(s) { return String(s).trim().replace(/\s+/g, " ").replace(/[–—]/g, "-"); }

function isRequired(sheetName, columnName) {
  const config = TEMPLATE_DATA.sheets[sheetName];
  if (!config || !config.required) return false;
  return config.required.some(function (r) { return norm(r) === norm(columnName); });
}

function getExcelSheetName(internalName) {
  const config = TEMPLATE_DATA.sheets[internalName];
  return (config && config.sheetNameInExcel) ? config.sheetNameInExcel : internalName;
}

function getInternalSheetName(excelName, workbookKey) {
  for (const [internalName, config] of Object.entries(TEMPLATE_DATA.sheets)) {
    if (config.workbook === workbookKey) {
      const excelSheetName = config.sheetNameInExcel || internalName;
      if (excelSheetName === excelName) {
        return internalName;
      }
    }
  }
  return null;
}
