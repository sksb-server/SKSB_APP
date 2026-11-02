# Staff_Manager_Template Rules Audit
**Workbook:** `Staff_Manager_Template.xltx`
**Generated:** 28-10-2026 14:13:11

---
## TABLE: Sheet=`All Staffs` | Table=`AllStaffsTable` | HeaderRow=1 | ArchetypeRow=2 | StartCol=1 | Columns=9

### Column Rules
| Col | Row | Address | Header | Style | NumberFormat | Type | InCellFormula | SourceList | SyncFormula | FilterExpr | TransformExpr | IsMulti | IsWarning | NativeType | NativeFormula | NativeAlert | NativeIgnoreBlank |
| :---: | :---: | :---: | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :---: | :---: | :--- | :--- | :---: | :---: |
| A | 2 | A2 | Staff ID | Unlocked | 0 | NativeValidation |   |   |   |   |   | FALSE | FALSE | WholeNumber | 0 | Stop | FALSE |
| B | 2 | B2 | Staff Name | Unlocked | @ | NativeValidation |   |   |   |   |   | FALSE | FALSE | CustomFormula | =LEN(TRIM(B2))>0 | Stop | FALSE |
| C | 2 | C2 | Department | Unlocked | @ | NativeValidation |   |   |   |   |   | FALSE | FALSE | List | =DepartmentList | Stop | FALSE |
| D | 2 | D2 | IC Number | Unlocked | @ | NativeValidation |   |   |   |   |   | FALSE | FALSE | CustomFormula | =OR(ISBLANK(D2), AND(ISNUMBER(VALUE(D2)), LEN(D2)=12)) | Warning | TRUE |
| E | 2 | E2 | Passport Number | Unlocked | @ | NativeValidation |   |   |   |   |   | FALSE | FALSE | CustomFormula | =OR(D2<>"", E2<>"") | Warning | TRUE |
| F | 2 | F2 | Phone Number | Unlocked | @ | StandardInput |   |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| G | 2 | G2 | Join Date | Unlocked | m/d/yyyy | NativeValidation |   |   |   |   |   | FALSE | FALSE | Date | 1/1/1990 | Stop | TRUE |
| H | 2 | H2 | Resign Date | Unlocked | m/d/yyyy | NativeValidation |   |   |   |   |   | FALSE | FALSE | CustomFormula | =OR(ISBLANK(H2), H2>=G2) | Stop | TRUE |
| I | 2 | I2 | Contract Status | DropDown | General | CustomDropdown |   | =IF(@left<>"", IF(@left<=TODAY(), "Resigned", ContractStatusList), ContractStatusList) |   | =IF(@item="Resigned", @left<>"", IF(@left<>"", @left>TODAY(), TRUE)) | @item | FALSE | TRUE |  |   |  | TRUE |

### Conditional Formatting
| AppliesTo | RuleType | Operator | Formula1 | Formula2 | Severity | StopIfTrue |
| :--- | :--- | :--- | :--- | :--- | :---: | :---: |
| $H$2:$I$2 | Expression |  | =AND($I2="Resigned", ISBLANK(H2)) |   | None | FALSE |
| $B$2:$C$2 | Expression |  | =AND($A2<>"", LEN(TRIM(B2))=0) |   | None | FALSE |
| $D$2:$E$2 | Expression |  | =AND($A2<>"", $D2="", $E2="") |   | None | FALSE |
| $A$2:$I$2 | Expression |  | =$I2="Suspended" |   | None | FALSE |
| $A$2:$I$2 | Expression |  | =$I2="Resigned" |   | None | FALSE |

---
## TABLE: Sheet=`Vehicle Licenses` | Table=`VehicleLicensesTable` | HeaderRow=1 | ArchetypeRow=2 | StartCol=1 | Columns=8

### Column Rules
| Col | Row | Address | Header | Style | NumberFormat | Type | InCellFormula | SourceList | SyncFormula | FilterExpr | TransformExpr | IsMulti | IsWarning | NativeType | NativeFormula | NativeAlert | NativeIgnoreBlank |
| :---: | :---: | :---: | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :---: | :---: | :--- | :--- | :---: | :---: |
| A | 2 | A2 | Vehicle License ID | DropDown | 0 | CustomDropdown |   | =MAX(VehicleLicensesTable[Vehicle License ID])+1 |   |   |   | FALSE | TRUE |  |   |  | TRUE |
| B | 2 | B2 | Staff ID | DropDown | @ | CustomDropdown |   | =AllStaffsTable[Staff ID] | @right LET(raw, XLOOKUP(@item, AllStaffsTable[Staff ID], AllStaffsTable[Staff Name], ""), p, TEXTSPLIT(TRIM(raw), " "), n, COUNTA(p), IF(raw="", "", IF(n<=3, raw, INDEX(p, 1) & " " & INDEX(p, 2) & " " & INDEX(p, n)))) | IF(@right<>"", XLOOKUP(@item, AllStaffsTable[Staff ID], AllStaffsTable[Staff Name], "") = @right, XLOOKUP(@item, AllStaffsTable[Staff ID], AllStaffsTable[Contract Status], "") = "Active") | @item | FALSE | FALSE |  |   |  | TRUE |
| C | 2 | C2 | Staff Name | DropDown | @ | CustomDropdown |   | =AllStaffsTable[Staff Name] | @left IFERROR(XLOOKUP(@item, AllStaffsTable[Staff Name], AllStaffsTable[Staff ID]), XLOOKUP(@display, AllStaffsTable[Staff Name], AllStaffsTable[Staff ID], "")) | IF(@left<>"", XLOOKUP(@left, AllStaffsTable[Staff ID], AllStaffsTable[Staff Name], "") = @item, XLOOKUP(@item, AllStaffsTable[Staff Name], AllStaffsTable[Contract Status], "") = "Active") | LET(raw, TRIM(@item), p, TEXTSPLIT(raw, " "), n, COUNTA(p), IF(raw="", "", IF(n<=3, raw, INDEX(p, 1) & " " & INDEX(p, 2) & " " & INDEX(p, n)))) | FALSE | FALSE |  |   |  | TRUE |
| D | 2 | D2 | Vehicle License Class | Unlocked | General | NativeValidation |   |   |   |   |   | FALSE | FALSE | List | =VehicleLicenseClassList | Stop | FALSE |
| E | 2 | E2 | Renew Date | DropDown | m/d/yyyy | CustomDropdown |   | =LET(wid, @left, prevCount, COUNTIF(VehicleLicensesTable[Staff ID], wid), joinD, XLOOKUP(wid, AllStaffsTable[Staff ID], AllStaffsTable[Join Date], ""), IF(prevCount<=1, IF(joinD<>"", joinD, TODAY()), TODAY())) |   |   |   | FALSE | TRUE |  |   |  | TRUE |
| F | 2 | F2 | Expiry Date | DropDown | m/d/yyyy | CustomDropdown |   | =IF(@left<>"", @left+365, TODAY()+365) |   |   |   | FALSE | TRUE |  |   |  | TRUE |
| G | 2 | G2 | Status | DropDown | General | CustomDropdown |   | =IF(@left <> "", IF(@left < TODAY(), "Invalid", VehicleLicenseStatusList), VehicleLicenseStatusList) |   | =IF(@item = "Valid", IF(@left <> "", @left >= TODAY(), TRUE), TRUE) | @item | FALSE | TRUE |  |   |  | TRUE |
| H | 2 | H2 | Cost | Unlocked | 0.00 | StandardInput |   |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |

### Conditional Formatting
| AppliesTo | RuleType | Operator | Formula1 | Formula2 | Severity | StopIfTrue |
| :--- | :--- | :--- | :--- | :--- | :---: | :---: |
| $F$2 | Type_6 |  |   |   | None | FALSE |
| $F$2:$G$2 | Expression |  | =AND($F2<>"", TODAY()>$F2, J2="Valid") |   | None | FALSE |
| $B$2:$D$2 | Expression |  | =AND(LEN(TRIM(CONCAT($A2:$G2)))>0, LEN(TRIM($B2))=0) |   | None | FALSE |
| $F$2 | Expression |  | =AND($E2<>"", $F2<>"", $F2<=$E2) |   | None | FALSE |
| $E$2:$G$2,$A$2 | Expression |  | =AND($B2<>"", ISBLANK(A2)) |   | None | FALSE |
| $A$2:$G$2 | Expression |  | =$G2="Suspended" |   | None | FALSE |
| $A$2:$G$2 | Expression |  | =$G2="Invalid" |   | None | FALSE |

---
## TABLE: Sheet=`Skill Licenses` | Table=`SkillLicensesTable` | HeaderRow=1 | ArchetypeRow=2 | StartCol=1 | Columns=8

### Column Rules
| Col | Row | Address | Header | Style | NumberFormat | Type | InCellFormula | SourceList | SyncFormula | FilterExpr | TransformExpr | IsMulti | IsWarning | NativeType | NativeFormula | NativeAlert | NativeIgnoreBlank |
| :---: | :---: | :---: | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :---: | :---: | :--- | :--- | :---: | :---: |
| A | 2 | A2 | Skill License ID | DropDown | 0 | CustomDropdown |   | =MAX(SkillLicensesTable[Skill License ID])+1 |   |   |   | FALSE | TRUE |  |   |  | TRUE |
| B | 2 | B2 | Staff ID | DropDown | @ | CustomDropdown |   | =AllStaffsTable[Staff ID] | @right LET(raw, XLOOKUP(@item, AllStaffsTable[Staff ID], AllStaffsTable[Staff Name], ""), p, TEXTSPLIT(TRIM(raw), " "), n, COUNTA(p), IF(raw="", "", IF(n<=3, raw, INDEX(p, 1) & " " & INDEX(p, 2) & " " & INDEX(p, n)))) | IF(@right<>"", XLOOKUP(@item, AllStaffsTable[Staff ID], AllStaffsTable[Staff Name], "") = @right, XLOOKUP(@item, AllStaffsTable[Staff ID], AllStaffsTable[Contract Status], "") = "Active") | @item | FALSE | FALSE |  |   |  | TRUE |
| C | 2 | C2 | Staff Name | DropDown | @ | CustomDropdown |   | =AllStaffsTable[Staff Name] | IFERROR(XLOOKUP(@item, AllStaffsTable[Staff Name], AllStaffsTable[Staff ID]), IFERROR(XLOOKUP(@display, AllStaffsTable[Staff Name], AllStaffsTable[Staff ID]), "")) | IF(@left<>"", XLOOKUP(@left, AllStaffsTable[Staff ID], AllStaffsTable[Staff Name], "") = @item, XLOOKUP(@item, AllStaffsTable[Staff Name], AllStaffsTable[Contract Status], "") = "Active") | LET(raw, TRIM(@item), p, TEXTSPLIT(raw, " "), n, COLUMNS(p), IF(raw="", "", IF(n<=3, raw, INDEX(p, 1, 1) & " " & INDEX(p, 1, 2) & " " & INDEX(p, 1, n)))) | FALSE | FALSE |  |   |  | TRUE |
| D | 2 | D2 | Skill License Type | Unlocked | @ | NativeValidation |   |   |   |   |   | FALSE | FALSE | List | =INDIRECT("SkillLicenseTypesTable[Skill License Types]") | Stop | TRUE |
| E | 2 | E2 | Renew Date | DropDown | m/d/yyyy | CustomDropdown |   | =LET(wid, @left, prevCount, COUNTIF(VehicleLicensesTable[Staff ID], wid), joinD, XLOOKUP(wid, AllStaffsTable[Staff ID], AllStaffsTable[Join Date], ""), IF(prevCount<=1, IF(joinD<>"", joinD, TODAY()), TODAY())) |   |   |   | FALSE | TRUE |  |   |  | TRUE |
| F | 2 | F2 | Expiry Date | DropDown | m/d/yyyy | CustomDropdown |   | =IF(@left<>"", @left+365, TODAY()+365) |   |   |   | FALSE | TRUE |  |   |  | TRUE |
| G | 2 | G2 | Status | DropDown | General | CustomDropdown |   | =IF(@left <> "", IF(@left < TODAY(), "Invalid", SkillLicenseStatusList), SkillLicenseStatusList) |   | =IF(@item = "Valid", IF(@left <> "", @left >= TODAY(), TRUE), TRUE) | @item | FALSE | TRUE |  |   |  | TRUE |
| H | 2 | H2 | Cost | DropDown | 0.00 | CustomDropdown |   | =SkillLicensesTable[Cost] |   |   |   | FALSE | FALSE |  |   |  | TRUE |

### Conditional Formatting
| AppliesTo | RuleType | Operator | Formula1 | Formula2 | Severity | StopIfTrue |
| :--- | :--- | :--- | :--- | :--- | :---: | :---: |
| $F$2 | Type_6 |  |   |   | None | FALSE |
| $F$2:$G$2 | Expression |  | =AND($F2<>"", TODAY()>$F2, F2="Valid") |   | None | FALSE |
| $B$2:$D$2 | Expression |  | =AND(LEN(TRIM(CONCAT($A2:$G2)))>0, LEN(TRIM($B2))=0) |   | None | FALSE |
| $F$2 | Expression |  | =AND($E2<>"", $F2<>"", $F2<=$E2) |   | None | FALSE |
| $E$2:$H$2,$A$2 | Expression |  | =AND($B2<>"", ISBLANK(A2)) |   | None | FALSE |
| $A$2:$H$2 | Expression |  | =$G2="Suspended" |   | None | FALSE |
| $A$2:$H$2 | Expression |  | =$G2="Invalid" |   | None | FALSE |

---
## TABLE: Sheet=`Skill Licenses` | Table=`SkillLicenseTypesTable` | HeaderRow=1 | ArchetypeRow=2 | StartCol=10 | Columns=1

### Column Rules
| Col | Row | Address | Header | Style | NumberFormat | Type | InCellFormula | SourceList | SyncFormula | FilterExpr | TransformExpr | IsMulti | IsWarning | NativeType | NativeFormula | NativeAlert | NativeIgnoreBlank |
| :---: | :---: | :---: | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :---: | :---: | :--- | :--- | :---: | :---: |
| J | 2 | J2 | Skill License Types | Unlocked | General | StandardInput |   |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |

