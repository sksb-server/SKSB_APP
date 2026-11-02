# Vehicle_Manager_Template1 Rules Audit
**Workbook:** `Vehicle_Manager_Template1`
**Generated:** 30-10-2026 12:09:58

---
## SHEET_HEADER_CELLS: Sheet=`All Vehicles`

| CellAddress | Row | Col | Style | NumberFormat | Formula | SourceList / NativeFormula | NativeType | NativeAlert | NativeIgnoreBlank |
| :---: | :---: | :---: | :--- | :--- | :--- | :--- | :--- | :---: | :---: |
| $B$1 | 1 | 2 | DisplayDate | m/d/yyyy | =TODAY() | =DATE(1990,1,1) | Date | Stop | TRUE |
| $C$1 | 1 | 3 | DisplayRange | @ | All_Time | =DateFilterScopeList | List | Stop | TRUE |

---
## TABLE: Sheet=`All Vehicles` | Table=`AllVehiclesTable` | HeaderRow=2 | ArchetypeRow=3 | StartCol=1 | Columns=25

### Column Rules
| Col | Row | Address | Header | Style | NumberFormat | Type | InCellFormula | SourceList | SyncFormula | FilterExpr | TransformExpr | IsMulti | IsWarning | NativeType | NativeFormula | NativeAlert | NativeIgnoreBlank |
| :---: | :---: | :---: | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :---: | :---: | :--- | :--- | :---: | :---: |
| A | 3 | A3 | Vehicle Plate | Unlocked | @ | StandardInput | BKK 9965 |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| B | 3 | B3 | Vehicle Type | DropDown | @ | CustomDropdown | Lorry_3T | =VehicleTypeList |   |   | @item | FALSE | FALSE |  |   |  | TRUE |
| C | 3 | C3 | Status | DropDown | @ | CustomDropdown | Active | =IF([@[Sold Date]] <> "", "Decommissioned", IF(OR(AND([@[Road Tax Expiry]] <> "—", [@[Road Tax Expiry]] < TODAY()), AND([@[Inspection Expiry]] <> "—", [@[Inspection Expiry]] < TODAY())), "Expired", VehicleStatusList)) |   | =IF(@item = "Decommissioned", [@[Sold Date]] <> "", IF(@item = "Active", AND(OR([@[Road Tax Expiry]] = "—", [@[Road Tax Expiry]] >= TODAY()), OR([@[Inspection Expiry]] = "—", [@[Inspection Expiry]] >= TODAY())), TRUE)) | @item | FALSE | TRUE |  |   |  | TRUE |
| D | 3 | D3 | VIN | Unlocked | @ | StandardInput | PL3T9965X001 |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| E | 3 | E3 | Brand | Unlocked | @ | StandardInput | Isuzu |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| F | 3 | F3 | Model | Unlocked | @ | StandardInput | NPR 75 |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| G | 3 | G3 | Years Old | Unlocked | General | Formula | =IF([@[Bought Date]]="", "—", IF(AND([@Status]="Decommissioned", [@[Sold Date]]<>""), DATEDIF([@[Bought Date]], [@[Sold Date]], "Y"), DATEDIF([@[Bought Date]], TODAY(), "Y"))) |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| H | 3 | H3 | Insurance Expiry | Normal | m/d/yyyy | Formula | =IFERROR(MAX(FILTER(InsurancesTable[Expiry Date], InsurancesTable[Vehicle Plate]=[@[Vehicle Plate]])), "—") |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| I | 3 | I3 | Permit Expiry | Normal | m/d/yyyy | Formula | =IFERROR(MIN(FILTER(PermitsTable[Expiry Date], PermitsTable[Vehicle Plate]=[@[Vehicle Plate]])), "—") |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| J | 3 | J3 | Road Tax Expiry | Normal | m/d/yyyy | Formula | =IFERROR(MAX(FILTER(RoadTaxesTable[Expiry Date], RoadTaxesTable[Vehicle Plate]=[@[Vehicle Plate]])), "—") |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| K | 3 | K3 | Inspection Expiry | Normal | m/d/yyyy | Formula | =IFERROR(MAX(FILTER(InspectionsTable[Expiry Date], InspectionsTable[Vehicle Plate]=[@[Vehicle Plate]])), "—") |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| L | 3 | L3 | Current Odometer | Normal | 0 | Formula | =IFERROR(MAX(FILTER(OdometersTable[Current Odometer], OdometersTable[Vehicle Plate]=[@[Vehicle Plate]])), 0) |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| M | 3 | M3 | Bought Date | Unlocked | m/d/yyyy | NativeValidation | 44270 |   |   |   |   | FALSE | FALSE | Date | =DATE(1990,1,1) | Stop | TRUE |
| N | 3 | N3 | Sold Date | Unlocked | m/d/yyyy | NativeValidation |   |   |   |   |   | FALSE | FALSE | Date | =M3 | Stop | TRUE |
| O | 3 | O3 | Bought Price | Unlocked | 0 | NativeValidation |   |   |   |   |   | FALSE | FALSE | Decimal | 0 | Stop | TRUE |
| P | 3 | P3 | Sold Price | Unlocked | 0 | NativeValidation |   |   |   |   |   | FALSE | FALSE | Decimal | 0 | Stop | TRUE |
| Q | 3 | Q3 | Driven By | Unuse | @ | NativeValidation |   |   |   |   |   | FALSE | FALSE | CustomFormula | FALSE | Stop | TRUE |
| R | 3 | R3 | Used At | Unuse | @ | NativeValidation |   |   |   |   |   | FALSE | FALSE | CustomFormula | FALSE | Stop | TRUE |
| S | 3 | S3 | Return | Normal | #,##0.00 | Formula | =IF([@[Sold Price]]>0, [@[Sold Price]] - [@[Bought Price]], 0) |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| T | 3 | T3 | Total Cost | Normal | #,##0.00 | Formula | =SUM([@[Fuel Cost]]:[@[Road Tax and Inspection Cost]]) |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| U | 3 | U3 | Fuel Cost | Normal | #,##0.00 | Formula | =LET(vPlate, [@[Vehicle Plate]], anchorD, $B$1, scope, $C$1, SUM(FILTER(OdometersTable[Transaction Amount], (OdometersTable[Vehicle Plate] = vPlate) * IFS(scope = "Day", OdometersTable[Date] = anchorD, scope = "Month", (YEAR(OdometersTable[Date]) = YEAR(anchorD)) * (MONTH(OdometersTable[Date]) = MONTH(anchorD)), scope = "Annual", YEAR(OdometersTable[Date]) = YEAR(anchorD), TRUE, TRUE), 0))) |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| V | 3 | V3 | Maintenance Cost | Normal | #,##0.00 | Formula | =LET(vPlate, [@[Vehicle Plate]], anchorD, $B$1, scope, $C$1, SUM(FILTER(MaintenancesTable[Repair Total Cost], (MaintenancesTable[Vehicle Plate] = vPlate) * IFS(scope = "Day", MaintenancesTable[Date in] = anchorD, scope = "Month", (YEAR(MaintenancesTable[Date in]) = YEAR(anchorD)) * (MONTH(MaintenancesTable[Date in]) = MONTH(anchorD)), scope = "Annual", YEAR(MaintenancesTable[Date in]) = YEAR(anchorD), TRUE, TRUE), 0))) |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| W | 3 | W3 | Insurance Cost | Normal | #,##0.00 | Formula | =LET(vPlate, [@[Vehicle Plate]], anchorD, $B$1, scope, $C$1, SUM(FILTER(InsurancesTable[Cost], (InsurancesTable[Vehicle Plate] = vPlate) * IFS(scope = "Day", InsurancesTable[Renewal Date] = anchorD, scope = "Month", (YEAR(InsurancesTable[Renewal Date]) = YEAR(anchorD)) * (MONTH(InsurancesTable[Renewal Date]) = MONTH(anchorD)), scope = "Annual", YEAR(InsurancesTable[Renewal Date]) = YEAR(anchorD), TRUE, TRUE), 0))) |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| X | 3 | X3 | Permit Cost | Normal | #,##0.00 | Formula | =LET(vPlate, [@[Vehicle Plate]], anchorD, $B$1, scope, $C$1, SUM(FILTER(PermitsTable[Cost], (PermitsTable[Vehicle Plate] = vPlate) * IFS(scope = "Day", PermitsTable[Renewal Date] = anchorD, scope = "Month", (YEAR(PermitsTable[Renewal Date]) = YEAR(anchorD)) * (MONTH(PermitsTable[Renewal Date]) = MONTH(anchorD)), scope = "Annual", YEAR(PermitsTable[Renewal Date]) = YEAR(anchorD), TRUE, TRUE), 0))) |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| Y | 3 | Y3 | Road Tax and Inspection Cost | Normal | #,##0.00 | Formula | =LET(vPlate, [@[Vehicle Plate]], anchorD, $B$1, scope, $C$1, rtCost, SUM(FILTER(RoadTaxesTable[Cost], (RoadTaxesTable[Vehicle Plate] = vPlate) * IFS(scope = "Day", RoadTaxesTable[Renewal Date] = anchorD, scope = "Month", (YEAR(RoadTaxesTable[Renewal Date]) = YEAR(anchorD)) * (MONTH(RoadTaxesTable[Renewal Date]) = MONTH(anchorD)), scope = "Annual", YEAR(RoadTaxesTable[Renewal Date]) = YEAR(anchorD), TRUE, TRUE), 0)), inspCost, SUM(FILTER(InspectionsTable[Cost], (InspectionsTable[Vehicle Plate] = vPlate) * IFS(scope = "Day", InspectionsTable[Schedule Date In] = anchorD, scope = "Month", (YEAR(InspectionsTable[Schedule Date In]) = YEAR(anchorD)) * (MONTH(InspectionsTable[Schedule Date In]) = MONTH(anchorD)), scope = "Annual", YEAR(InspectionsTable[Schedule Date In]) = YEAR(anchorD), TRUE, TRUE), 0)), rtCost + inspCost) |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |

### Conditional Formatting
| AppliesTo | RuleType | Operator | Formula1 | Formula2 | Severity | StopIfTrue |
| :--- | :--- | :--- | :--- | :--- | :---: | :---: |
| $G$3:$G$6 | Type_6 |  |   |   | None | FALSE |
| $H$3:$K$6 | Type_6 |  |   |   | None | FALSE |
| $S$3:$Y$6 | Type_6 |  |   |   | None | FALSE |
| $B$3:$B$6 | Expression |  | =AND($A3<>"", ISBLANK(B3)) |   | None | FALSE |
| $A$3:$P$6 | Expression |  | =$C3="Decommissioned" |   | None | FALSE |
| $A$3:$P$6 | Expression |  | =$C3="Expired" |   | None | FALSE |
| $A$3:$P$6 | Expression |  | =$C3="In_Shop" |   | None | FALSE |

---
## TABLE: Sheet=`Odometers` | Table=`OdometersTable` | HeaderRow=1 | ArchetypeRow=2 | StartCol=1 | Columns=16

### Column Rules
| Col | Row | Address | Header | Style | NumberFormat | Type | InCellFormula | SourceList | SyncFormula | FilterExpr | TransformExpr | IsMulti | IsWarning | NativeType | NativeFormula | NativeAlert | NativeIgnoreBlank |
| :---: | :---: | :---: | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :---: | :---: | :--- | :--- | :---: | :---: |
| A | 2 | A2 | Date | DropDown | m/d/yyyy | CustomDropdown |   | =TODAY() |   |   | @item | FALSE | TRUE |  |   |  | TRUE |
| B | 2 | B2 | Vehicle Plate | DropDown | @ | CustomDropdown |   | =AllVehiclesTable[Vehicle Plate] | @right XLOOKUP(@item, AllVehiclesTable[Vehicle Plate], AllVehiclesTable[Vehicle Type], "") | IF(@right<>"", XLOOKUP(@item, AllVehiclesTable[Vehicle Plate], AllVehiclesTable[Vehicle Type], "") = @right, TRUE) | @item | FALSE | FALSE |  |   |  | TRUE |
| C | 2 | C2 | Vehicle Type | DropDown | @ | CustomDropdown |   | =VehicleTypeList | @left XLOOKUP(@item, AllVehiclesTable[Vehicle Type], AllVehiclesTable[Vehicle Plate], "") | TRUE | @item | FALSE | FALSE |  |   |  | TRUE |
| D | 2 | D2 | Odometer Before | Normal | #,##0 | Formula | =IF([@[Vehicle Plate]]="", 0,   LET(     priorLogs, FILTER([Current Odometer], ([Vehicle Plate]=[@[Vehicle Plate]]) * (ROW([Current Odometer]) < ROW([@[Current Odometer]])), 0),     maxPrior, MAX(priorLogs),     IF(maxPrior > 0, maxPrior, IFERROR(XLOOKUP([@[Vehicle Plate]], AllVehiclesTable[Vehicle Plate], AllVehiclesTable[Current Odometer], 0), 0))   ) ) |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| E | 2 | E2 | Current Odometer | Unlocked | 0 | NativeValidation |   |   |   |   |   | FALSE | FALSE | CustomFormula | =AND(ISNUMBER(E2), E2>=D2) | Stop | TRUE |
| F | 2 | F2 | Odometer Difference | Normal | #,##0 | Formula | =IF(AND(ISNUMBER([@[Current Odometer]]), ISNUMBER([@[Odometer Before]])), MAX(0, [@[Current Odometer]] - [@[Odometer Before]]), 0) |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| G | 2 | G2 | Fuel Transaction ID | DropDown | 0 | CustomDropdown |   | =MAX(OdometersTable[Fuel Transaction ID])+1 |   |   | @item | FALSE | TRUE |  |   |  | TRUE |
| H | 2 | H2 | Card Number | DropDown | @ | CustomDropdown |   | =FuelCardsTable[Fuel Card Number] | @target[Card Owner] XLOOKUP("" & @item, FuelCardsTable[Fuel Card Number], FuelCardsTable[Card Owner Name], ""); @target[Staff ID] XLOOKUP("" & @item, FuelCardsTable[Fuel Card Number], FuelCardsTable[Card Owner ID], "") | TRUE | @item | FALSE | FALSE |  |   |  | TRUE |
| I | 2 | I2 | Card Owner | Normal | General | Formula | =IFERROR(XLOOKUP([@[Card Number]], FuelCardsTable[Fuel Card Number], FuelCardsTable[Card Owner Name], "—"), "—") |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| J | 2 | J2 | Staff ID | DropDown | 0 | CustomDropdown |   | =StaffIDList | @right XLOOKUP(@item, StaffIDList, StaffNameList, "") | TRUE | @item | FALSE | FALSE |  |   |  | TRUE |
| K | 2 | K2 | Staff Name | DropDown | @ | CustomDropdown |   | =StaffNameList | @left XLOOKUP(@item, StaffNameList, StaffIDList, "") | TRUE | @item | FALSE | FALSE |  |   |  | TRUE |
| L | 2 | L2 | Station Name | Unlocked | @ | StandardInput |   |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| M | 2 | M2 | Project | DropDown | @ | CustomDropdown |   | =ProjectIDList |   |   | @item | FALSE | TRUE |  |   |  | TRUE |
| N | 2 | N2 | Transaction Amount | Unlocked | 0.00 | NativeValidation |   |   |   |   |   | FALSE | FALSE | Decimal | 0 | Stop | TRUE |
| O | 2 | O2 | Time | Unlocked | @ | StandardInput |   |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| P | 2 | P2 | Reciept Number | Unlocked | @ | StandardInput |   |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |

### Conditional Formatting
| AppliesTo | RuleType | Operator | Formula1 | Formula2 | Severity | StopIfTrue |
| :--- | :--- | :--- | :--- | :--- | :---: | :---: |
| $N$2 | Type_12 |  |   |   | None | FALSE |
| $F$2 | Type_6 |  |   |   | None | FALSE |
| $B$2 | Expression |  | =AND(OR($E2<>"", $N2<>""), ISBLANK($B2)) |   | None | FALSE |
| $E$2 | Expression |  | =AND(ISNUMBER($D2), ISNUMBER($E2), $E2<$D2) |   | None | FALSE |

---
## TABLE: Sheet=`Odometers` | Table=`FuelCardsTable` | HeaderRow=1 | ArchetypeRow=2 | StartCol=18 | Columns=4

### Column Rules
| Col | Row | Address | Header | Style | NumberFormat | Type | InCellFormula | SourceList | SyncFormula | FilterExpr | TransformExpr | IsMulti | IsWarning | NativeType | NativeFormula | NativeAlert | NativeIgnoreBlank |
| :---: | :---: | :---: | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :---: | :---: | :--- | :--- | :---: | :---: |
| R | 2 | R2 | Fuel Card Number | Unlocked | @ | NativeValidation | 70838153012929800 |   |   |   |   | FALSE | FALSE | CustomFormula | =LEN(TRIM(R2))>0 | Stop | FALSE |
| S | 2 | S2 | Fuel Company Names | DropDown | @ | CustomDropdown | Petronas | =FuelCompanyNamesTable[Fuel Company Names] |   |   | @item | FALSE | FALSE |  |   |  | TRUE |
| T | 2 | T2 | Card Owner ID | DropDown | @ | CustomDropdown | 276 | =VSTACK(AllVehiclesTable[Vehicle Plate], StaffIDList) | @right IFERROR(XLOOKUP(@item, StaffIDList, StaffNameList), IFERROR(TRIM(XLOOKUP(@item, AllVehiclesTable[Vehicle Plate], AllVehiclesTable[Brand]) & " " & XLOOKUP(@item, AllVehiclesTable[Vehicle Plate], AllVehiclesTable[Model])), "")) | IF(@right<>"", IFERROR(XLOOKUP(@item, StaffIDList, StaffNameList), IFERROR(TRIM(XLOOKUP(@item, AllVehiclesTable[Vehicle Plate], AllVehiclesTable[Brand]) & " " & XLOOKUP(@item, AllVehiclesTable[Vehicle Plate], AllVehiclesTable[Model])), "")) = @right, TRUE) | @item | FALSE | FALSE |  |   |  | TRUE |
| U | 2 | U2 | Card Owner Name | DropDown | @ | CustomDropdown | ABDUL HAFIY AZRI BIN MOHD ANUAR | =VSTACK(AllVehiclesTable[Brand] & " " & AllVehiclesTable[Model], StaffNameList) | @left IFERROR(XLOOKUP(@item, StaffNameList, StaffIDList), IFERROR(XLOOKUP(@item, AllVehiclesTable[Brand] & " " & AllVehiclesTable[Model], AllVehiclesTable[Vehicle Plate]), "")) | IF(@left<>"", TEXT(IFERROR(XLOOKUP(@item, StaffNameList, StaffIDList), IFERROR(XLOOKUP(@item, AllVehiclesTable[Brand] & " " & AllVehiclesTable[Model], AllVehiclesTable[Vehicle Plate]), "")), "@") = TEXT(@left, "@"), TRUE) | @item | FALSE | FALSE |  |   |  | TRUE |

### Conditional Formatting
| AppliesTo | RuleType | Operator | Formula1 | Formula2 | Severity | StopIfTrue |
| :--- | :--- | :--- | :--- | :--- | :---: | :---: |
| $R$2:$R$500 | Expression |  | =AND($S2<>"", ISBLANK($R2)) |   | Error | FALSE |

---
## TABLE: Sheet=`Odometers` | Table=`FuelCompanyNamesTable` | HeaderRow=1 | ArchetypeRow=2 | StartCol=23 | Columns=1

### Column Rules
| Col | Row | Address | Header | Style | NumberFormat | Type | InCellFormula | SourceList | SyncFormula | FilterExpr | TransformExpr | IsMulti | IsWarning | NativeType | NativeFormula | NativeAlert | NativeIgnoreBlank |
| :---: | :---: | :---: | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :---: | :---: | :--- | :--- | :---: | :---: |
| W | 2 | W2 | Fuel Company Names | Unlocked | @ | NativeValidation | Petronas |   |   |   |   | FALSE | FALSE | CustomFormula | =LEN(TRIM(W2))>0 | Stop | FALSE |

---
## TABLE: Sheet=`Maintenances` | Table=`MaintenancesTable` | HeaderRow=1 | ArchetypeRow=2 | StartCol=1 | Columns=25

### Column Rules
| Col | Row | Address | Header | Style | NumberFormat | Type | InCellFormula | SourceList | SyncFormula | FilterExpr | TransformExpr | IsMulti | IsWarning | NativeType | NativeFormula | NativeAlert | NativeIgnoreBlank |
| :---: | :---: | :---: | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :---: | :---: | :--- | :--- | :---: | :---: |
| A | 2 | A2 | Vehicle Plate | DropDown | @ | CustomDropdown | BKK 9965 | =AllVehiclesTable[Vehicle Plate] | @right XLOOKUP(@item, AllVehiclesTable[Vehicle Plate], AllVehiclesTable[Vehicle Type], "") | IF(@right<>"", XLOOKUP(@item, AllVehiclesTable[Vehicle Plate], AllVehiclesTable[Vehicle Type], "") = @right, TRUE) | @item | FALSE | FALSE |  |   |  | TRUE |
| B | 2 | B2 | Vehicle Types | DropDown | @ | CustomDropdown | Lorry_3T | =VehicleTypeList | @left XLOOKUP(@item, AllVehiclesTable[Vehicle Type], AllVehiclesTable[Vehicle Plate], "") | TRUE | @item | FALSE | FALSE |  |   |  | TRUE |
| C | 2 | C2 | Maintenance ID | DropDown | 0 | CustomDropdown | 2 | =MAX(MaintenancesTable[Maintenance ID])+1 |   |   | @item | FALSE | TRUE |  |   |  | TRUE |
| D | 2 | D2 | Service Type | Unlocked | @ | NativeValidation | Routine_Service |   |   |   |   | FALSE | FALSE | List | =ServiceTypeList | Stop | FALSE |
| E | 2 | E2 | Date in | DropDown | m/d/yyyy | CustomDropdown | 46248 | =TODAY() |   |   | @item | FALSE | TRUE |  |   |  | TRUE |
| F | 2 | F2 | Estimated Return | DropDown | m/d/yyyy | CustomDropdown | 46250 | =IF(@left<>"", @left+3, TODAY()+3) |   |   | @item | FALSE | TRUE |  |   |  | TRUE |
| G | 2 | G2 | Date completed | Unlocked | m/d/yyyy | NativeValidation | 46249 |   |   |   |   | FALSE | FALSE | Date | =E2 | Stop | TRUE |
| H | 2 | H2 | Status | DropDown | @ | CustomDropdown | Done | =IF(@left <> "", "Done", IF([@[Date in]] > TODAY(), "Scheduled", MaintenanceStatusList)) |   | =IF(@item = "Done", @left <> "", IF(@left <> "", FALSE, TRUE)) | @item | FALSE | FALSE |  |   |  | TRUE |
| I | 2 | I2 | Odometer In | Unlocked | 0 | NativeValidation | 0 |   |   |   |   | FALSE | FALSE | WholeNumber | 0 | Stop | TRUE |
| J | 2 | J2 | Odometer Out | Unlocked | 0 | NativeValidation | 0 |   |   |   |   | FALSE | FALSE | WholeNumber | 0 | Stop | TRUE |
| K | 2 | K2 | Repair Total Cost | Normal | #,##0.00 | Formula | =SUM(MaintenancesTable[@[Tyre]:[Parts]]) |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| L | 2 | L2 | Workshop Name | DropDown | @ | CustomDropdown | Pusat Tayar & Servis Klang | =WorkshopsTable[Workshop Name] | @right XLOOKUP(@item, WorkshopsTable[Workshop Name], WorkshopsTable[Workshop Address], "") | TRUE | @item | FALSE | FALSE |  |   |  | TRUE |
| M | 2 | M2 | Workshop Address | Normal | @ | StandardInput | No 44, Jalan Kapar, Klang |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| N | 2 | N2 | Remarks | Unlocked | @ | StandardInput |   |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| O | 2 | O2 | Tyre | Unlocked | 0.00 | NativeValidation | 464654 |   |   |   |   | FALSE | FALSE | Decimal | 0 | Stop | TRUE |
| P | 2 | P2 | Brake | Unlocked | 0.00 | NativeValidation |   |   |   |   |   | FALSE | FALSE | Decimal | 0 | Stop | TRUE |
| Q | 2 | Q2 | Clutch | Unlocked | 0.00 | NativeValidation | 897987 |   |   |   |   | FALSE | FALSE | Decimal | 0 | Stop | TRUE |
| R | 2 | R2 | Service | Unlocked | 0.00 | NativeValidation | 98798 |   |   |   |   | FALSE | FALSE | Decimal | 0 | Stop | TRUE |
| S | 2 | S2 | Fuel Pump | Unlocked | 0.00 | NativeValidation | 8979854 |   |   |   |   | FALSE | FALSE | Decimal | 0 | Stop | TRUE |
| T | 2 | T2 | Aircond | Unlocked | 0.00 | NativeValidation | 465 |   |   |   |   | FALSE | FALSE | Decimal | 0 | Stop | TRUE |
| U | 2 | U2 | Wiring | Unlocked | 0.00 | NativeValidation | 46546 |   |   |   |   | FALSE | FALSE | Decimal | 0 | Stop | TRUE |
| V | 2 | V2 | Battery | Unlocked | 0.00 | NativeValidation | 56 |   |   |   |   | FALSE | FALSE | Decimal | 0 | Stop | TRUE |
| W | 2 | W2 | Towing | Unlocked | 0.00 | NativeValidation |   |   |   |   |   | FALSE | FALSE | Decimal | 0 | Stop | TRUE |
| X | 2 | X2 | Cushion | Unlocked | 0.00 | NativeValidation |   |   |   |   |   | FALSE | FALSE | Decimal | 0 | Stop | TRUE |
| Y | 2 | Y2 | Parts | Unlocked | 0.00 | StandardInput |   |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |

### Conditional Formatting
| AppliesTo | RuleType | Operator | Formula1 | Formula2 | Severity | StopIfTrue |
| :--- | :--- | :--- | :--- | :--- | :---: | :---: |
| $O$2:$Y$3 | Type_12 |  |   |   | None | FALSE |
| $K$2:$K$3 | Type_12 |  |   |   | None | FALSE |
| $A$2:$A$3 | Expression |  | =AND(OR($E2<>"", $L2<>""), ISBLANK($A2)) |   | None | FALSE |
| $A$2:$H$3 | Expression |  | =$H2="InProgress" |   | None | FALSE |

---
## SHEET_HEADER_CELLS: Sheet=`Workshops`

| CellAddress | Row | Col | Style | NumberFormat | Formula | SourceList / NativeFormula | NativeType | NativeAlert | NativeIgnoreBlank |
| :---: | :---: | :---: | :--- | :--- | :--- | :--- | :--- | :---: | :---: |
| $B$1 | 1 | 2 | DisplayDate | m/d/yyyy | =TODAY() |   |  |  | TRUE |
| $C$1 | 1 | 3 | DisplayRange | @ | All_Time | =DateFilterScopeList | List | Stop | TRUE |

---
## TABLE: Sheet=`Workshops` | Table=`WorkshopsTable` | HeaderRow=2 | ArchetypeRow=3 | StartCol=1 | Columns=9

### Column Rules
| Col | Row | Address | Header | Style | NumberFormat | Type | InCellFormula | SourceList | SyncFormula | FilterExpr | TransformExpr | IsMulti | IsWarning | NativeType | NativeFormula | NativeAlert | NativeIgnoreBlank |
| :---: | :---: | :---: | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :---: | :---: | :--- | :--- | :---: | :---: |
| A | 3 | A3 | Workshop ID | DropDown | 0 | CustomDropdown | 1 | =MAX(WorkshopsTable[Workshop ID])+1 |   |   | @item | FALSE | TRUE |  |   |  | TRUE |
| B | 3 | B3 | Workshop Name | Unlocked | @ | NativeValidation | Ah Keong Auto Repairs |   |   |   |   | FALSE | FALSE | CustomFormula | =LEN(TRIM(B3))>0 | Stop | FALSE |
| C | 3 | C3 | Workshop Address | Unlocked | @ | StandardInput | Lot 12, Jalan Utarid U5/11, Shah Alam |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| D | 3 | D3 | Workshop Owner | Unlocked | @ | StandardInput | Tan Ah Keong |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| E | 3 | E3 | Workshop Primary Contact | Unlocked | @ | StandardInput | 012-3456789 |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| F | 3 | F3 | Workshop Secondary Contact | Unlocked | @ | StandardInput | 03-78451122 |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| G | 3 | G3 | Number of Vehicle Fixed | Normal | #,##0 | Formula | =COUNTIF(MaintenancesTable[Workshop Name], [@[Workshop Name]]) |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| H | 3 | H3 | Fixed Vehicle Plates | Normal | General | Formula | =TEXTJOIN(", ", TRUE, UNIQUE(FILTER(MaintenancesTable[Vehicle Plate], MaintenancesTable[Workshop Name]=[@[Workshop Name]], ""))) |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| I | 3 | I3 | Total Cost | Normal | #,##0.00 | Formula | =SUMIF(MaintenancesTable[Workshop Name], [@[Workshop Name]], MaintenancesTable[Repair Total Cost]) |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |

### Conditional Formatting
| AppliesTo | RuleType | Operator | Formula1 | Formula2 | Severity | StopIfTrue |
| :--- | :--- | :--- | :--- | :--- | :---: | :---: |
| $G$3:$G$4 | Type_6 |  |   |   | None | FALSE |
| $I$3:$I$4 | Type_12 |  |   |   | None | FALSE |
| $B$3:$B$4 | Expression |  | =AND($A3<>"", ISBLANK($B3)) |   | None | FALSE |

---
## TABLE: Sheet=`Permits` | Table=`PermitsTable` | HeaderRow=1 | ArchetypeRow=2 | StartCol=1 | Columns=8

### Column Rules
| Col | Row | Address | Header | Style | NumberFormat | Type | InCellFormula | SourceList | SyncFormula | FilterExpr | TransformExpr | IsMulti | IsWarning | NativeType | NativeFormula | NativeAlert | NativeIgnoreBlank |
| :---: | :---: | :---: | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :---: | :---: | :--- | :--- | :---: | :---: |
| A | 2 | A2 | Permit ID | DropDown | 0 | CustomDropdown | 1 | =MAX(PermitsTable[Permit ID])+1 |   |   | @item | FALSE | TRUE |  |   |  | TRUE |
| B | 2 | B2 | Vehicle Plate | DropDown | @ | CustomDropdown | BKK 9965 | =AllVehiclesTable[Vehicle Plate] | @right XLOOKUP(@item, AllVehiclesTable[Vehicle Plate], AllVehiclesTable[Vehicle Type], "") | IF(@right<>"", XLOOKUP(@item, AllVehiclesTable[Vehicle Plate], AllVehiclesTable[Vehicle Type], "") = @right, TRUE) | @item | FALSE | FALSE |  |   |  | TRUE |
| C | 2 | C2 | Vehicle Types | DropDown | @ | CustomDropdown | Lorry_3T | =VehicleTypeList | @left XLOOKUP(@item, AllVehiclesTable[Vehicle Type], AllVehiclesTable[Vehicle Plate], "") | TRUE | @item | FALSE | FALSE |  |   |  | TRUE |
| D | 2 | D2 | Permit Types | DropDown | @ | CustomDropdown | APAD_Permit_C | =VehiclePermitTypeList |   |   | @item | FALSE | FALSE |  |   |  | TRUE |
| E | 2 | E2 | Renewal Date | DropDown | m/d/yyyy | CustomDropdown | 46143 | =TODAY() |   |   | @item | FALSE | TRUE |  |   |  | TRUE |
| F | 2 | F2 | Expiry Date | DropDown | m/d/yyyy | CustomDropdown | 46508 | =IF(@left<>"", @left+365, TODAY()+365) |   |   | @item | FALSE | TRUE |  |   |  | TRUE |
| G | 2 | G2 | Status | DropDown | @ | CustomDropdown | Active | =IF(@left <> "", IF(@left < TODAY(), "Expired", VehiclePermitStatusList), VehiclePermitStatusList) |   | =IF(@item = "Active", IF(@left <> "", @left >= TODAY(), TRUE), TRUE) | @item | FALSE | TRUE |  |   |  | TRUE |
| H | 2 | H2 | Cost | Unlocked | 0.00 | NativeValidation | 120 |   |   |   |   | FALSE | FALSE | Decimal | 0 | Stop | TRUE |

### Conditional Formatting
| AppliesTo | RuleType | Operator | Formula1 | Formula2 | Severity | StopIfTrue |
| :--- | :--- | :--- | :--- | :--- | :---: | :---: |
| $H$2:$H$4 | Type_12 |  |   |   | None | FALSE |
| $F$2:$F$4 | Type_6 |  |   |   | None | FALSE |
| $B$2:$B$4 | Expression |  | =AND(OR($E2<>"", $D2<>""), ISBLANK($B2)) |   | None | FALSE |
| $F$2:$F$4 | Expression |  | =AND($F2<>"", TODAY()>$F2) |   | None | FALSE |
| $A$2:$H$4 | Expression |  | =$G2="Expired" |   | None | FALSE |

---
## TABLE: Sheet=`Road Taxes` | Table=`RoadTaxesTable` | HeaderRow=1 | ArchetypeRow=2 | StartCol=1 | Columns=7

### Column Rules
| Col | Row | Address | Header | Style | NumberFormat | Type | InCellFormula | SourceList | SyncFormula | FilterExpr | TransformExpr | IsMulti | IsWarning | NativeType | NativeFormula | NativeAlert | NativeIgnoreBlank |
| :---: | :---: | :---: | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :---: | :---: | :--- | :--- | :---: | :---: |
| A | 2 | A2 | Road Tax ID | DropDown | 0 | CustomDropdown | 1 | =MAX(RoadTaxesTable[Road Tax ID])+1 |   |   | @item | FALSE | TRUE |  |   |  | TRUE |
| B | 2 | B2 | Vehicle Plate | DropDown | @ | CustomDropdown | BKK 9965 | =AllVehiclesTable[Vehicle Plate] | @right XLOOKUP(@item, AllVehiclesTable[Vehicle Plate], AllVehiclesTable[Vehicle Type], "") | IF(@right<>"", XLOOKUP(@item, AllVehiclesTable[Vehicle Plate], AllVehiclesTable[Vehicle Type], "") = @right, TRUE) | @item | FALSE | FALSE |  |   |  | TRUE |
| C | 2 | C2 | Vehicle Types | DropDown | @ | CustomDropdown | Lorry_3T | =VehicleTypeList | @left XLOOKUP(@item, AllVehiclesTable[Vehicle Type], AllVehiclesTable[Vehicle Plate], "") | TRUE | @item | FALSE | FALSE |  |   |  | TRUE |
| D | 2 | D2 | Renewal Date | DropDown | m/d/yyyy | CustomDropdown | 46174 | =TODAY() |   |   | @item | FALSE | TRUE |  |   |  | TRUE |
| E | 2 | E2 | Expiry Date | DropDown | m/d/yyyy | CustomDropdown | 46539 | =IF(@left<>"", @left+365, TODAY()+365) |   |   | @item | FALSE | TRUE |  |   |  | TRUE |
| F | 2 | F2 | Status | DropDown | @ | CustomDropdown | Active | =IF(@left <> "", IF(@left < TODAY(), "Expired", RoadTaxStatusList), RoadTaxStatusList) |   | =IF(@left <> "", IF(@left < TODAY(), "Expired", RoadTaxStatusList), RoadTaxStatusList) | @item | FALSE | TRUE |  |   |  | TRUE |
| G | 2 | G2 | Cost | Unlocked | 0.00 | NativeValidation | 850 |   |   |   |   | FALSE | FALSE | Decimal | 0 | Stop | TRUE |

### Conditional Formatting
| AppliesTo | RuleType | Operator | Formula1 | Formula2 | Severity | StopIfTrue |
| :--- | :--- | :--- | :--- | :--- | :---: | :---: |
| $G$2:$G$4 | Type_12 |  |   |   | None | FALSE |
| $E$2:$E$4 | Type_6 |  |   |   | None | FALSE |
| $B$2:$B$4 | Expression |  | =AND($D2<>"", ISBLANK($B2)) |   | None | FALSE |
| $E$2:$E$4 | Expression |  | =AND($E2<>"", TODAY()>$E2) |   | None | FALSE |
| $A$2:$G$4 | Expression |  | =$F2="Expired" |   | None | FALSE |

---
## TABLE: Sheet=`Inspections` | Table=`InspectionsTable` | HeaderRow=1 | ArchetypeRow=2 | StartCol=1 | Columns=12

### Column Rules
| Col | Row | Address | Header | Style | NumberFormat | Type | InCellFormula | SourceList | SyncFormula | FilterExpr | TransformExpr | IsMulti | IsWarning | NativeType | NativeFormula | NativeAlert | NativeIgnoreBlank |
| :---: | :---: | :---: | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :---: | :---: | :--- | :--- | :---: | :---: |
| A | 2 | A2 | Inspection ID | DropDown | 0 | CustomDropdown | 1 | =MAX(InspectionsTable[Inspection ID])+1 |   |   | @item | FALSE | TRUE |  |   |  | TRUE |
| B | 2 | B2 | Vehicle Plate | DropDown | @ | CustomDropdown | BKK 9965 | =AllVehiclesTable[Vehicle Plate] | @right XLOOKUP(@item, AllVehiclesTable[Vehicle Plate], AllVehiclesTable[Vehicle Type], "") | IF(@right<>"", XLOOKUP(@item, AllVehiclesTable[Vehicle Plate], AllVehiclesTable[Vehicle Type], "") = @right, TRUE) | @item | FALSE | FALSE |  |   |  | TRUE |
| C | 2 | C2 | Vehicle Types | DropDown | @ | CustomDropdown | Lorry_3T | =VehicleTypeList | @left XLOOKUP(@item, AllVehiclesTable[Vehicle Type], AllVehiclesTable[Vehicle Plate], "") | TRUE | @item | FALSE | FALSE |  |   |  | TRUE |
| D | 2 | D2 | Schedule Date In | DropDown | m/d/yyyy | CustomDropdown | 46275 | =TODAY() |   |   | @item | FALSE | TRUE |  |   |  | TRUE |
| E | 2 | E2 | Date Out | DropDown | m/d/yyyy | CustomDropdown | 46275 | =IF(@left<>"", @left, TODAY()) |   |   | @item | FALSE | TRUE |  |   |  | TRUE |
| F | 2 | F2 | Odometer In | Unlocked | 0 | NativeValidation | 142800 |   |   |   |   | FALSE | FALSE | WholeNumber | 0 | Stop | TRUE |
| G | 2 | G2 | Odometer Out | Unlocked | 0 | NativeValidation | 142815 |   |   |   |   | FALSE | FALSE | WholeNumber | 0 | Stop | TRUE |
| H | 2 | H2 | Expiry Date | DropDown | m/d/yyyy | CustomDropdown | 46456 | =EDATE(IF([@[Date Out]]<>"", [@[Date Out]], TODAY()), 6) |   |   | @item | FALSE | TRUE |  |   |  | TRUE |
| I | 2 | I2 | Status | DropDown | @ | CustomDropdown | Active | =IF(@item = "Active", IF(@left <> "", @left >= TODAY(), TRUE), TRUE) |   | =IF(@item = "Active", IF(@left <> "", @left >= TODAY(), TRUE), TRUE) | @item | FALSE | TRUE |  |   |  | TRUE |
| J | 2 | J2 | Cost | Unlocked | 0.00 | NativeValidation | 75 |   |   |   |   | FALSE | FALSE | Decimal | 0 | Stop | TRUE |
| K | 2 | K2 | Address | Unlocked | @ | StandardInput | Puspakom Padang Jawa |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| L | 2 | L2 | Remarks | Unlocked | @ | StandardInput |   |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |

### Conditional Formatting
| AppliesTo | RuleType | Operator | Formula1 | Formula2 | Severity | StopIfTrue |
| :--- | :--- | :--- | :--- | :--- | :---: | :---: |
| $J$2:$J$4 | Type_12 |  |   |   | None | FALSE |
| $H$2:$H$4 | Type_6 |  |   |   | None | FALSE |
| $B$2:$B$4 | Expression |  | =AND(OR($D2<>"", $E2<>""), ISBLANK($B2)) |   | None | FALSE |
| $H$2:$H$4 | Expression |  | =AND($H2<>"", TODAY()>$H2) |   | None | FALSE |
| $A$2:$L$4 | Expression |  | =$I2="Expired" |   | None | FALSE |

---
## TABLE: Sheet=`Insurances` | Table=`InsurancesTable` | HeaderRow=1 | ArchetypeRow=2 | StartCol=1 | Columns=13

### Column Rules
| Col | Row | Address | Header | Style | NumberFormat | Type | InCellFormula | SourceList | SyncFormula | FilterExpr | TransformExpr | IsMulti | IsWarning | NativeType | NativeFormula | NativeAlert | NativeIgnoreBlank |
| :---: | :---: | :---: | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :---: | :---: | :--- | :--- | :---: | :---: |
| A | 2 | A2 | Renewal Date | DropDown | m/d/yyyy | CustomDropdown | 46174 | =TODAY() |   |   | @item | FALSE | TRUE |  |   |  | TRUE |
| B | 2 | B2 | Vehicle Plate | DropDown | @ | CustomDropdown | BKK 9965 | =AllVehiclesTable[Vehicle Plate] | @target[Vehicle Types] XLOOKUP(@item, AllVehiclesTable[Vehicle Plate], AllVehiclesTable[Vehicle Type], "") | IF(@right<>"", XLOOKUP(@item, AllVehiclesTable[Vehicle Plate], AllVehiclesTable[Vehicle Type], "") = @right, TRUE) | @item | FALSE | FALSE |  |   |  | TRUE |
| C | 2 | C2 | Vehicle Types | DropDown | @ | CustomDropdown | Lorry_3T | =VehicleTypeList | @left XLOOKUP(@item, AllVehiclesTable[Vehicle Type], AllVehiclesTable[Vehicle Plate], "") | TRUE | @item | FALSE | FALSE |  |   |  | TRUE |
| D | 2 | D2 | Insurance Renewal ID | DropDown | 0 | CustomDropdown | 1 | =MAX(InsurancesTable[Insurace Renewal ID])+1 |   |   | @item | FALSE | TRUE |  |   |  | TRUE |
| E | 2 | E2 | Insurance Number | Unlocked | @ | NativeValidation | POL-ALZ-88912 |   |   |   |   | FALSE | FALSE | CustomFormula | =LEN(TRIM(E2))>0 | Stop | FALSE |
| F | 2 | F2 | Insurance Type | DropDown | @ | CustomDropdown | Comprehensive | =InsuranceTypeList |   |   | @item | FALSE | FALSE |  |   |  | TRUE |
| G | 2 | G2 | Status | DropDown | @ | CustomDropdown | Active | =IF([@[Expiry Date]] <> "", IF([@[Expiry Date]] < TODAY(), "Expired", InsuranceStatusList), InsuranceStatusList) |   | =IF(@item = "Active", IF([@[Expiry Date]] <> "", [@[Expiry Date]] >= TODAY(), TRUE), TRUE) | @item | FALSE | TRUE |  |   |  | TRUE |
| H | 2 | H2 | Insurance Company | DropDown | @ | CustomDropdown | Allianz General Insurance | =InsuranceCompaniesTable[Insurance Company] | @target[Company Contact Number] XLOOKUP(@item, InsuranceCompaniesTable[Insurance Company], InsuranceCompaniesTable[Company Contact Number], ""); @target[Insurance Agent] XLOOKUP(@item, InsuranceCompaniesTable[Insurance Company], InsuranceCompaniesTable[Insurance Agent], "") | IF([@[Insurance Agent]]<>"", XLOOKUP(@item, InsuranceCompaniesTable[Insurance Company], InsuranceCompaniesTable[Insurance Agent], "")=[@[Insurance Agent]], TRUE) | @item | FALSE | FALSE |  |   |  | TRUE |
| I | 2 | I2 | Company Contact Number | Unlocked | @ | StandardInput | 1300-22-5542 |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| J | 2 | J2 | Insurance Agent | DropDown | @ | CustomDropdown | Vincent Agency | =InsuranceCompaniesTable[Insurance Agent] | @target[Insurance Company] XLOOKUP(@item, InsuranceCompaniesTable[Insurance Agent], InsuranceCompaniesTable[Insurance Company], ""); @target[Company Contact Number] XLOOKUP(@item, InsuranceCompaniesTable[Insurance Agent], InsuranceCompaniesTable[Company Contact Number], "") | IF([@[Insurance Company]]<>"", XLOOKUP(@item, InsuranceCompaniesTable[Insurance Agent], InsuranceCompaniesTable[Insurance Company], "")=[@[Insurance Company]], TRUE) | @item | FALSE | FALSE |  |   |  | TRUE |
| K | 2 | K2 | Expiry Date | DropDown | m/d/yyyy | CustomDropdown | 46539 | =IF(@left<>"", @left+365, TODAY()+365) |   |   | @item | FALSE | TRUE |  |   |  | TRUE |
| L | 2 | L2 | Cost | Unlocked | 0.00 | NativeValidation | 2850 |   |   |   |   | FALSE | FALSE | Decimal | 0 | Stop | TRUE |
| M | 2 | M2 | Remarks | Unlocked | @ | StandardInput |   |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |

### Conditional Formatting
| AppliesTo | RuleType | Operator | Formula1 | Formula2 | Severity | StopIfTrue |
| :--- | :--- | :--- | :--- | :--- | :---: | :---: |
| $L$2:$L$3 | Type_12 |  |   |   | None | FALSE |
| $K$2:$K$3 | Type_6 |  |   |   | None | FALSE |
| $B$2:$B$3 | Expression |  | =AND(OR($E2<>"", $L2<>""), ISBLANK($B2)) |   | None | FALSE |
| $K$2:$K$3 | Expression |  | =AND($K2<>"", TODAY()>$K2) |   | None | FALSE |
| $A$2:$G$3 | Expression |  | =$G2="Expired" |   | None | FALSE |

---
## TABLE: Sheet=`Insurances` | Table=`InsuranceCompaniesTable` | HeaderRow=1 | ArchetypeRow=2 | StartCol=15 | Columns=7

### Column Rules
| Col | Row | Address | Header | Style | NumberFormat | Type | InCellFormula | SourceList | SyncFormula | FilterExpr | TransformExpr | IsMulti | IsWarning | NativeType | NativeFormula | NativeAlert | NativeIgnoreBlank |
| :---: | :---: | :---: | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :---: | :---: | :--- | :--- | :---: | :---: |
| O | 2 | O2 | Insurance Company | Unlocked | @ | NativeValidation | Allianz General Insurance |   |   |   |   | FALSE | FALSE | CustomFormula | =LEN(TRIM(O2))>0 | Stop | FALSE |
| P | 2 | P2 | Company Contact Number | Unlocked | @ | StandardInput | 1300-22-5542 |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| Q | 2 | Q2 | Company Address | Unlocked | @ | StandardInput | Menara Allianz Sentral, KL |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| R | 2 | R2 | Insurance Agent | Unlocked | @ | StandardInput | Vincent Agency |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| S | 2 | S2 | Agent Contact Number | Unlocked | @ | StandardInput | 012-2233445 |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| T | 2 | T2 | Number of Vehicles Under Agent | Unlocked | 0 | Formula | =COUNTIF(InsurancesTable[Insurance Agent], [@[Insurance Agent]]) |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |
| U | 2 | U2 | Vehicle Plates Under Agent | Unlocked | @ | Formula | =TEXTJOIN(", ", TRUE, UNIQUE(FILTER(InsurancesTable[Vehicle Plate], InsurancesTable[Insurance Agent]=[@[Insurance Agent]], ""))) |   |   |   |   | FALSE | FALSE |  |   |  | TRUE |

### Conditional Formatting
| AppliesTo | RuleType | Operator | Formula1 | Formula2 | Severity | StopIfTrue |
| :--- | :--- | :--- | :--- | :--- | :---: | :---: |
| $T$2:$T$3 | Type_6 |  |   |   | None | FALSE |
| $O$2:$O$500 | Expression |  | =AND($R2<>"", ISBLANK($O2)) |   | Error | FALSE |

