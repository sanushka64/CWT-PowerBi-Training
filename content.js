/* Creative Web Technology — step-by-step Power BI training content. */
const FORMULAS = [
{ id:"month-table", title:"Month Table (January–December)", tag:"Calculated table", description:"Creates a small disconnected table for the chart X-axis, matrix column headers and month slicer. Sort Month by MonthNo.", code:`Month Table =
DATATABLE(
    "Month", STRING,
    "MonthNo", INTEGER,
    {
        {"January", 1},
        {"February", 2},
        {"March", 3},
        {"April", 4},
        {"May", 5},
        {"June", 6},
        {"July", 7},
        {"August", 8},
        {"September", 9},
        {"October", 10},
        {"November", 11},
        {"December", 12}
    }
)`},
{ id:"actual-month",title:"Actual Revenue by Month",tag:"Measure", description:"Reads the selected month from Month Table and adds the matching Actual Revenue column. Designed for month columns kept in wide format.",code:`Actual Revenue by Month =
VAR SelectedMonth = SELECTEDVALUE('Month Table'[Month])
RETURN
IF(
    ISBLANK(SelectedMonth),
    BLANK(),
    SUMX(
        'Actual Revenue',
        VAR RawAmount =
            SWITCH(
                SelectedMonth,
                "January", 'Actual Revenue'[January] & "",
                "February", 'Actual Revenue'[February] & "",
                "March", 'Actual Revenue'[March] & "",
                "April", 'Actual Revenue'[April] & "",
                "May", 'Actual Revenue'[May] & "",
                "June", 'Actual Revenue'[June] & "",
                "July", 'Actual Revenue'[July] & "",
                "August", 'Actual Revenue'[August] & "",
                "September", 'Actual Revenue'[September] & "",
                "October", 'Actual Revenue'[October] & "",
                "November", 'Actual Revenue'[November] & "",
                "December", 'Actual Revenue'[December] & "",
                BLANK()
            )
        VAR CleanAmount =
            SUBSTITUTE(
                SUBSTITUTE(
                    SUBSTITUTE(
                        SUBSTITUTE(TRIM(RawAmount), ",", ""),
                        "$", ""
                    ),
                    "(", "-"
                ),
                ")", ""
            )
        RETURN IFERROR(VALUE(CleanAmount), 0)
    )
)`},
{ id:"estimated-month",title:"Estimated Revenue by Month",tag:"Measure",description:"Matches the selected month from the Estimated Revenue table. Check whether source estimates are stored in thousands before setting the scaling factor.",code:`Estimated Revenue by Month =
VAR SelectedMonth = SELECTEDVALUE('Month Table'[Month])
VAR ScaleFactor = 1  // Change to 1000 ONLY if estimates are in thousands
RETURN
IF(
    ISBLANK(SelectedMonth),
    BLANK(),
    SUMX(
        'Estimated Revenue',
        VAR RawAmount =
            SWITCH(
                SelectedMonth,
                "January", 'Estimated Revenue'[January] & "",
                "February", 'Estimated Revenue'[February] & "",
                "March", 'Estimated Revenue'[March] & "",
                "April", 'Estimated Revenue'[April] & "",
                "May", 'Estimated Revenue'[May] & "",
                "June", 'Estimated Revenue'[June] & "",
                "July", 'Estimated Revenue'[July] & "",
                "August", 'Estimated Revenue'[August] & "",
                "September", 'Estimated Revenue'[September] & "",
                "October", 'Estimated Revenue'[October] & "",
                "November", 'Estimated Revenue'[November] & "",
                "December", 'Estimated Revenue'[December] & "",
                BLANK()
            )
        VAR CleanAmount =
            SUBSTITUTE(
                SUBSTITUTE(
                    SUBSTITUTE(
                        SUBSTITUTE(TRIM(RawAmount), ",", ""),
                        "$", ""
                    ),
                    "(", "-"
                ),
                ")", ""
            )
        RETURN IFERROR(VALUE(CleanAmount), 0) * ScaleFactor
    )
)`},
{ id:"total-actual",title:"Total Actual Revenue",tag:"Measure",description:"Sums the month-level Actual Revenue measure across the months currently selected. Useful for the revenue-code contribution pie chart.",code:`Total Actual Revenue =
SUMX(
    VALUES('Month Table'[Month]),
    [Actual Revenue by Month]
)`},
{ id:"total-estimated",title:"Total Estimated Revenue",tag:"Measure",description:"Sums the selected monthly estimates. Does not inherit an Actual Revenue Year slicer unless a separate filter is applied.",code:`Total Estimated Revenue =
SUMX(
    VALUES('Month Table'[Month]),
    [Estimated Revenue by Month]
)`},
{ id:"estimated-year",title:"Estimated Revenue (Year Filtered)",tag:"Measure",description:"Optional bridge for a Year slicer based on Actual Revenue[Year]. No extra Year Table needed. Applies that year selection to Estimated Revenue.",code:`Estimated Revenue (Year Filtered) =
CALCULATE(
    [Estimated Revenue by Month],
    TREATAS(
        VALUES('Actual Revenue'[Year]),
        'Estimated Revenue'[Year]
    )
)`},
{ id:"revenue-variance",title:"Revenue Variance by Month",tag:"Measure",description:"Actual minus estimate for each revenue code, year and month. TREATAS connects filter context without new physical relationships.",code:`Revenue Variance by Month =
VAR ActualAmount = [Actual Revenue by Month]
VAR EstimatedAmount =
    CALCULATE(
        [Estimated Revenue by Month],
        TREATAS(
            VALUES('Actual Revenue'[Revenue Code]),
            'Estimated Revenue'[Revenue Code]
        ),
        TREATAS(
            VALUES('Actual Revenue'[Year]),
            'Estimated Revenue'[Year]
        )
    )
RETURN
    IF(
        ISBLANK(ActualAmount) || ISBLANK(EstimatedAmount),
        BLANK(),
        ActualAmount - EstimatedAmount
    )`},
{ id:"variance-color",title:"Variance Background Color (optional)",tag:"Measure",description:"Optional simpler field-value approach for red / green matrix formatting.",code:`Variance Background Color =
VAR VarianceAmount = [Revenue Variance by Month]
RETURN
    SWITCH(
        TRUE(),
        ISBLANK(VarianceAmount), "#FFFFFF",
        VarianceAmount < 0, "#F8D7DA",
        VarianceAmount >= 0, "#D1E7DD"
    )`}
];

const STEPS = [
{
 id:"excel", title:"Import the Excel workbook", category:"Getting started", duration:"3–5 min", subtitle:"Load the four revenue worksheets without changing the source data.",
 instructions:[
 "Open Power BI Desktop and choose a blank report.",
 "Go to Home → Get data → Excel workbook.",
 "Browse to Revenue Dashboard 1.xlsx, select the file, and click Open.",
 "In Navigator, tick Actual Revenue, Estimated Revenue, Daily Revenue, and Revenue Codes.",
 "Click Load. Wait until the four tables appear in the Data pane on the right.",
 "Expand Actual Revenue and Estimated Revenue and locate Year, Revenue Code, and January–December."
 ], note:"These instructions work with the original wide-format month columns. No Power Query transformation or unpivot is required.", formulas:[], fields:[], result:"Four tables appear in the Data pane."
},
{
 id:"page",title:"Set up the report and Month Table",category:"Report setup",duration:"4–6 min",subtitle:"Make room for the visuals and create a January–December axis.",
 instructions:[
 "Click + at the bottom of Power BI Desktop to create a new report page.",
 "Rename the page to Revenue Analysis Demo.",
 "For more vertical space, click the empty canvas → Format page → Canvas settings → Type: Custom; set Width to 1280 and Height to 1200.",
 "Go to Modeling → New table and paste the Month Table DAX below.",
 "In the Data pane, select Month Table[Month] → Column tools → Sort by column → MonthNo.",
 "Leave Month Table without a relationship to the revenue tables."
 ],note:"Month Table is a disconnected helper table; it is the only additional table used in this basic workflow.",formulas:["month-table"],fields:[],result:"A correctly sorted Month field is available for visuals and month filtering."
},
{
 id:"actual",title:"Calculate Actual Revenue by Month",category:"DAX measures",duration:"4–6 min",subtitle:"Read each month's value from the existing Actual Revenue columns.",
 instructions:[
 "Select Modeling → New measure.",
 "Copy the Actual Revenue by Month DAX below into the formula bar.",
 "Press Enter and confirm Power BI creates a measure with a calculator icon.",
 "The measure reads the month selected in Month Table and sums the matching Actual Revenue column.",
 "Verify one value against the source worksheet; check that totals or heading rows are not counted twice."
 ],note:"The measure uses VALUE to convert numeric-looking text and IFERROR to handle conversion failures as zero. Invalid text may hide data-quality issues, so validate a sample.",formulas:["actual-month"],fields:[],result:"Actual Revenue by Month is ready to use in the clustered chart and heatmap."
},
{
 id:"estimated",title:"Calculate Estimated Revenue by Month",category:"DAX measures",duration:"4–6 min",subtitle:"Add the estimated series using the same month selection.",
 instructions:[
 "Go to Modeling → New measure.",
 "Copy the Estimated Revenue by Month DAX below and press Enter.",
 "Check the units in the Excel workbook. If estimates are stored in thousands but actual amounts are in full units, update ScaleFactor from 1 to 1000.",
 "Compare Actual and Estimated for one month before making a chart."
 ],note:"Do not multiply by 1000 simply because a chart looks small. Confirm the two source tables use different units first.",formulas:["estimated-month"],fields:[],result:"Actual and Estimated monthly measures are both available."
},
{
 id:"column",title:"Build the clustered column chart",category:"Charts",duration:"5–7 min",subtitle:"Show Actual and Estimated side by side for each month.",
 instructions:[
 "Go to Report view and click an empty part of the canvas.",
 "Select Clustered column chart in Visualizations.",
 "Place the fields below in the appropriate wells.",
 "Leave Legend empty. The two measures automatically define the two series.",
 "Click the chart's three-dot menu → Sort axis → Month → Ascending.",
 "Change the title to Monthly Actual vs Estimated Revenue.",
 "Turn on the legend and data labels. Set Actual to blue and Estimated to yellow."
 ],fields:[["X-axis","Month Table[Month]"],["Y-axis","Actual Revenue by Month"],["Y-axis","Estimated Revenue by Month"],["Legend","Leave empty"]],note:"The chart combines available years until year filtering is added. Month should display January through December, not alphabetical order.",formulas:[],result:"Two columns per month compare actual and estimated collections."
},
{
 id:"pie",title:"Build the revenue-code pie chart",category:"Charts",duration:"5–7 min",subtitle:"Show the share of ACTUAL collected revenue contributed by each revenue code.",
 instructions:[
 "Go to Modeling → New measure and create Total Actual Revenue with the DAX below.",
 "Create Total Estimated Revenue as an optional reference measure if it was part of your earlier setup; it is NOT used as a pie slice.",
 "Click an empty part of the canvas and choose Pie chart.",
 "Place Actual Revenue[Revenue Code] in Legend and Total Actual Revenue in Values.",
 "Format visual → Detail labels / Data labels → On → Percent of total. Show category names if readable.",
 "Set the title to Revenue Contribution by Revenue Code (%).",
 "For many tiny slices, use a Top N visual filter. The full pie represents 100% of Actual Revenue in the current filter context."
 ],fields:[["Legend","Actual Revenue[Revenue Code]"],["Values","Total Actual Revenue"],["Tooltips (optional)","Actual Revenue[Description]"]],note:"This pie is NOT comparing collected against estimated totals and is NOT a target-achievement chart.",formulas:["total-actual","total-estimated"],result:"Each code shows its percentage contribution to Actual Revenue."
},
{
 id:"heatmap",title:"Create the revenue heatmap",category:"Matrix visuals",duration:"5–7 min",subtitle:"Use a matrix to compare Actual Revenue across codes and months.",
 instructions:[
 "Click the canvas and select Matrix from Visualizations.",
 "Drag the fields into Rows, Columns and Values as shown below.",
 "With the matrix selected, open Format visual → Cell elements.",
 "Enable Background color → fx → Format style: Gradient.",
 "Based on field: Actual Revenue by Month. Use minimum #FFF0BB (light yellow) and maximum #DB8D18 (dark orange).",
 "Apply to Values only. Turn off unnecessary grand totals.",
 "Set the title to Revenue Heatmap – Revenue Code by Month."
 ],fields:[["Rows","Actual Revenue[Revenue Code]"],["Columns","Month Table[Month]"],["Values","Actual Revenue by Month"]],note:"Light cells indicate smaller values; darker orange cells indicate larger values. The matrix uses Actual Revenue only.",formulas:[],result:"A yellow-to-orange heatmap comparing revenue codes and months."
},
{
 id:"format",title:"Format your visuals",category:"Appearance",duration:"4–6 min",subtitle:"Improve readability with titles, borders, labels and tooltips.",
 instructions:[
 "Select the clustered column chart → Format visual → General → Effects → turn Background on and choose a light colour.",
 "In General → Effects, turn Visual border on and choose a subtle grey border.",
 "In General → Title, use a descriptive title and a readable font size (for example 16).",
 "In Visual → Data labels, turn labels on; adjust text size and display units to suit the source amounts.",
 "Hover over a column to show the default tooltip. Optionally add a compatible measure in Build visual → Tooltips.",
 "Keep Actual and Estimated colours consistent throughout the report.",
 "Arrange the clustered chart and pie chart on top, then the heatmap below."
 ],note:"Avoid many competing colours and excessive labels. Formatting should make the values clearer.",formulas:[],fields:[],result:"The dashboard is organized and readable."
},
{
 id:"slicers",title:"Add Year and Month slicers",category:"Interactivity",duration:"6–8 min",subtitle:"Filter the completed report without creating a Year Table.",
 instructions:[
 "Insert a Slicer → drag Month Table[Month] into Field → Format visual → Slicer settings → Dropdown.",
 "Select September to check how the month selection affects the monthly chart, pie and heatmap.",
 "Insert a second Slicer → drag Actual Revenue[Year] into Field → choose Dropdown.",
 "Actual Revenue visuals respond to that Year slicer. Estimated Revenue is in a separate table, so create the Year Filtered measure below to make its series respond too.",
 "Replace Estimated Revenue by Month with Estimated Revenue (Year Filtered) in the clustered chart's Y-axis.",
 "Click a month bar. If the pie shows highlighted parts rather than only the selected month, select the column chart → Format ribbon → Edit interactions → choose Filter on the pie and heatmap.",
 "Explain visual-level, page-level and report-level filters in the Filters pane."
 ],note:"A slicer built from Actual Revenue[Year] does not automatically filter all measures from Estimated Revenue. If both source tables use Year, the TREATAS measure transfers the selected year. Column types must match.",formulas:["estimated-year"],fields:[["Month slicer Field","Month Table[Month]"],["Year slicer Field","Actual Revenue[Year]"]],result:"Month and Year selections control the intended visuals."
},
{
 id:"variance",title:"Show variance with conditional formatting",category:"Conditional formatting",duration:"6–8 min",subtitle:"Colour positive Actual-minus-Estimated differences green and negative differences red.",
 instructions:[
 "Go to Modeling → New measure and create Revenue Variance by Month with the DAX below.",
 "Add another Matrix; keep your original orange heatmap unchanged.",
 "Use Revenue Code in Rows, Month in Columns and Revenue Variance by Month in Values.",
 "Select the variance matrix → Format visual → Cell elements → Background color → fx.",
 "Format style: Rules. Based on field: Revenue Variance by Month. Apply to: Values only.",
 "Rule 1: 0 Percent (minimum) up to less than 0 Number → light red #F8D7DA.",
 "Click + New rule. Rule 2: greater than or equal to 0 Number up to 100 Percent (maximum) → light green #D1E7DD.",
 "Alternatively copy the optional Variance Background Color measure and choose Format style: Field value.",
 "Select another year/month and verify the cell values and colours update."
 ],fields:[["Rows","Actual Revenue[Revenue Code]"],["Columns","Month Table[Month]"],["Values","Revenue Variance by Month"]],note:"A positive value means Actual exceeds Estimated. Check scaling and Revenue Code/Year data types before interpreting the variance.",formulas:["revenue-variance","variance-color"],result:"A second matrix with green positive variances and red negative variances."
},
{
 id:"review",title:"Review and save the dashboard",category:"Finishing up",duration:"2–4 min",subtitle:"Check the numbers and interactions before using the report.",
 instructions:[
 "Check that the chart shows January–December in the correct order.",
 "Confirm Actual and Estimated source units are aligned, and verify at least one month's values against Excel.",
 "Confirm the pie represents revenue-code shares of Actual Revenue (not achievement against Estimated).",
 "Check that yellow/orange heatmap colours and green/red variance rules reflect the values.",
 "Select different years and months, and check that all intended visuals respond.",
 "Use Format page → Canvas settings → Custom page height if more space is needed.",
 "Save the Power BI report as Revenue Analysis Demo.pbix."
 ],formulas:[],fields:[],note:"Keep your workbook and .pbix together for future practice. Use only data that you have permission to share.",result:"Your completed practice dashboard is saved and ready for review."
}
];
window.CWT_DATA = { FORMULAS, STEPS };
