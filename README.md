# DistHeaven

Excel Task Pane Add-in to create Monte Carlo Trials
Current implementation creates 2 correlated uniform distributions.

## To use:
- npm install
- npm start
- Visit https://localhost:3000/ to confirm no certificate errors
- From excel 'sideload' the add-in using Insert->My Add-ins (dropdown menu to see developer add-ins)
- Enter correlation coeff and number of trials into a row and select them. eg
C1 = 0.7
D1 = 100
- Select the Add-in from the toolbar and hit the create dists button.
- A new worksheet will be created with two correlated uniform distributions.
- As a bonus a third uniform distribution is also created using the HDR1 Random Access Random Number Generator.
- The nice thing about HDR1 is that it fits into one cell and works nicely with Excel's data table.


### Based on
https://docs.microsoft.com/en-us/office/dev/add-ins/quickstarts/excel-quickstart-react

## SIPMath Compatible
To use these distributions and their trails in SIPMath models, you will need to manually set named ranges in Excel. Add PM_Trials named range in cell A1. Then also create ranges for each distribution. Those named ranges should each contain all of the trials for a given distribution. Start in cell C4 and go down then move to D4 and create another named range. Save that workbook and then use it as and Input Library in SIPMath Tools from https://ProbabilityManagement.org The named ranges are necessary during the input process in SIPMath Tools. You may use these distributions without SIPMath Tools by simply creating a data table. Here are some useful learning resources https://app.memphis.io/publish/SIP-Math/5cc1edd2af1274057ac0d4ca

## TODOs are found in app.tsx

## Publish, Deploy and learn more here
https://app.popdoc.io/kevin/creating-excel-js-api-taskpanel-add-in-with-react/5cbe79f8af1274057ac0d3ad
