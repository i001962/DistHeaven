# DistHeaven

Excel Task Pane Add-in to create Monte Carlo Trials
Current implementation creates 2 correlated uniform distributions.

To use:
npm install
npm start

From excel 'sideload' the add-in using Insert->My Add-ins (dropdown menu to see developer add-ins)
Enter correlation coeff and number of trials into a row and select them. eg
C1 = 0.7
D1 = 100
Select the Add-in from the toolbar and hit the create dists button.
A new worksheet will be created with two correlated uniform distributions.
As a bonus a third uniform distribution is also created using the HDR1 Random Access Random Number Generator.
The nice thing about HDR1 is that it fits into one cell and works nicely with Excel's data table.


To read other learning sources and learn how debug add-ins see:
https://app.memphis.io/publish/Creating-Excel-Js-API-Taskpanel-Add-in-with-React/5cbe79f8af1274057ac0d3ad
