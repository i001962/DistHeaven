/* global Office, Excel */

import * as React from 'react';
import { Header } from './Header';
import { Content } from './Content';
import Progress from './Progress';

import * as OfficeHelpers from '@microsoft/office-js-helpers';
// TODO Replace jStats with a seeded RNG npm package
var jStat = require('jStat').jStat;

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
  }

  deleteWorksheet = async () => {
    try {
      await Excel.run(async context => {
        var sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();

        if (sheets.items.length === 1) {
          console.log("Unable to delete the only worksheet in the workbook");
        } else {
          var lastSheet = sheets.items[sheets.items.length - 1];
          // TODO Be smart and only delete worksheet with specific name
          console.log(`Deleting worksheet named "${lastSheet.name}"`);
          lastSheet.delete();
        };
        await context.sync();

      });
    } catch (error) {
      OfficeHelpers.UI.notify(error);
      OfficeHelpers.Utilities.log(error);
    }
  }

  createCorrelUnifs = async () => {
    try {
      await Excel.run(async context => {
        // console.log("this is it");
        var newsheets = context.workbook.worksheets;
        var newsheet = newsheets.add("PM_Table");
        newsheet.load("name, position");
        await context.sync();

        //  console.log(`Added worksheet named "${newsheet.name}" in position ${newsheet.position}`);
        const selectedRange = context.workbook.getSelectedRange();
        selectedRange.load("values");
        await context.sync();

        console.log(JSON.stringify(selectedRange.values, null, 4));
        console.log('you are here', selectedRange.values[0]);
        //TAKE ONLY FIRST VALUE for Correlation and SECOND for Trial Count OR ADD SOME ERROR CHECKING LIKE A GOOD DEV WOULD
        var correlInput = selectedRange.values[0][0],
          trialsInput = selectedRange.values[0][1], // number of trials
          columns = 2,  // number of Distributions being created TODO make this a user input
          // TODO Allow user to select type of distributions to gerneate eg normal
          // mu = [0, 0], // for normal distributions
          // sigma = [0.25, 0.5], // for normal
          correlation = [[1.0, correlInput], [correlInput, 1.0]]; // TODO Expand to more than 2 dists
          // var data = await generateCorrLognorm(trialsInput, mu, sigma, correlation); // Corr Lognormals

        // Cheleski for corr uniforms et al
        var copula = generateCopula(columns, trialsInput, correlation);
        var // normal = jStat.normal(0, 1),
          // normal2 = jStat.normal(0, 1),
          // outputdists = [normal, normal2];
          //lognormal = jStat.lognormal(0, 0.5),
          uniform = jStat.uniform(0, 1),      // Three lines for corr uniform
          uniform2 = jStat.uniform(0, 1),     // Three lines for corr uniform
          outputdists = [uniform, uniform2];  // Three lines for corr uniform
          //console.log(outputdists);
        var samples = copula.map(function(x, row) { return outputdists[row].inv(x); });
        //console.log(samples);
        // console.log(data[0]["0"]);
        //  console.log(selectedRange.values);

        // Queue a command to add a new table to contain the results
        // Using ProbabilityManagement.org SIPMath standard for sheetname and format of distribution
        // trials but can not name ranges programmatically
        var sheettest = context.workbook.worksheets.getItem("PM_Table");
        var cell = sheettest.getCell(0, 0);
        cell.load("address, values");
        await context.sync();

        // console.log(`The headers start in cell "${cell.address}"`);
        for (let i = 0; i < columns; i++) {
          for (let j = 0; j < trialsInput; j++) {
            let cell = sheettest.getCell(j, i);
            cell.values = samples[i][j];

          }
        }

        // TODO Get smarter here and set range based on user inputs not artificial max of 10k
        var range = sheettest.getRange("A1:A10000");
        range.insert(Excel.InsertShiftDirection.right);
        await context.sync();

        for (let j = 0; j < trialsInput; j++) {
          let cell = sheettest.getCell(j, 0);
          cell.values = [[j + 1]];

        }
        var range = sheettest.getRange("A1:A10000");
        range.insert(Excel.InsertShiftDirection.right);
        await context.sync();

        range = sheettest.getRange("A1:E1");
        range.insert(Excel.InsertShiftDirection.down);
        await context.sync();

        range = sheettest.getRange("A1:E1");
        range.values = [[trialsInput, "Seeds", "Unknown", "Unknown", 6818051]]
        await context.sync();

        range = sheettest.getRange("B2:E2");
        range.insert(Excel.InsertShiftDirection.down);
        await context.sync();

        range = sheettest.getRange("B2:E2");
        range.values = [["Index", "CorrelUniform1", "CorrelUniform2", "Uniform-HDR1"]]
        await context.sync();

        range = sheettest.getRange("B3:D3");
        range.insert(Excel.InsertShiftDirection.down);
        await context.sync();

        range = sheettest.getRange("B3:D3");
        // TODO Ideally the PM_Table sheet could create ranges and be used in generation model of SIPMath Tools.
        // For now we use the sheet as an Input Library for SIPMath Tools instead.
        // range.values = [["Values", "=VLOOKUP(A1,B4:C13,2,)", "=VLOOKUP($A$1,B4:D13,3,)"]]
        // await context.sync();

        // Bonus adding the new HDR1 Random access random number generator for kicks
        range = sheettest.getRange("E3");
        range.formulas = [["=IF( 10 = 0, 50, NORMINV( ( MOD((( MOD( (E1+1000000)^2 + (E1+1000000)*($A$1+10000000), 99999989 )) + 1000007 ) * (( MOD( ($A$1+10000000)^2 + ($A$1+10000000) * ( MOD( (E1+1000000)^2 + (E1+1000000)*($A$1+10000000), 99999989 )), 99999989 )) + 1000013 ), 2147483647 ) + 0.5 ) / 2147483647, 50, 10 ))"]];
        await context.sync();

      });
    } catch (error) {
      OfficeHelpers.UI.notify(error);
      OfficeHelpers.Utilities.log(error);
    }
  }

  render() {
    const {
      title,
      isOfficeInitialized,
    } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
                    title= { title }
      logo = 'assets/logo-filled.png'
      message = 'Please sideload your add-in to see app body.'
        />
            );
    }

    return (
      <div className= 'ms-welcome' >
      <Header title='Generate Distributions' />
        <Content  buttonLabel = 'Create Uniform Dists' click = { this.createCorrelUnifs } />
          <Content  buttonLabel = 'Delete last sheet' click = { this.deleteWorksheet } />

            </div>
        );
  }
}
function generateCopula(rows, columns, correlation) {
  //https://en.wikipedia.org/wiki/Copula_(probability_theory)
  //Create uncorrelated standard normal samples
  var normSamples = jStat.randn(rows, columns);
  //Create lower triangular cholesky decomposition of correlation matrix
  var A = jStat(jStat.cholesky(correlation));
  //Create correlated samples through matrix multiplication
  var normCorrSamples = A.multiply(normSamples);
  //Convert to uniform correlated samples over 0,1 using normal CDF
  var normDist = jStat.normal(0, 1);
  var uniformCorrSamples = normCorrSamples.map(function(x) { return normDist.cdf(x); });
  return uniformCorrSamples;
}
// TODO add ux for Correlated Lognormals
// async function generateCorrLognorm(number, mu, sigma, correlation) {
//
//   //Create uniform correlated copula
//   var copula = await generateCopula(mu.length, number, correlation);
//
//   //Create unique lognormal distribution for each marginal
//   var lognormDists = [];
//   for (var i = 0; i < mu.length; i++) {
//     lognormDists.push(jStat.lognormal(mu[i], sigma[i]));
//     // console.log(lognormDists);
//   }
//
//   //Generate correlated lognormal samples using the inverse transform method:
//   //https://en.wikipedia.org/wiki/Inverse_transform_sampling
//   var lognormCorrSamples = await copula.map(function(_x, _row, _col) { return lognormDists[_row].inv(_x); });
//   return lognormCorrSamples;
// }
