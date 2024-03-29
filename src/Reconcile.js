'use strict';
/*
 * Copyright (c) 2013-2023 b3devs@gmail.com
 * MIT License: https://spdx.org/licenses/MIT.html
 */

import {Const} from './Constants.js';
import {Utils, toast} from './Utils.js';
import {SpreadsheetUtils} from './SpreadsheetUtils.js';
import {Debug} from './Debug.js';
import {Mint} from './MintApi.js';
import {Sheets} from './Sheets.js';


export const Reconcile = {

  RECON_ROW_COLOR: '#c5e1ff',
  RECON_COL_COLOR: '#ffdddd',
  RECON_RECORD_AMOUNT: -0.01,

  isReconcileAlreadyInProgress() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Const.SHEET_NAME_RECONCILE);
    const reconAccount = sheet.getRange(Const.RECON_ROW_TITLE, Const.RECON_COL_ACCOUNT, 1, 1);
//    Debug.log('Reconcile Account: ' + reconAccount.getValue());
    return (!!reconAccount.getValue());
  },

  getReconcileParams() {
    // Get reconcile params from cell
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Const.SHEET_NAME_RECONCILE);
    const paramsCell = sheet.getRange(Const.RECON_ROW_TARGET, Const.RECON_COL_SAVED_PARAMS, 1, 1);
    const paramsJson = paramsCell.getValue();
    if (paramsJson == null || paramsJson === '') {
      throw new Error('Reconcile parameters not found.');
    }
    return JSON.parse(paramsJson);
  },

  cancelReconcile(showPrompt) {
    if (showPrompt) {
      if (!this.isReconcileAlreadyInProgress()) {
        toast('No reconcile is in progress.', 'Reconcile');
        return;
      }

      // Cancel reconciling?
      if ('yes' !== Browser.msgBox('Reconcile', 'Cancel reconcile?', Browser.Buttons.YES_NO)) {
        return;
      }
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Const.SHEET_NAME_RECONCILE);
    sheet.activate();
    this.initReconcileSheet(sheet, null, 0, 0);
  },

  startReconcile() {
    try {
      if (!Mint.getClearedTag() || !Mint.getReconciledTag()) {
        Browser.msgBox('Clear / Reconcile not enabled', 'You cannot reconcile an account until you enable this feature by specifying the corresponding tags on the Settings sheet. Refer to the Help sheet for instructions on how to do this.', Browser.Buttons.OK);
        return;
      }

      if (this.isReconcileAlreadyInProgress()) {
        // The account name cell is already filled in. Is another account already 
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Const.SHEET_NAME_RECONCILE);
        sheet.activate();
        
        if ('no' === Browser.msgBox('Reconcile', 'Another account is currently being reconciled. Do you want to cancel it and reconcile a different account instead?', Browser.Buttons.YES_NO)) {
          return;
        }
      }

      toast('Getting reconcile history', 'Reconcile', 4);
      Reconcile.Window.show();
    }
    catch (e)
    {
        Debug.log(Debug.getExceptionInfo(e));
        toast('Exception: ' + Debug.getExceptionInfo(e), 'Reconcile', 15);
    }
  },

  /**
   * Called from reconcile_start.html
   * @param args {{account, accountType, mintAccount, endDate, prevBalance, newBalance}}
   */
  continueReconcile(args) {
    const {account, accountType, mintAccount, endDate, prevBalance, newBalance} = args;
    const {
      IDX_TXN_DATE,
      IDX_TXN_ACCOUNT,
      IDX_TXN_MINT_ACCOUNT,
      IDX_TXN_MERCHANT,
      IDX_TXN_AMOUNT,
      IDX_TXN_CLEAR_RECON,
      IDX_TXN_STATE,
      TXN_STATUS_PENDING,
      IDX_TXN_ID,
      IDX_TXN_PARENT_ID,
      IDX_RECON_DATE,
      IDX_RECON_TXN_ID,
      IDX_RECON_AMOUNT,
      IDX_RECON_RECONCILE,
      RECON_ROW_SUM,
      RECON_ROW_TARGET,
      RECON_COL_SAVED_PARAMS,
      SHEET_NAME_RECONCILE
    } = Const;

    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Const.SHEET_NAME_RECONCILE);

      this.initReconcileSheet(sheet, account, prevBalance, newBalance, endDate);

      // Get txns from TxnData sheet
      const txnRange = Utils.getTxnDataRange();
      let txnValues = txnRange.getValues();
      const txnDataLen = txnValues.length;
      let reconValues = [];
      let splitValues = {};
      
      // Loop through txns and copy the ones for the specified account that have not been
      // reconciled yet.
      for (let i = txnDataLen - 1; i >= 0; --i) {
        if (String(txnValues[i][IDX_TXN_ACCOUNT]) === account &&
            (!mintAccount || txnValues[i][IDX_TXN_MINT_ACCOUNT] === mintAccount) &&
            txnValues[i][IDX_TXN_CLEAR_RECON].toUpperCase() !== 'R' && // Ignore previously reconciled txns
            txnValues[i][IDX_TXN_STATE] !== TXN_STATUS_PENDING)  // Ignore pending txns
        {
          const parentId = txnValues[i][IDX_TXN_PARENT_ID];

          let reconRow = [
            txnValues[i][IDX_TXN_DATE],
            txnValues[i][IDX_TXN_MERCHANT],
            txnValues[i][IDX_TXN_AMOUNT],
            '',
            txnValues[i][IDX_TXN_CLEAR_RECON],
            (parentId ? 'S' : null),
            txnValues[i][IDX_TXN_ID]
          ];

          // Aggregate split txns into a single txn
          if (parentId) {
            // It's a split txn, sum up the separate split txns in a separate map
            const splitTxn = splitValues[`${parentId}`];
            if (!splitTxn) {
              reconRow[IDX_RECON_TXN_ID] = parentId;
              splitValues[parentId] = reconRow;
            } else {
              splitTxn[IDX_RECON_AMOUNT] += reconRow[IDX_RECON_AMOUNT];
            }
          } else {
            // Not a split txn, just add it to the reconValues array
            reconValues.push(reconRow);
          }
        }
      }
      for (let id in splitValues) {
        reconValues.push(splitValues[id]);
      }

      if (reconValues.length === 0) {
        toast(`There are no transactions to reconcile for account "${account}"`, 'Reconcile');
        return;
      }

      // Sort transactions by date, descending
      reconValues.sort(function(a, b) { return (a[IDX_RECON_DATE] > b[IDX_RECON_DATE] ? -1 : (a[IDX_RECON_DATE] < b[IDX_RECON_DATE] ? 1 : 0)); });
      
      let reconRange = Utils.getDataRange(SHEET_NAME_RECONCILE, IDX_RECON_TXN_ID + 1);
      reconRange = reconRange.offset(0, 0, reconValues.length, reconValues[0].length);
      reconRange.setValues(reconValues);

      const rowCount = reconRange.getNumRows();
      const amountColRange = reconRange.offset(0, IDX_RECON_AMOUNT, rowCount, 1);
      let rColRange = reconRange.offset(0, IDX_RECON_RECONCILE, rowCount, 1);

      const reconTotalCell = sheet.getRange(RECON_ROW_SUM, IDX_RECON_AMOUNT + 1, 1, 1);
      reconTotalCell.setFormula(`=SUMIF(${rColRange.getA1Notation()}, "R", ${amountColRange.getA1Notation()})`);
      reconTotalCell.setFontWeight('bold');
      reconTotalCell.setBackground(this.RECON_ROW_COLOR);
//      reconTotalCell.setBorder(true, true, true, true, false, false);

      // Set number format of Amount column to '$0.00'
      let numberFormats = amountColRange.getNumberFormats();
      for (let i = 0; i < rowCount; ++i) {
        numberFormats[i][0] = '$0.00';
      }
      amountColRange.setNumberFormats(numberFormats);

      // Set R column to bold and light red
      rColRange = rColRange.offset(-1, 0, rColRange.getNumRows() + 1, 1);
      let fontWeights = rColRange.getFontWeights();
      for (let i = 0; i < rowCount; ++i) {
        fontWeights[i][0] = 'bold';
      }
      rColRange.setFontWeights(fontWeights);
      SpreadsheetUtils.setRowColors(rColRange, null, false, this.RECON_COL_COLOR, false);

      // Set the text color of the 'internal only' columns to light gray
      const internalRange = reconRange.offset(-1, IDX_RECON_TXN_ID, rowCount + 1, 2);
      SpreadsheetUtils.setRowColors(internalRange, null, false, Const.COLOR_TXN_INTERNAL_FIELD, true);

      // Set current cell to first R cell
      const rCell = reconRange.offset(0, IDX_RECON_RECONCILE, 1, 1);
      rCell.activate();

      // Save reconcile params in a cell
      const reconcileParams = {
        account,
        accountType,
        mintAccount,
        endDate,
        prevBalance, // make sure balances are float type
        newBalance,
      };
      const paramsJson = JSON.stringify(reconcileParams);
      const paramsCell = sheet.getRange(RECON_ROW_TARGET, RECON_COL_SAVED_PARAMS, 1, 1);
      paramsCell.setWrap(false);
      paramsCell.setBackground(Const.NO_COLOR);
      paramsCell.setFontColor(Const.NO_COLOR);
      paramsCell.setValue(paramsJson);

    } catch (e) {
        Debug.log(Debug.getExceptionInfo(e));
        toast('Exception: ' + Debug.getExceptionInfo(e), 'Reconcile', 15);
    }

  },

  autoReconcileTransactions() {
    if (!this.isReconcileAlreadyInProgress()) {
      toast('No reconcile is in progress.', 'Reconcile');
      return;
    }

    const reconcileParams = this.getReconcileParams();
    const endDate = new Date(reconcileParams.endDate);

    const {
      IDX_RECON_DATE,
      IDX_RECON_TXN_ID,
      IDX_RECON_RECONCILE,
      SHEET_NAME_RECONCILE,
    } = Const;
    const reconRange = Utils.getDataRange(SHEET_NAME_RECONCILE, IDX_RECON_TXN_ID + 1);

    let setRecon = false;
    let reconValues = reconRange.getValues();
    for (let reconRow of reconValues) {
      if (reconRow[IDX_RECON_DATE] <= endDate) {
        reconRow[IDX_RECON_RECONCILE] = 'r';
        setRecon = true;
      }
    }

    if (setRecon) {
      let rColValues = reconValues.map(row => [row[IDX_RECON_RECONCILE]]);
      let rColRange = reconRange.offset(0, IDX_RECON_RECONCILE, rColValues.length, 1);
      rColRange.setValues(rColValues);
    }

    this.checkIfReconcileAmountsMatch(reconRange.getSheet(), true, true);
  },

  finishReconcile() {
    if (!this.isReconcileAlreadyInProgress()) {
      toast('No reconcile is in progress.', 'Reconcile');
      return false;
    }

    const {
      IDX_TXN_DATE,
      IDX_TXN_ACCOUNT,
      IDX_TXN_CLEAR_RECON,
      IDX_TXN_ID,
      IDX_TXN_PARENT_ID,
      IDX_RECON_DATE,
      IDX_RECON_TXN_ID,
      IDX_RECON_MERCHANT,
      IDX_RECON_AMOUNT,
      IDX_RECON_RECONCILE,
      IDX_RECON_SPLIT_FLAG,
      RECON_ROW_FINISH_MSG,
      SHEET_NAME_RECONCILE,
      EDITTYPE_EDIT
    } = Const;
    const reconRange = Utils.getDataRange(SHEET_NAME_RECONCILE, IDX_RECON_TXN_ID + 1);
    const sheet = reconRange.getSheet();

    if (!this.checkIfReconcileAmountsMatch(sheet, false, false)) {
      toast('Reconcile is not finished. Amounts do not match.', 'Reconile');
      const finishReconcilingCell = sheet.getRange(RECON_ROW_FINISH_MSG, IDX_RECON_RECONCILE + 1, 1, 1);
      finishReconcilingCell.setValue('');
      return false;
    }

    try
    {
      toast('Finishing ...', 'Reconcile', 5);

      const reconcileParams = this.getReconcileParams();

      Debug.log(`Starting reconcile of regular (non-split) txn rows.`);
      // Get just the reconciled txns
      // Loop through backwards so the index 'i' doesn't get messed up if we delete an entry
      let reconValuesFull = reconRange.getValues();
      let reconValues = [];
      let splitValues = [];
      for (let reconRow of reconValuesFull) {
        if (reconRow[IDX_RECON_RECONCILE].toUpperCase() === 'R') {
          if (reconRow[IDX_RECON_SPLIT_FLAG] === 'S') {
            splitValues.push(reconRow);
          } else {
            reconValues.push(reconRow);
          }
        }
      }

      let reconLen = reconValues.length;
      toast(`Applying ${reconLen} reconciled transaction changes.`, 'Reconcile', 90);

      const txnRange = Utils.getTxnDataRange();
      const txnValues = txnRange.getValues();
      const txnDataLen = txnValues.length;

      // Mark transactions as reconciled, 'R' and highlight them
      Sheets.TxnData.clearRowHighlights();
      const highlightStartCol = txnRange.getSheet().getFrozenColumns();

      for (let i = 0; i < reconLen; ++i) {
        const txnRow = Sheets.TxnData.findTxnRowUnsorted(txnValues, txnDataLen, [IDX_TXN_ID], [ reconValues[i][IDX_RECON_TXN_ID] ]);
        if (txnRow < 0) {
          throw new Error(Utilities.formatString('Transaction cannot be found: %s, %s, $%f', reconValues[i][IDX_RECON_DATE], reconValues[i][IDX_RECON_MERCHANT], reconValues[i][IDX_RECON_AMOUNT]));
        }
        const txnReconCell = txnRange.offset(txnRow, IDX_TXN_CLEAR_RECON, 1, 1);
        txnReconCell.setValue('R');
        // TODO: Make this more efficient
        Sheets.TxnData.validateTransactionEdit(txnRange.getSheet(), txnRange.getRow() + txnRow, IDX_TXN_CLEAR_RECON + 1, IDX_TXN_CLEAR_RECON + 1, EDITTYPE_EDIT);
        if (Debug.traceEnabled) Debug.trace('Reconciling txn: ' + reconValues[i][IDX_RECON_TXN_ID]);

        // Highlight the row
        // (disabled to speed up performance)
        //const txnReconRowRange = txnRange.offset(txnRow, highlightStartCol, 1, Const.IDX_TXN_MATCHES - highlightStartCol);
        //SpreadsheetUtils.setRowColors(txnReconRowRange, null, false, this.RECON_ROW_COLOR, false);
      }

      // Mark split transactions as reconciled, 'R'
      const splitLen = splitValues.length;
      Debug.log(`Finished regular rows. Starting ${splitLen} split txns`);

      for (let i = 0; i < splitLen; ++i) {
        const splitRow = splitValues[i];
        const txnRows = Sheets.TxnData.findAllTxnRowsUnsorted(txnValues, txnDataLen, [IDX_TXN_PARENT_ID], [ splitRow[IDX_RECON_TXN_ID] ]);
        if (txnRows.length == 0) {
          throw new Error(Utilities.formatString('Child transactions of parent %d cannot be found: %s, %s, $%f', splitRow[IDX_RECON_TXN_ID], splitRow[IDX_RECON_DATE], splitRow[IDX_RECON_MERCHANT], splitRow[IDX_RECON_AMOUNT]));
        }

        if (Debug.enabled) Debug.log(`Reconciling ${txnRows.length} child split txn rows (parent id: ${splitRow[IDX_RECON_TXN_ID]})`);
        for (let j = 0; j < txnRows.length; ++j) {
          const txnReconCell = txnRange.offset(txnRows[j], IDX_TXN_CLEAR_RECON, 1, 1);
          txnReconCell.setValue('R');
          Sheets.TxnData.validateTransactionEdit(txnRange.getSheet(), txnRange.getRow() + txnRows[j], IDX_TXN_CLEAR_RECON + 1, IDX_TXN_CLEAR_RECON + 1, EDITTYPE_EDIT);

          // Highlight the row
          // (disabled to speed up performance)
          //const txnReconRowRange = txnRange.offset(txnRows[j], highlightStartCol, 1, Const.IDX_TXN_MATCHES - highlightStartCol);
          //SpreadsheetUtils.setRowColors(txnReconRowRange, null, false, this.RECON_ROW_COLOR, false);
        }
      }

      // Insert dummy row that records the reconcile date and balance
      toast('Creating a "Reconcile transaction" so you can track your reconcile history in Mint.', 'Reconcile', 90);
      Debug.log(`Finished reconciling rows. Creating dummy "reconcile txn"`);

      reconcileParams.newBalance *= 1.00; // Make sure balance is a 'float'
      const endDate = new Date(reconcileParams.endDate);
      const merchant = Utilities.formatString(Const.RECON_RECORD_MERCHANT_FMT, reconcileParams.account);
      const memo = Utilities.formatString(Const.RECON_RECORD_MEMO_FMT, reconcileParams.newBalance);
      const propsJson = Utilities.formatString(Const.RECON_RECORD_PROPS_FMT, reconcileParams.newBalance);

      Sheets.TxnData.insertNewTransaction(Const.EDITTYPE_NEW, 1, endDate, reconcileParams.account,
                                                 merchant, this.RECON_RECORD_AMOUNT, 'Financial', '', 'R',
                                                 memo, null, propsJson, reconcileParams.mintAccount);

      // Sort the transactions by date then account so all of the reconcile txns are displayed together
      txnRange.sort([{column: IDX_TXN_ACCOUNT + 1, ascending: true}, {column: IDX_TXN_DATE + 1, ascending: false}]);

      // Add dummy reconcile txn to txnValues array too, then sort the values and determine where the dummy row is.
      // All this so we can select the row so the reconciled txns are visible.
      let txnRow = [];
      txnRow[Const.IDX_TXN_ACCOUNT] = reconcileParams.account;
      txnRow[Const.IDX_TXN_DATE] = endDate;
      txnRow[Const.IDX_TXN_EDIT_STATUS] = 'N';
      txnValues.push(txnRow);
      txnValues.sort(function(a, b) {
        const aAccount = `${a[IDX_TXN_ACCOUNT]}`;
        const bAccount = `${b[IDX_TXN_ACCOUNT]}`;
        if (aAccount < bAccount) { return -1; }
        if (aAccount > bAccount) { return 1; }
        if (a[IDX_TXN_DATE] > b[IDX_TXN_DATE]) { return -1; }
        if (a[IDX_TXN_DATE] < b[IDX_TXN_DATE]) { return 1; }
        return 0;
      });

      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Const.SHEET_NAME_TXNDATA).activate();

      const selectRow = Sheets.TxnData.findTxnRowUnsorted(txnValues, txnValues.length, [Const.IDX_TXN_DATE, Const.IDX_TXN_EDIT_STATUS, Const.IDX_TXN_ACCOUNT], [endDate, 'N', reconcileParams.account], -1, true);
      if (selectRow >= 0) {
//        --selectRow; //HACK: Subtract 1 from selectRow so it highlights the right row. No why idea selectRow isn't exactly right ...
        const dummyRowRange = txnRange.offset(selectRow - 1, highlightStartCol, 1, Const.IDX_TXN_MATCHES - highlightStartCol);
        SpreadsheetUtils.setRowColors(dummyRowRange, null, false, this.RECON_ROW_COLOR, false);
        SpreadsheetApp.getActiveSpreadsheet().setActiveRange(dummyRowRange);
      }

      toast('Finished. Remember to save the updated txns.', 'Reconcile', 5);
      Debug.log(`Reconcile complete`);

      // Reconcile is complete. Clear the reconcile sheet (rows and header area).
      this.initReconcileSheet(sheet, null, 0, 0);

      // Prompt the user to save the changes to Mint
      // DISABLED: because Spreadsheet.show() is not allowed here anymore? Use buttons instead.
      //Utils.getPrivateCache().put(Const.CACHE_RECONCILE_SUBMIT_PARAMS, paramsJson);
      //const htmlOutput = HtmlService.createTemplateFromFile('reconcile_submit.html').evaluate();
      //htmlOutput.setTitle('Reconcile').setHeight(200).setWidth(300).setSandboxMode(HtmlService.SandboxMode.IFRAME);
      //SpreadsheetApp.getActiveSpreadsheet().show(htmlOutput);

    } catch (e) {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox('Exception: ' + Debug.getExceptionInfo(e));
    }
  },
  
  /**
   * submitReconciledTransactions - Called from reconcile_submit.html
   * args = reconcileParams (see above)
   * -- NOT USED -- We let user choose when to save reconciled txns instead
   */
  submitReconciledTransactions(args) {
    if (Utils.checkDemoMode()) return;

    try
    {
      if (Debug.enabled) Debug.log('Submitting reconcile txns for mint account "%s"', args.mintAccount);
      const cookies = Mint.Session.getCookies();
      if (cookies) {
        Sheets.TxnData.saveModifiedTransactions(args.mintAccount, true);

        // Reconciled txns have been saved. Clear the txn highlights.
        Sheets.TxnData.clearRowHighlights();
      }
    } catch (e) {
        Debug.log(Debug.getExceptionInfo(e));
        toast('Exception: ' + Debug.getExceptionInfo(e), 'Reconcile', 15);
    }
  },

  onReconcileSheetEdit(e)
  {
    const sheet = e.range.getSheet();
    const editColFirst = e.range.getColumn();
    const editRowFirst = e.range.getRow();
    const firstDataRow = e.range.getSheet().getFrozenRows() + 1;

    if (editRowFirst === Const.RECON_ROW_FINISH_MSG && editColFirst === Const.IDX_RECON_RECONCILE + 1) {

      if (e.value && e.value.toUpperCase() === 'X') {
        this.cancelReconcile(true);
      }
      else if (e.value && e.value.toUpperCase() === 'R') {
        // User is done reconciling this account
        this.finishReconcile();
      }
      else {
        // User didn't enter 'R'. Just clear it.
        const finishReconcilingCell = sheet.getRange(Const.RECON_ROW_FINISH_MSG, Const.IDX_RECON_RECONCILE + 1, 1, 1);
        finishReconcilingCell.setValue('');
      }
    }
    else if ((editRowFirst >= firstDataRow && editColFirst === Const.IDX_RECON_RECONCILE + 1) ||
               (editRowFirst === Const.RECON_ROW_TARGET && editColFirst === Const.IDX_RECON_AMOUNT + 1)) {

      if (editColFirst === Const.IDX_RECON_RECONCILE + 1) {
        const reconCell = e.range.offset(0, 0, 1, 1);
        const rowOffset = Math.min(0, e.range.getLastRow() - e.range.getRow());
        const isReconciled = (reconCell.getValue().toUpperCase() === 'R');
        if (isReconciled) {
          // Highlight reconciled rows
          let range = e.range.offset(rowOffset, -Const.IDX_RECON_RECONCILE, e.range.getNumRows(), Const.IDX_RECON_TXN_ID + 1);
          SpreadsheetUtils.setRowColors(range, null, false, this.RECON_ROW_COLOR, false);
        } else {
          // Restore background of just the R column
          let range = e.range.offset(rowOffset, -Const.IDX_RECON_RECONCILE, e.range.getNumRows(), Const.IDX_RECON_TXN_ID + 1);
          SpreadsheetUtils.setRowColors(range, null, false, Const.NO_COLOR, false);
          range = e.range.offset(rowOffset, 0, e.range.getNumRows(), 1);
          SpreadsheetUtils.setRowColors(range, null, false, this.RECON_COL_COLOR, false);
        }
      }

      //const reconRange = sheet.getDataRange();
      //if (Debug.enabled) Debug.log(Utilities.formatString('Reconcile Sheet range: (%s, %s), (%s,%s)', reconRange.getRow(), reconRange.getColumn(), reconRange.getNumRows(), reconRange.getNumColumns()));
      this.checkIfReconcileAmountsMatch(sheet, true, true);
    }
  },

  checkIfReconcileAmountsMatch(sheet, changeCellBackgrounds, showFinishMsg) {
    let amountsMatch = false;
    const reconTargetCell = sheet.getRange(Const.RECON_ROW_TARGET, Const.IDX_RECON_AMOUNT + 1, 1, 1);
    const reconTotalCell = sheet.getRange(Const.RECON_ROW_SUM, Const.IDX_RECON_AMOUNT + 1, 1, 1);
    const target = reconTargetCell.getValue();
    const total = reconTotalCell.getValue();
    
    let color = Const.NO_COLOR;
    // if (Debug.enabled) Debug.log(Utilities.formatString('target: %f, total: %f', target, total));
    if (0 == Math.round(Math.abs(total * 100)) - Math.round(Math.abs(target * 100))) {
      Debug.log(`Reconcile amount matches target (${target})`);
      amountsMatch = true;
      color = '#aaeeaa'; // light green
    }

    if (changeCellBackgrounds) {
      reconTargetCell.setBackground(color);
      reconTotalCell.setBackground(amountsMatch ? color : this.RECON_ROW_COLOR);
    }

    if (showFinishMsg) {
      const finishReconcilingMsgCell = sheet.getRange(Const.RECON_ROW_FINISH_MSG, Const.RECON_COL_FINISH_MSG, 1, 1);

      let msgToFinish = '';
      if (amountsMatch) {
        msgToFinish = Const.RECON_MSG_FINISH;
      }
      finishReconcilingMsgCell.setValue(msgToFinish);
    }

    return amountsMatch;
  },

  initReconcileSheet(sheet, account, prevBalance, newBalance, endDate) {
    const endDateStr = (endDate instanceof Date ? endDate.toISOString().split('T')[0] : endDate || '');
    const reconRange = Utils.getDataRange(Const.SHEET_NAME_RECONCILE, Const.IDX_RECON_TXN_ID + 1);
    reconRange.clear();

    const reconAccount = sheet.getRange(Const.RECON_ROW_TITLE, Const.RECON_COL_ACCOUNT, 1, 1);
    const reconTargetCell = sheet.getRange(Const.RECON_ROW_TARGET, Const.IDX_RECON_AMOUNT + 1, 1, 1);
    const reconTotalCell = sheet.getRange(Const.RECON_ROW_SUM, Const.IDX_RECON_AMOUNT + 1, 1, 1);
    const finishReconcilingMsgCell = sheet.getRange(Const.RECON_ROW_FINISH_MSG, Const.RECON_COL_FINISH_MSG, 1, 1);
    const paramsCell = sheet.getRange(Const.RECON_ROW_TARGET, Const.RECON_COL_SAVED_PARAMS, 1, 1);

    reconAccount.setValue(account);

    const targetAmount = Math.round((newBalance - prevBalance)*100)/100; // Truncate to cents
    reconTargetCell.setValue(targetAmount);
    reconTargetCell.setBackground(Const.NO_COLOR);
    reconTargetCell.setNumberFormat('$0.00');
    if (account == null) {
      reconTargetCell.clearNote();
    } else {
      reconTargetCell.setNote(Utilities.formatString('Prev. Balance: $%f\nNew Balance:   $%f\nTarget amount: $%f\nEnd date:          %s', prevBalance, newBalance, targetAmount, endDateStr));
    }

    reconTotalCell.setBackground(Const.NO_COLOR);
    reconTotalCell.setNumberFormat('$0.00');

    finishReconcilingMsgCell.setValue('');

    paramsCell.setValue('');
  },

  Window: {
    show() {
      try {
        const mintAccounts = Mint.getMintAccounts(null);
        if (mintAccounts == null || mintAccounts.length === 0) {
          Browser.msgBox('No Mint accounts were found. Make sure you have imported your transactions from Mint.');
          return;
        }

        if (Debug.enabled) Debug.log(mintAccounts);

        const acctInfoMap = Sheets.AccountData.getAccountInfoMap();
        if (!acctInfoMap) {
          Browser.msgBox('No accounts were found. Make sure you have imported your accounts from Mint.');
          return;
        }

        const args = { mintAccounts , acctInfoMap };
        Utils.getPrivateCache().put(Const.CACHE_RECONCILE_WINDOW_ARGS, JSON.stringify(args), 60);

        const htmlOutput = HtmlService.createTemplateFromFile('reconcile_start.html').evaluate();
        htmlOutput.setTitle("Reconcile an Account").setHeight(255).setWidth(375).setSandboxMode(HtmlService.SandboxMode.IFRAME);
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        if (ss != null) ss.show(htmlOutput);
      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox(e);
      }
    },

    //-------------------------------------------------------------------------------------
    // Event handlers

    onOkClicked: function(e)
    {

      const uiApp = UiApp.getActiveApplication();
      uiApp.close();

      const accountInfo = e.parameter.accountInfo.split(Const.DELIM_2);
      Reconcile.continueReconcile(accountInfo[0], accountInfo[1], e.parameter.mintAccount, e.parameter.endDate, e.parameter.previousBalance, e.parameter.newBalance);

      return uiApp;
    },

    onCancelClicked: function(e)
    {
      const uiApp = UiApp.getActiveApplication();
      uiApp.close();
      return uiApp;
    },

    onAccountChanged: function(e)
    {
      const uiApp = UiApp.getActiveApplication();
      const prevousBalanceField = uiApp.getElementById('previousBalance');
      const prevousReconDateField = uiApp.getElementById('previousReconDate');

      const accountInfo = e.parameter.accountInfo.split(Const.DELIM_2);
      const prevBalance = (accountInfo[2] || accountInfo[2] === 0 ? String(accountInfo[2]) : null);
      const prevDate = (accountInfo[3] ? String(accountInfo[3]) : 'None');
      //Debug.log('Previous reconcile info:  balance: %s, date: %s', prevBalance, prevDate);

      prevousBalanceField.setValue(prevBalance);
      prevousReconDateField.setText(prevDate);
      return uiApp;
    }
  },
};
