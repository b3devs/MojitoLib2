'use strict';
/*
 * Copyright (c) 2013-2019 b3devs@gmail.com
 * MIT License: https://spdx.org/licenses/MIT.html
 */

import {Const} from './Constants.js'
import {Mint} from './MintApi.js';
import {Sheets} from './Sheets.js';
import {Utils, Settings, EventServiceX, toast} from './Utils.js';
import {Upgrade} from './Upgrade.js';
import {Debug} from './Debug.js';


export const Ui = {

  Menu: {

    setupMojitoMenu: function(isOnOpen) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();

      // Add a custom "Mojito" menu to the active spreadsheet.
      let entries = [];

//      entries.push({ name: "Show / hide sidebar",           functionName: "onMenu_" + Const.ID_TOGGLE_SIDEBAR });
      entries.push({ name: "Connect Mojito to your active Mint session",  functionName: "onMenu_" + Const.ID_SET_MINT_AUTH });
      entries.push(null); // separator

      entries.push({ name: "Sync all with Mint",            functionName: "onMenu_" + Const.ID_SYNC_ALL_WITH_MINT });
      entries.push({ name: "Sync: Import txn data",         functionName: "onMenu_" + Const.ID_IMPORT_TXNS });
      entries.push({ name: "Sync: Import account balances", functionName: "onMenu_" + Const.ID_IMPORT_ACCOUNT_DATA});
      entries.push({ name: "Sync: Save txn changes",        functionName: "onMenu_" + Const.ID_UPLOAD_CHANGES});
      entries.push(null); // separator
      entries.push({ name: "Reconcile an account",          functionName: "onMenu_" + Const.ID_RECONCILE_ACCOUNT});
      // Don't show following menu item until it's implemented
      //entries.push(null); // separator
      //entries.push({ name: "Check for Mojito updates",      functionName: "onMenu_" + Const.ID_CHECK_FOR_UPDATES});

      if (Debug.enabled) {
        entries.push(null); // separator
        entries.push({ name: "Display log window", functionName: "onMenu_displayLogWindow" });
      }

      if (isOnOpen) {    
        ss.addMenu("Mojito", entries);
      } else {
        ss.updateMenu("Mojito", entries);
      }
    },

    onMenu: function(id) {
      switch (id) {
        case Const.ID_SYNC_ALL_WITH_MINT:
          this.onSyncAllWithMint();
          break;

        case Const.ID_IMPORT_TXNS:
          this.onImportTxnData();
          break;

        case Const.ID_IMPORT_ACCOUNT_DATA:
          this.onImportAccountData();
          break;

        case Const.ID_UPLOAD_CHANGES:
          this.onUploadChanges();
          break;

        case Const.ID_RECONCILE_ACCOUNT:
          this.onReconcileAccount();
          break;

        case Const.ID_CANCEL_RECONCILE:
          this.onCancelReconcile();
          break;

        case Const.ID_CHECK_FOR_UPDATES:
          this.onCheckForUpdates();
          break;

        case Const.ID_TOGGLE_SIDEBAR:
          //this.onToggleSidebar();
          break;

        case Const.ID_SET_MINT_AUTH:
          Mint.Session.showManualMintAuth();
          break;

        case 11: {
          toast('test call 11')
          break;
        }
        default:
          if (Debug.enabled) Debug.log("Unknown menu id: " + id);
          break;
      }
    },
    
    //-----------------------------------------------------------------------------
    onSyncAllWithMint: function()
    {
      this.syncWithMint(true, true, false, true, true, true);
    },
    
    //-----------------------------------------------------------------------------
    onImportTxnData: function()
    {
      this.syncWithMint(false, false, true, true, false, false);
    },

    //-----------------------------------------------------------------------------
    onImportAccountData: function()
    {
      this.syncWithMint(false, false, false, false, true, false);
    },

    //-----------------------------------------------------------------------------
    onUploadChanges: function()
    {
      this.syncWithMint(false, true, false, false, false, false);
    },

    //-----------------------------------------------------------------------------
    onReconcileAccount: function()
    {
      Reconcile.startReconcile();
      Sheets.About.turnOffAuthMsg();
    },

    //-----------------------------------------------------------------------------
    onCancelReconcile: function()
    {
      Reconcile.cancelReconcile();
      // Refresh the Mojito menu
      Ui.Menu.setupMojitoMenu(false);
      Sheets.About.turnOffAuthMsg();
    },

    //-----------------------------------------------------------------------------
    onCheckForUpdates: function()
    {
      // TODO: Open web page showing release history
    },

    //-----------------------------------------------------------------------------
    onToggleSidebar: function() {
      // Experimenting with this...
      const htmloutput = HtmlService.createTemplateFromFile('sidebar.html').evaluate()
                        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
                        .setTitle('Mojito sidebar')
                        .setWidth(300);
      SpreadsheetApp.getUi().showSidebar(htmloutput);
    },
    
    //-----------------------------------------------------------------------------
    onDisplayLogWindow: function()
    {
      Debug.displayLogWindow();
    },
    
    /////////////////////////////////////////////////////////////////////////////
    // Helpers

    syncWithMint: function(syncAll, saveEdits, promptSaveEdits, importTxns, importAccountBal, importCatsTags)
    {
      // If we're running int "demo mode" then don't do anything
      if (Utils.checkDemoMode()) {
        return;
      }

      // Apply Mojito updates, if any. If the spreadsheet version cannot be updated
      // to this MojitoLib version, then abort.
      const version = Upgrade.autoUpgradeMojitoIfApplicable();
      if (version !== Const.CURRENT_MOJITO_VERSION) {
        // Upgrade function already displayed a message. Just return.
        return;
      }

      // Before doing anything, make sure the mint session is still valid.
      // If expired (no cookies), then show message box and abort.
      const cookies = Mint.Session.getCookies(false);
      if (!cookies) {
        Browser.msgBox('Mint authentication expired', 'The Mint authentication token has not been provided or it has expired. Please re-enter the Mint authentication headers then retry the operation.', Browser.Buttons.OK);
        Debug.log('No auth cookies found. Showing mint auth ui.');
        Mint.Session.showManualMintAuth();
        return;
      }

      let saveFailed = false;

      try
      {
        let txnDateRange = null;
        let acctDateRange = null;

        if (syncAll) {
          importAccountBal = !Sheets.AccountData.isUpToDate();
          // If account balances aren't up to date, get the date range for the latest balances
          acctDateRange = (importAccountBal ? Sheets.AccountData.determineImportDateRange() : null);

          // Get date range for latest transactions
          txnDateRange = Sheets.TxnData.determineImportDateRange(Utils.getMintLoginAccount());
        }

        // Wait for mint to get data from financial institutions
        //Disabled: Mint doesn't seem to support this any more
        const isMintDataReady = Mint.waitForMintFiRefresh(true);

        if (!isMintDataReady) {
          return;
        }

        if (importAccountBal) {
          Sheets.AccountData.getSheet().activate();

          if (acctDateRange)
          {
            // Download latest account balances
            if (Debug.enabled) Debug.log("Sync account date range: " + Utilities.formatDate(acctDateRange.startDate, "GMT", "MM/dd/yyyy") + " - " + Utilities.formatDate(acctDateRange.endDate, "GMT", "MM/dd/yyyy"));

            const args = {
                startDateTs: acctDateRange.startDate.getTime(),
                endDateTs: acctDateRange.endDate.getTime(),
                replaceExistingData: false
              };

            Ui.AccountBalanceImportWindow.onImport(args);
          } else {
            // No date range specified. Show the account import window
            const args = Sheets.AccountData.determineImportDateRange();
            Ui.AccountBalanceImportWindow.show(args);
         }
        }
        
        if (saveEdits === true || promptSaveEdits === true) {
          const mintAccount = Utils.getMintLoginAccount();

          if (promptSaveEdits === true) {
            const pendingUpdates = Sheets.TxnData.getModifiedTransactionRows(mintAccount);
            if (pendingUpdates != null && pendingUpdates.length > 0) {
              if ("yes" === Browser.msgBox("", "Would you like to save your modified transactions first? (If you click \"No\", your changes may be overwritten.)", Browser.Buttons.YES_NO)) {
                saveEdits = true;
              }
            }
          }

          if (saveEdits === true) {
            Sheets.TxnData.getSheet().activate();

            // Upload any edited txns
            const success = Sheets.TxnData.saveModifiedTransactions(mintAccount, true);
            saveFailed = !success;
          }
        }

        if (importTxns) {
          if (saveFailed) {
            toast("Not all changes were saved. Skipping transaction import.");
            Utilities.sleep(3000);
          }
          else {
            Sheets.TxnData.getSheet().activate();
  
            // Import the "latest" txns, or show import window?
            if (txnDateRange) {
              // Download the latest txns
              if (Debug.enabled) Debug.log("Sync txn date range: " + Utilities.formatDate(txnDateRange.startDate, "GMT", "MM/dd/yyyy") + " - " + Utilities.formatDate(txnDateRange.endDate, "GMT", "MM/dd/yyyy"));
  
              const args = {
                  startDate: txnDateRange.startDate.getTime(),
                  endDate: txnDateRange.endDate.getTime(),
                  replaceExistingData: false,
                };
  
              Mint.TxnData.importTransactions(args, true, !syncAll);
  
            }
            else {
              // No date range specified. Show the txn import window
              
              // Default date range will be year-to-date
              const today = new Date;
              const startDate = new Date(today.getYear(), 0, 1);

              const args = { startDate: startDate, endDate: today };
              Ui.TxnImportWindow.show(args);
            }
          }
        }

        if (importCatsTags) {
          // Download latest categories and tags
          Debug.log("Syncing categories");
          Mint.Categories.import(false);
          Debug.log("Syncing tags");
          Mint.Tags.import(false);
        }

        Sheets.About.turnOffAuthMsg();
      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox(e);
      }
    },


  }, // Menu

  ///////////////////////////////////////////////////////////////////////////////
  // LoginWindow class

  /**
   * @obsolete The login window no longer works because Mint now requires two-factor
   * authentication (via emailed code) if it doesn't recognize the computer where
   * the login request is coming from (which is a Google server in this case).
   */
  LoginWindow: {

    login: function()
    {
      // Only allow one login call at a time
      let loginMutex = Utils.getDocumentLock();
      if (!loginMutex.tryLock(1000))
      {
        toast("Multiple logins are occurring at once. Please wait a moment ...", "", 10);

        if (!loginMutex.tryLock(Const.MINT_LOGIN_START_TIMEOUT_SEC * 1000))
        {
          if (Debug.enabled) Debug.log("Unable to aquire login lock");
          return null;
        }
      }

      try
      {
        Debug.log("Login required");
        Ui.LoginWindow.show();
        
        // The code that follows only exists because we want to open the login window and wait for it
        // to close, either because the user successfully logged in, or the login was cancelled. This sort of
        // "modal dialog" behavior is not supported by Google Apps Script for windows created by container-bound
        // scripts. To complicate matters, the user could click the little "X" in the upper right corner to
        // close the window. There is no way (that I have found) to intercept this action and notify this login
        // script that the login has been canceled. So to handle this case, we have the login window send a 
        // "window ping" event every few seconds so we know it is still open. If we don't see the ping 
        // for 10 seconds, then we assume the user closed the window with the "X". The fact that any of this
        //  code needs to exist is pretty lame. Google Apps Script should support this simple use case.
        let loginFinished = false;
        let loginSucceeded = false;
        let loginWaitEvents = [Const.EVT_MINT_LOGIN_SUCCEEDED, Const.EVT_MINT_LOGIN_FAILED, Const.EVT_MINT_LOGIN_CANCELED, Const.EVT_MINT_LOGIN_WINDOW_PING];
        let timeoutSec = Const.MINT_LOGIN_TIMEOUT_SEC;
        let windowOpened = false;
        let timeoutCount = 0;
        
        while (true) {
          let loginEvent = EventServiceX.waitForEvents(loginWaitEvents, timeoutSec);
          switch (loginEvent) {
              
            case Const.EVT_MINT_LOGIN_SUCCEEDED:
              if (Debug.enabled) Debug.log("Wait event: Login succeeded");
              loginFinished = true;
              loginSucceeded = true;
              break;
              
            case Const.EVT_MINT_LOGIN_CANCELED:
              if (Debug.enabled) Debug.log("Wait event: Login was canceled");
              loginFinished = true;
              loginSucceeded = false;
              break;
              
            case Const.EVT_MINT_LOGIN_FAILED:
              if (Debug.enabled) Debug.log("Wait event: Login failed.");
              break;
              
            case Const.EVT_MINT_LOGIN_WINDOW_PING:
              // Login window is still open, waiting for user to click OK or Cancel. Keep waiting ...
              windowOpened = true;
              if (Debug.traceEnabled) Debug.trace("Wait event: Login window ping");
              break;
              
            default:
              if (windowOpened || timeoutCount > 2)
              {
                if (Debug.enabled) Debug.log("Login timeout: Assuming login window has been closed.");
                toast("Login timed out.", "Mint login", 5);
                loginFinished = true;
                loginSucceeded = false;
              } else {
                ++timeoutCount;
                // don't give up until the window has opened. Could just be a really slow network connection.
                if (Debug.enabled) Debug.log("Login timeout: Ignoring. Window is not open yet.");
              }
              break;
          }
          
          if (loginFinished)
            break;
        }
      }
      finally
      {
        loginMutex.releaseLock();
      }

      return loginSucceeded;
    },

    show: function()
    {
      EventServiceX.clearEvent(Const.EVT_MINT_LOGIN_STARTED);
      EventServiceX.clearEvent(Const.EVT_MINT_LOGIN_CANCELED);

      const htmlOutput = HtmlService.createTemplateFromFile('mint_login.html').evaluate();
      htmlOutput.setHeight(150).setWidth(310).setSandboxMode(HtmlService.SandboxMode.IFRAME);
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Log in to Mint');
    },

    onDoLogin: function(args) {
      let success = false;

      const loginCookies = Mint.Session.loginMintUser(args.email, args.password);
      
      if (loginCookies)
      {
        // Save successful email to settings
        Settings.setSetting(Const.IDX_SETTING_MINT_LOGIN, args.email);

        EventServiceX.triggerEvent(Const.EVT_MINT_LOGIN_SUCCEEDED, {result:"success", cookies: loginCookies});
        success = true;
      }
      else
      {
        EventServiceX.triggerEvent(Const.EVT_MINT_LOGIN_FAILED, null);
      }

      return success;
    },

    onCancel: function() {
      EventServiceX.triggerEvent(Const.EVT_MINT_LOGIN_CANCELED, null);
    },
    
    onWindowPing: function() {
      EventServiceX.triggerEvent(Const.EVT_MINT_LOGIN_WINDOW_PING, null);
    },
  },
  
  ///////////////////////////////////////////////////////////////////////////////
  TxnImportWindow: {
    /**
     *
     * @param dates {{ startDate: Date, endDate: Date }}
     */
    show: function(dates) {
      try
      {
        const args = { startDate: dates.startDate.getTime(), endDate: dates.endDate.getTime() };
        Utils.getPrivateCache().put(Const.CACHE_TXN_IMPORT_WINDOW_ARGS, JSON.stringify(args), 60);

        const htmlOutput = HtmlService.createTemplateFromFile('txn_import.html').evaluate();
        htmlOutput.setTitle("Import Transactions from Mint").setHeight(190).setWidth(250).setSandboxMode(HtmlService.SandboxMode.IFRAME);
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        if (ss != null) ss.show(htmlOutput);
      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox(e);
      }
    },

    /**
     * onImport - Called from txn_import.html
     * args = { startDate: <date>, endDate: <date>, replaceExistingData: <true/false> }
     */
    onImport: function(args) {
      try
      {
        Mint.TxnData.importTransactions(args, true, true);
      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox(e);
      }

      toast("Done");
    }

  },

  ///////////////////////////////////////////////////////////////////////////////
  AccountBalanceImportWindow: {
    show: function(dates) {
      // dates = { startDate: <date>, endDate: <date> }
      try
      {
        const args = { startDate: dates.startDate.getTime(), endDate: dates.endDate.getTime() };
        Utils.getPrivateCache().put(Const.CACHE_ACCOUNT_IMPORT_WINDOW_ARGS, JSON.stringify(args), 60);

        const htmlOutput = HtmlService.createTemplateFromFile('account_balance_import.html').evaluate();
        htmlOutput.setTitle("Import account balances").setHeight(220).setWidth(250).setSandboxMode(HtmlService.SandboxMode.IFRAME);
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        if (ss) ss.show(htmlOutput);
      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox(e);
      }
    },

    /**
     * onImport - Called from account_balance_import.html
     * @param args {{
     *   startDate: Date,
     *   endDate: Date,
     *   replaceExistingData: boolean,
     *   saveReplaceExisting: boolean,
     *   importTodaysBalance: undefined|boolean,
     *   }}
     */
    onImport(args) {
      try
      {
        const { IDX_ACCT_NAME, IDX_ACCT_FINANCIAL_INST, IDX_ACCT_TYPE, IDX_ACCT_ID } = Const;
        // if (Debug.enabled) Debug.log("AccountBalanceImportWindow.onImport(): %s", args.toSource());

        let accountRanges = Utils.getAccountDataRanges(false); // Call Utils impl to get ranges with 1 row if no data exists.
        let balRange = accountRanges.balanceRange;

        // Replace (clear) the existing account data, if requested
        if (args.replaceExistingData === true) {
          const balCount = balRange.getNumRows();

          if (!accountRanges.isEmpty && balCount > 0) {
            const button = Browser.msgBox("Replace existing balances?", "Are  you sure you want to REPLACE the " + balCount + " existing account balance(s)?", Browser.Buttons.OK_CANCEL);
            if (button === "cancel")
              return;
          }

          if (accountRanges.hdrRange != null) {
            accountRanges.hdrRange.clear();
            accountRanges.hdrRange.setWrap(true);
          }
          const range = balRange.offset(0, -1, balRange.getNumRows(), balRange.getNumColumns() + 1);
          range.clear();
        }

        // Fetch the account list
        toast('Fetching account list', 'Account import', 120);
        const sessionHeaders = Mint.Session.getHeaders();
        let acctInfoArray = Mint.AccountData.downloadAccountInfo() || [];
        if (Debug.enabled) Debug.log(`Found ${acctInfoArray.length} accounts(s) for balance import. Filtering out hidden and closed accounts.`);
        // Filter out hidden and closed accounts from the list
        acctInfoArray = acctInfoArray.filter((acct) => {
          if (!acct.isVisible)
          {
            if (Debug.enabled) Debug.log(`Ignoring hidden account '${acct.name}'`);
            return false;
          }
          if (acct.isClosed || acct.isDeleted)
          {
            if (Debug.enabled) Debug.log(`Ignoring closed / deleted account '${acct.name}'`);
            return false;
          }
          if (Debug.enabled) Debug.log(`Including account '${acct.name}'`);
          return true;
        });

        if (acctInfoArray.length === 0) {
          toast('No accounts found', 'Account import', 2000);
          return;
        }

        if (args.saveReplaceExisting) {
          Settings.setSetting(Const.IDX_SETTING_REPLACE_ALL_ON_ACCT_IMPORT, args.replaceExistingData);
        }

        let acctBalDataMap = {}; // key is date as string, value is Array of account+data maps

        // If we are not replacing existing data then get the existing account data from the spreadsheed
        if (!args.replaceExistingData && !accountRanges.isEmpty) {
          // Re-order acctInfoArray to match the order chosen by the user (i.e column order in the spreadsheet)
          const existingAcctInfo = accountRanges.hdrRange?.getValues() || [];
          const existingAcctCount = existingAcctInfo[0].length;
          let acctInfoArrayMerged = [];
          // FIXME: Merge is broken, fix it.
          for (let i = 0; i < existingAcctCount; ++i) {
            // Debug.log(`Finding id: ${existingAcctInfo[IDX_ACCT_ID][i]}, name: ${existingAcctInfo[IDX_ACCT_NAME][i]}`)
            // let idx = Utils.findArrayIndex(acctInfoArray, (a) => a.id === existingAcctInfo[IDX_ACCT_ID][i]);
            let idx = Utils.findArrayIndex(acctInfoArray, (a) => {
              // Debug.log(`comparing "${a?.id}" to "${existingAcctInfo[IDX_ACCT_ID][i]}"`);
              return (a?.id === existingAcctInfo[IDX_ACCT_ID][i]);
            });
            if (idx < 0) {
              idx = Utils.findArrayIndex(acctInfoArray, (a) => a?.name === existingAcctInfo[IDX_ACCT_NAME][i]);
            }

            if (idx >= 0) {
              Debug.log(`Index of acct "${acctInfoArray[idx].name}" found at ${idx}`)
              acctInfoArrayMerged.push(acctInfoArray[idx]);
              acctInfoArray[idx] = null;
            }
            else {
              Debug.log(`Existing acct not found: "${existingAcctInfo[IDX_ACCT_NAME][i]}"`);
              acctInfoArrayMerged.push({
                name: existingAcctInfo[IDX_ACCT_NAME][i],
                fiName: existingAcctInfo[IDX_ACCT_FINANCIAL_INST][i],
                type: existingAcctInfo[IDX_ACCT_TYPE][i],
                id: existingAcctInfo[IDX_ACCT_ID][i],
                // Add flag indicating that Mint doesn't know about this acc
                isUnknownAcct: true
              });
            }
          }
          // Add any new accounts to the end
          let acctCount = acctInfoArray.length;
          for (let i = 0; i < acctCount; ++i) {
            if (acctInfoArray[i]) {
              Debug.log(`New acct (idx ${i}): "${existingAcctInfo[IDX_ACCT_NAME][i]}"`);
              acctInfoArrayMerged.push(acctInfoArray[i]);
            }
          }
          acctInfoArray = acctInfoArrayMerged;
          acctCount = acctInfoArray.length;

          // Read the existing account data into acctBalDataMap
          const dateValues = accountRanges.dateRange.getValues().map((entry) => entry[0]);
          const balValues = balRange.getValues();

          const balCount = balValues.length;
          for (let i = 0; i < balCount; ++i) {
            // Make sure length of array at balValues[i] has the same length as acctInfoArray.length
            const currRowLen = balValues[i].length;
            if (currRowLen !== acctCount) {
              balValues[i].length = acctCount;
              // Change new array values from undefined to empty string '' 
              for (let j = currRowLen; j < acctCount; ++j) { balValues[i][j] = ''; }
            }
            // Convert dateValues[i] to date string and assign balance values to it
            // Debug.log(`dateValues[${i}]: "${dateValues[i]}"`);
            acctBalDataMap[dateValues[i].toISOString().split('T')[0]] = balValues[i];
          }
        }

        // Activate the last date cell so user can see new data that is imported.
        let lastDateCell = balRange.offset(balRange.getNumRows() - 1, -1, 1, 1);
        if (lastDateCell) {
          lastDateCell.activate();
        }

        const { startDateTs, endDateTs } = args; // dates are epoch timestamps
        const acctCount = acctInfoArray.length; // this count includes custom columns added by user
        const actualAcctCount = acctInfoArray.reduce((count, acct) => { return count + (acct.id ? 1 : 0)}, 0);
        let actualAcctIndex = 0;

        toast(`Retrieving balances for ${actualAcctCount} account(s)`, 'Account balance import', 120);
        let showToastForEachAcct = false;

        for (let acctIdx = 0; acctIdx < acctCount; ++acctIdx) {
          const timeStart = Date.now();

          const currAcct = acctInfoArray[acctIdx];
          if (!currAcct.id) {
            // Skip any custon inserted columns
            continue;
          }

          if (currAcct.isUnknownAcct) {
            toast(`Skipping unknown account: ${currAcct.name}`);
            continue;
          }

          if (showToastForEachAcct) {
            toast(`Retrieving balances for account ${++actualAcctIndex} of ${actualAcctCount}:  ${currAcct.name}`, "Account balance import", 60);
          }

          const acctWithBalances = Mint.AccountData.downloadBalanceHistory(sessionHeaders, currAcct, startDateTs, endDateTs, args.importTodaysBalance, true);
          if (acctWithBalances.balanceHistoryNotAvailable || !acctWithBalances.balanceHistory) {
            toast(`No balance history found for account: ${currAcct.name}`, "Account balance import", 3);
            continue;
          }

          for (let balanceEntry of currAcct.balanceHistory) {
            // Debug.log(`${balanceEntry.dateStr}   ${balanceEntry.amount}`);
            // If this date doesn't exist in the map yet, add it (with 0 amount for each account)
            if (!acctBalDataMap.hasOwnProperty(balanceEntry.dateStr)) {
              // Pre-allocating an array in JS is non-obvious. Found this solution on stackoverflow:
              // https://stackoverflow.com/questions/1295584/most-efficient-way-to-create-a-zero-filled-javascript-array
              acctBalDataMap[balanceEntry.dateStr] = Array.apply(null, Array(acctCount)).map(Number.prototype.valueOf, 0);
            }

            acctBalDataMap[balanceEntry.dateStr][acctIdx] = balanceEntry.amount;
          }

          const timeElapsed = Date.now() - timeStart;
          if (timeElapsed > 1000) {
            showToastForEachAcct = true;
          }
        }

        // Convert acctInfoArray into a 2D array that we can insert the as
        // account headers into the spreadsheet
        const acctValues = [
          /* IDX_ACCT_NAME */           acctInfoArray.map(a => a.name),
          /* IDX_ACCT_FINANCIAL_INST */ acctInfoArray.map(a => a.fiName),
          /* IDX_ACCT_TYPE */           acctInfoArray.map(a => a.type),
          /* IDX_ACCT_ID */             acctInfoArray.map(a => a.id),
        ];
        const hdrRange = accountRanges.hdrRange.offset(0, 0, acctValues.length, acctValues[0].length);
        hdrRange.setValues(acctValues);

        // Get all of the dates (keys of the map) as an array, and sort them ascending while we're at it
        let dates = Object.keys(acctBalDataMap)?.sort();

        // Populate date column
        let range = balRange.offset(0, -1, dates.length, 1);
        range.setValues(dates.map(d => [d]));

        // Activate the last date cell so user can quickly see the latest balances
        lastDateCell = range.offset(dates.length - 1, 0, 1, 1);
        lastDateCell.activate();

        // Loop through accounts again so we can populate one account (column) at a time.
        // We only do one account at a time because we don't want to overwrite custom columns
        //  containing formulas, etc.
        for (let acctIdx = 0; acctIdx < acctCount; ++acctIdx) {
          const currAcct = acctInfoArray[acctIdx];
          if (!currAcct.id) {
            // Skip any custon inserted columns
            continue;
          }

          if (currAcct.isUnknownAcct) {
            // Skip unknown account
            continue;
          }

          range = balRange.offset(0, acctIdx, dates.length, 1);
          range.setValues(dates.map(d => [ acctBalDataMap[d][acctIdx] ]));
        }
      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox(e);
      }

      toast("Done", "Account balance import");
    }

  }

};
