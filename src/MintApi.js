'use strict';
/*
 * Copyright (c) 2013-2023 b3devs@gmail.com
 * MIT License: https://spdx.org/licenses/MIT.html
 */

import {Const} from './Constants.js';
import {Utils, Settings, toast} from './Utils.js';
import {Sheets} from './Sheets.js';
import {Debug} from './Debug.js';

///////////////////////////////////////////////////////////////////////////////
// Mint class

export const Mint = {

  testFunc() {
    Logger.log('testFunc');
    let txnValues = Mint.TxnData.fetchSplitTransactions('43248151_1281033579_0');
    Debug.log(`split txns: ${JSON.stringify(txnValues, null, '  ')}`);
  },

  ///////////////////////////////////////////////////////////////////////////////
  // Mint.TxnData class

  TxnData: {

    // Mint.TxnData
    importTransactions(options, interactive, showImportCount)
    {
      try
      {
        if (interactive == undefined) {
          interactive = false;
        }
        if (showImportCount == undefined) {
          showImportCount = true;
        }

        if (interactive && showImportCount) toast("Starting ...", "Mint transaction import");

        let range = Utils.getTxnDataRange();

        // If replacing existing txns, prompt user to make sure
        if (options.replaceExistingData && range.getNumRows() > 1) {
          let button = Browser.msgBox("Replace existing transactions?", "Are  you sure you want to REPLACE the " + range.getNumRows() + " existing transaction(s)?", Browser.Buttons.OK_CANCEL);
          if (button === "cancel")
            return;
        }

        // Clear all txn formatting (row colors, text colors, bold, italics, etc.)
        // We will re-apply it after the txns have been imported
        range.clear({formatOnly: true});

        let headers = Mint.Session.getHeaders();
        if (!headers) {
          throw new Error('Mint session has expired');
        }

        if (interactive) {
          if (showImportCount) {
            toast("Retrieving transaction data", "Mint transaction import", 60);
          } else {
            toast("Retrieving the latest transaction data", "Mint transaction import", 60);
          }
        }

        let allTxns = [];
        let offset = 0;
        let progress = 0;
        let mintAccount = Settings.getSetting(Const.IDX_SETTING_MINT_LOGIN);
        let importDate = new Date();

        do
        {
          let startDate = new Date(options.startDate);
          let endDate = new Date(options.endDate);

          let txns = this.downloadTransactions(headers, offset, startDate, endDate);
          if (!txns || txns.length <= 0)
            break;
          
          let txnValues = this.copyDataIntoValueArray(txns, importDate, mintAccount);
          // Use push.apply() for fast, in-place concatenation
          allTxns.push.apply(allTxns, txnValues);

          offset += txns.length;
          
          // Display toast with status every 500 txns
          progress += txns.length;
          if (progress >= 500)
          {
            if (interactive && showImportCount) toast("Downloaded " + offset + " transactions. More coming ...", "Mint transaction import", 10);
            progress -= 500;
            
//            if (offset % 1000 === 0 && "no" === Browser.msgBox("Mint transaction import", String(offset) + " transactions have been downloaded. Continue?", Browser.Buttons.YES_NO)) {
//              break;
//            }
          }

        } while (true);

        if (interactive && showImportCount) toast("Download complete. Importing " + offset + " transactions into Mojito", "Mint transaction import", 60);

        Sheets.TxnData.insertData(allTxns, options.replaceExistingData);

        if (interactive) {
          if (showImportCount) {
            toast("A total of " + offset + " transactions were imported.", "Mint transaction import", 5);
          } else {
            toast("Transactions imported.", "Mint transaction import", 5);
          }
        }
      }
      catch(e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox("Error: " + e.toString());
      }

    },

    // Mint.TxnData
    downloadTransactions(sessionHeaders, offset, startDate, endDate)
    {
      if (!sessionHeaders)
        return null;

      const limit = 200; // arbitrary chunk size of txns to fetch
      let startDateString = startDate.toISOString().split('T')[0];
      let endDateString = endDate.toISOString().split('T')[0];
      
      let url = "https://mint.intuit.com/pfm/v1/transactions";
      url += `?limit=${limit}&offset=${offset}&fromDate=${startDateString}&toDate=${endDateString}`;
      Debug.log(`Downloading transactions: ${url}`);

      let headers = {...sessionHeaders,
        Accept: "application/json"
      };
      // Debug.log(`headers: ${JSON.stringify(headers)}`);

      let options = {
        method: "GET",
        headers
      };
    
      let response = null;
      let txnData = [];
      
      try
      {
        response = UrlFetchApp.fetch(url, options);
        if (response && (response.getResponseCode() === 200))
        {
          //let respHeaders = response.getAllHeaders();
          //Debug.log(respHeaders.toSource());
          
          let respBody = response.getContentText();
          let respCheck = Mint.checkJsonResponse(respBody);
          if (respCheck.success)
          {
            let results = JSON.parse(respBody);
            //Debug.log("Results: " + results.toSource());

            if (results.Transaction) {
              // Debug.log(`Transaction json: ${JSON.stringify(results, null, '  ')}`);
              Debug.assert(Array.isArray(results.Transaction), 'results.Transaction object is an array');
              txnData = results.Transaction;
              Debug.assert(!!txnData, "txnData !== null");
            }
          }
          else
          {
            throw new Error("Data retrieval failed. " + respBody);
          }
        }
        else
        {
          toast("Data retrieval failed. HTTP status code: " + response.getResponseCode(), "Download error");
        }
      }
      catch(e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox("Error: " + e.toString());
      }
      
      return txnData;
    },

    // Mint.TxndData
    fetchSplitTransactions(txnId) {
      let url = `https://mint.intuit.com/pfm/v1/transactions/${txnId}`;
      Debug.log(`Downloading split transactions: ${url}`);

      let sessionHeaders = Mint.Session.getHeaders();
      if (!sessionHeaders) {
        throw new Error('Mint session has expired');
      }

      let headers = {...sessionHeaders,
        Accept: 'application/json'
      };
      // Debug.log(`headers: ${JSON.stringify(headers)}`);

      let options = {
        method: 'GET',
        headers
      };
    
      let response = null;
      let txnData = null;

      try
      {
        response = UrlFetchApp.fetch(url, options);
        if (response && (response.getResponseCode() === 200))
        {
          let respBody = response.getContentText();
          let result = JSON.parse(respBody);
          // Debug.log(`result: ${JSON.stringify(result, null, '  ')}`);
          let importDate = new Date();
          let mintAccount = Utils.getMintLoginAccount();
          txnData = this.copyDataIntoValueArray((Array.isArray(result) ? result : [result]), importDate, mintAccount);
        } else {
          throw new Error(`Failed to fetch transaction(s) for id "${txnId}". Status code ${response.getResponseCode()}, ${response.getContentText()}`);
        }
      }
      catch(e) {
        Debug.log(Debug.getExceptionInfo(e));
      }

      return txnData;
    },

    // Mint.TxnData
    updateTransaction(txnId, payload, editType)
    {
      let result = { success: false, responseJson: null };

      sessionHeaders = Mint.Session.getHeaders();

      try
      {
        let response = null;
 
        if (Debug.enabled) Debug.log(`updateTransaction: txnId: ${txnId}, payload: ${JSON.stringify(payload)}`);

        let headers = {...sessionHeaders,
          'content-type': 'application/json',
          accept: 'application/json',
          origin: 'https://mint.intuit.com',
          referer: 'https://mint.intuit.com/transactions'
        };

        let options = {
          method: (editType === Const.EDITTYPE_NEW ? 'POST' : 'PUT'),
          headers,
          payload: JSON.stringify(payload),
          followRedirects: false,
          muteHttpExceptions: true
        };

        response = UrlFetchApp.fetch(`https://mint.intuit.com/pfm/v1/transactions/${txnId}`, options);
        let respBody = '';
        const respCode = response.getResponseCode();
        if (respCode === 204) {
          if (Debug.enabled) Debug.log('updateTransaction: succeeded');
          result.success = true;
        }
        else {
          respBody = response.getContentText();

          if (respCode === 201) {
            result.success = true;
          }
          else if (respCode === 200) {
            //let respHeaders = response.getAllHeaders();
            //Debug.log("Response headers: " + respHeaders.toSource());

            respBody = response.getContentText();
            if (Debug.enabled) Debug.log("updateTransaction: Response Body: " + respBody);
            let respCheck = Mint.checkJsonResponse(respBody);
            if (respCheck.success) {
              result.success = true;
              result.responseJson = JSON.parse(respBody);
            }
          }
        }

        if (!result.success) {
          throw new Error("Update failed. Status code: " + response.getResponseCode() + ", body:  " + respBody);
        }
      }
      catch (e) {
        Debug.log(Debug.getExceptionInfo(e));
        toast("Update failed. Exception encountered: " + e.toString());
//        Browser.msgBox("Error: " + e.toString());
      }

      return result;
    },

    // Mint.TxnData
    copyDataIntoValueArray(txnData, importDate, mintAccount)
    {
      const {
        IDX_TXN_DATE,
        IDX_TXN_EDIT_STATUS,
        IDX_TXN_ACCOUNT,
        IDX_TXN_MERCHANT,
        IDX_TXN_AMOUNT,
        IDX_TXN_CATEGORY,
        IDX_TXN_TAGS,
        IDX_TXN_CLEAR_RECON,
        IDX_TXN_MEMO,
        IDX_TXN_MATCHES,
        IDX_TXN_STATE,
        IDX_TXN_MINT_ACCOUNT,
        IDX_TXN_ORIG_MERCHANT_INFO,
        IDX_TXN_ID,
        IDX_TXN_PARENT_ID,
        IDX_TXN_CAT_ID,
        IDX_TXN_TAG_IDS,
        IDX_TXN_MOJITO_PROPS,
        IDX_TXN_YEAR_MONTH,
        IDX_TXN_ORIG_AMOUNT,
        IDX_TXN_IMPORT_DATE,
        IDX_TXN_LAST_COL,
        TXN_STATUS_PENDING,
        TXN_STATUS_SPLIT,
        DELIM
      } = Const;
      const TXN_COL_COUNT = IDX_TXN_LAST_COL + 1;
      const numRows = txnData.length;
      const tagCleared = Mint.getClearedTag();
      const tagReconciled = Mint.getReconciledTag();

      // local helper function
      const self = this;
      function createTxnFromData(dataJson, childItem, parentId, mintAccount, importDate) {
        let txnRow = new Array(TXN_COL_COUNT);

        childItem = childItem || dataJson;
        const txnDate = new Date(dataJson.date);
        const amount = childItem.amount;

        let cleared = false;
        let reconciled = false;

        let tagArray = childItem.tagData?.tags || [];
        let tags = '';
        let tagIds = '';
        for (let t of tagArray)
        {
          // Show the 'cleared' and 'reconciled' tags in c/R column
          let tagName = t.name;
          if (tagName === tagCleared) {
            cleared = true;
          }
          else if (tagName === tagReconciled) {
            reconciled = true;
          }
          else {
            tags += tagName + DELIM;
          }
          // Include all tag IDs
          tagIds += t.id + DELIM;
        }
        
        // See if the memo field contains any Mojito properties
        // If so, move the props to a separate column to avoid confusing the user
        let memo = childItem.notes || '';
        let props = null;
        let propsJson = null;

        if (memo) {
          let extractedParts = self.extractPropsFromString(memo);
          if (extractedParts) {
            memo = extractedParts.text;
            props = extractedParts.props;
            propsJson = extractedParts.propsJson;
          }
        }
        
        // Determine 'state' column
        let state = null;
        if ((childItem.isPending === true || childItem.manualTransactionType === 'PENDING')
            && (!props || props.pending !== 'ignore')) {
          state = TXN_STATUS_PENDING;
        } else if (parentId) {
          state = TXN_STATUS_SPLIT;
        }

        txnRow[IDX_TXN_DATE] = txnDate;
        txnRow[IDX_TXN_EDIT_STATUS] = null;
        txnRow[IDX_TXN_ACCOUNT] = dataJson.accountRef?.name || 'Unknown';
        txnRow[IDX_TXN_MERCHANT] = childItem.description;
        txnRow[IDX_TXN_AMOUNT] = amount;
        txnRow[IDX_TXN_CATEGORY] = childItem.category?.name || '';
        txnRow[IDX_TXN_TAGS] = tags;
        txnRow[IDX_TXN_CLEAR_RECON] = (reconciled ? 'R': (cleared ? 'c': null));
        txnRow[IDX_TXN_MEMO] = memo;
        txnRow[IDX_TXN_MATCHES] = '';
        txnRow[IDX_TXN_STATE] = state;
        // Internal values
        txnRow[IDX_TXN_MINT_ACCOUNT] = mintAccount;
        txnRow[IDX_TXN_ORIG_MERCHANT_INFO] = dataJson.fiData?.description || '';
        txnRow[IDX_TXN_ID] = childItem.id;
        txnRow[IDX_TXN_PARENT_ID] = parentId || '';
        txnRow[IDX_TXN_CAT_ID] = childItem.category?.id || '';
        txnRow[IDX_TXN_TAG_IDS] = tagIds;
        txnRow[IDX_TXN_MOJITO_PROPS] = propsJson;
        txnRow[IDX_TXN_YEAR_MONTH] = txnDate.getFullYear() * 100 + (txnDate.getMonth() + 1);
        // We keep a second copy of txn amount so, if user changes amount,
        // we can compare the before and after value
        txnRow[IDX_TXN_ORIG_AMOUNT] = amount;
        txnRow[IDX_TXN_IMPORT_DATE] = importDate;

        return txnRow;
      } // end local function

      let txnValues = [];
      let invalidCount = 0;
      
      for (let i = 0; i < numRows; ++i)
      {
        const currRow = txnData[i];
        if (currRow.type !== 'CashAndCreditTransaction')
          continue; // skip anything that's not a cash or credit txn

        if (!currRow.date || !currRow.id || isNaN(currRow.amount)) {
          ++invalidCount;
          if (invalidCount < 4) {
            Browser.msgBox('Skipping invalid transaction: ' + JSON.stringify(currRow));
          }
          Debug.log(`Skipping invalid transaction: ${JSON.stringify(currRow)}`);
          continue;
        }
        
        if (currRow.isDuplicate)
        {
          // Skip txns flagged as duplicates
          if (Debug.enabled) Debug.log('Skipping duplicate: %s', JSON.stringify(currRow));
          continue;
        }

        // Does this row contain split data? If so, add a row for each split item.
        if (currRow.splitData?.children) {
          // "splitData" : {
          //   "usePercentages" : false,
          //   "children" : [ {
          //     "id" : "43248151_1282014871_0",
          //     "description" : "Market Basket",
          //     "category" : {
          //       "id" : "43248151_701",
          //       "name" : "Groceries"
          //     },
          //     "amount" : -2.63
          //   }, ...
          Debug.assert(!currRow.splitData.usePercentages, 'Expected currRow.splitData.usePercentages === false');
          for (const childItem of currRow.splitData.children) {
            txnValues.push(createTxnFromData(currRow, childItem, currRow.id, mintAccount, importDate));
          }
        }
        else { // else, it's a single txn
          txnValues.push(createTxnFromData(currRow, null, currRow.parentId, mintAccount, importDate));
        }
      }

      if (invalidCount > 0) {
        Debug.log(`Skipped ${invalidCount} invalid transaction(s)`);
      }

      return txnValues;
    },
    
    // Mint.TxnData
    getUpdatePayload(txnRow, editType, includeTypeAndDate = false)
    {
      let payload = null;

      if (editType === Const.EDITTYPE_DELETE) {
        //FIXME:
        Debug.assert(`Edit type 'DELETE' is not yet implemented`);
      }
      else {
/*
payload: {"description":"Foreside House of Pizza","type":"CashAndCreditTransaction"}
payload: {"category":{"id":"43248151_706"},"type":"CashAndCreditTransaction","notes":"test note","tagData":{"tags":[{"id":"43248151_803867"}]}}
*/
        let memo = txnRow[Const.IDX_TXN_MEMO];
        let propsJson = txnRow[Const.IDX_TXN_MOJITO_PROPS];
        if (propsJson) {
          // If Mojito properties exist, append them to the end of the memo field with a delimeter
          memo = this.appendPropsToString(memo, propsJson);
          if (Debug.enabled) Debug.log("Properties appended to memo: " + propsJson);
        }

        payload = {
          description: txnRow[Const.IDX_TXN_MERCHANT],
          category: { id: `${txnRow[Const.IDX_TXN_CAT_ID]}` },
          notes: memo,
          amount: `${txnRow[Const.IDX_TXN_AMOUNT]}`,
        };

        if (includeTypeAndDate) {
          payload.type = "CashAndCreditTransaction";
          payload.date = txnRow[Const.IDX_TXN_DATE].toISOString().split('T')[0];
        }

        if (editType === Const.EDITTYPE_NEW) {
          let acctInfoMap = Sheets.AccountData.getAccountInfoMap();
          let acctInfo = (!acctInfoMap ? null : acctInfoMap[`${txnRow[Const.IDX_TXN_ACCOUNT]}`]);
          if (!acctInfo) {
            throw new Error(`Account "${txnRow[Const.IDX_TXN_ACCOUNT]}" not found. Unable to determine account id.`);
          }

          payload.accountId = acctInfo.id;
          payload.parentId = null;
          payload.id = null;
          payload.isExpense = (payload.amount <= 0);
          payload.isPending = false;
          payload.isDuplicate = false;
          payload.splitData = null;
          payload.manualTransactionType = 'PENDING';
          payload.checkNumber = null;
          payload.isLinkedToRule = false;
          payload.shouldPullFromAtmWithdrawals = false;
        }

        // Add tags, if any
        let tagIds = txnRow[Const.IDX_TXN_TAG_IDS];
        let tagIdArray = tagIds.split(Const.DELIM);
        let tags = [];
        for (const tagId of tagIdArray) {
          if (!tagId) continue;

          tags.push({ id: tagId });
        }
        if (tags.length > 0) {
          payload.tagData = { tags };
        }
      }

      return payload;
    },

    // Mint.TxnData
    getSplitUpdatePayload(splitRows)
    {
      const { EDITTYPE_EDIT } = Const;
      const splitCount = splitRows.length;
      let payload = null;


      if (splitCount === 1) {
        // If there is only one split txn in this group, then we are effectively deleting the split group
        // and reverting back to a 'normal' transaction.
        payload = this.getUpdatePayload(splitRows[0], EDITTYPE_EDIT);
        payload.splitData = { children: [] };
      }
      else if (splitCount > 1) {
        const totalAmt = splitRows.reduce((sum, nextSplit) => sum + nextSplit[Const.IDX_TXN_AMOUNT], 0);
        payload = {
          type: "CashAndCreditTransaction",
          amount: totalAmt,
          splitData: {
            children: splitRows.map(row => this.getUpdatePayload(row, EDITTYPE_EDIT, false))
          }
        };

        /*
        "splitData": {
          "children": [
            {
            "amount":"-2.63","description":"Market Basket","category":{"id":"43248151_701"}
            },
            {
              "amount":"-40.00","description":"Market Basket","category":{"id":"43248151_2001"}
            }
          ]
        }
        */
      }

      return payload;
    },

    // Mint.TxnData
    appendPropsToString(strValue, propsJson) {
      return Utilities.formatString("%s\n\n\n%s%s", (strValue || ''), Const.DELIM_2, propsJson);
    },

    // Mint.TxnData
    extractPropsFromString(strValue) {
      let extractedParts = {
        text: null,
        props: null,
        propsJson: null,
      };

      let propDelim = strValue.indexOf(Const.DELIM_2);
      if (propDelim >= 0) {
        extractedParts.propsJson = strValue.substr(propDelim + Const.DELIM_2.length);
        if (Debug.enabled) Debug.log("Mojito props: " + extractedParts.propsJson);
        try {
          extractedParts.props = JSON.parse(extractedParts.propsJson);
        } catch (e) {
          if (Debug.enabled) Debug.log(`Unable to parse mojito props. ${e.toString()}`);
          extractedParts.props = null;
        }
        extractedParts.text = strValue.substr(0, propDelim).trim();

      } else {
        extractedParts.text = strValue;
      }

      return extractedParts;
    },
    
  },

  ///////////////////////////////////////////////////////////////////////////////
  // Mint.AccountData

  AccountData: {

    mintTimestampForToday: 0,

    // Mint.AccountData
    downloadAccountInfo()
    {
      let jsonData = Mint.getJsonData(null, 'accounts', 'Account', { offset:0, limit:1000 });
      if (Debug.enabled) Debug.log("%s account(s) found", (!jsonData ? `0`: String(jsonData.length)));
      // Debug.log(`Accounts: ${JSON.stringify(jsonData, null, '  ')}`);
      
      return jsonData;
    },

    // Mint.AccountData
    downloadBalanceHistory(sessionHeaders, account, startDateTs, endDateTs, importTodaysBalance, interactive)
    {
      if (!sessionHeaders)
        return null;

      let balanceHistory = [];

      try
      {
        let payload = {
          offset: 0,
          limit: 50,
          reportView: { type: "ASSETS_TIME" }, // assume asset (but check for debt below)
          dateFilter: { type: "CUSTOM", startDate: "", endDate: "" }, // dates filled before each request
          searchFilters: [
            {
              matchAll: true,
              filters: [{ type: "AccountIdFilter", accountId: account.id }]
            },
            {
              matchAll: false,
              filters:[]
            }
          ]
        }

        // Determine if the reportView should be "DEBTS_TIME" instead of "ASSETS_TIME"
        let isDebtAcct = (account.type === "CreditAccount" || account.type === "LoanAccount");
        if (isDebtAcct) {
          payload.reportView.type = "DEBTS_TIME";
        }

        let daysToFetch = Math.ceil(Math.max(0, (endDateTs - startDateTs)/Const.ONE_DAY_IN_MILLIS));
        let iterStartDate = new Date(startDateTs);

        while (daysToFetch > 0) {
          // If we want to get daily account balances (which we do) then
          // we can only fetch about 40 days at a time
          let nextEndTs = Math.min(iterStartDate.getTime() + (40 * Const.ONE_DAY_IN_MILLIS), endDateTs);
          let iterEndDate = new Date(nextEndTs);

          let strStart = iterStartDate.toISOString().split('T')[0];
          let strEnd = iterEndDate.toISOString().split('T')[0];
          payload.dateFilter.startDate = strStart;
          payload.dateFilter.endDate = strEnd;
          dayCount = Math.floor((iterEndDate.getTime() - iterStartDate.getTime())/Const.ONE_DAY_IN_MILLIS);
          if (Debug.enabled) Debug.log(`Fetching account "${account.name}" (${account.id}) trend data for date range ${strStart} to ${strEnd}, (${dayCount} day(s))`);

          let headers = {...sessionHeaders,
            'content-type': 'application/json',
            Accept: 'application/json',
          };
          let options = {
              method: 'POST',
              headers,
              payload: JSON.stringify(payload),
              // followRedirects: false,
              muteHttpExceptions: true,
          };
    
          // if (Debug.enabled) Debug.log(`downloadBalanceHistory(${account.id}): Starting fetch. payload: ${JSON.stringify(payload)}`);

          let url = 'https://mint.intuit.com/pfm/v1/trends';
          let response = UrlFetchApp.fetch(url, options);

          if (Debug.enabled) Debug.log(`downloadBalanceHistory(${account.id}): Fetch completed. Response code: ${response ? response.getResponseCode(): '<unknown>'}`);

          if (response && (response.getResponseCode() === 200 || response.getResponseCode() === 302))
          {
            //let respHeaders = response.getAllHeaders();
            //Debug.log(respHeaders.toSource());

            let respBody = response.getContentText();
            // if (Debug.enabled) Debug.log(`account trends response: ${respBody}`);
            let respCheck = Mint.checkJsonResponse(respBody);
            if (respCheck.success)
            {
              let results = JSON.parse(respBody);
              // if (Debug.enabled) Debug.log(`Results: ${JSON.stringify(results, null, '  ')}`);

              // Success. We got the account balances.
              // The results.Trend array contains an entry for each day
              // Debug.assert(!!results.Trend, "results.Trend does not exist");
              let data = results.Trend;

              if (!data || data.length === 0) {
                data = [];
                // The Trends array was empty. Should we use today's balance in account.bal?
                if (Debug.enabled) Debug.log("Trend data was empty for account '%s'.", account.name);
                if (Debug.enabled) Debug.log(`Results: ${JSON.stringify(results, null, '  ')}`);
                let now = new Date();
                let todayTs = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());

                if (importTodaysBalance === undefined) {
                  let includeToday = (startDateTs <= todayTs && todayTs <= endDateTs);
                  if (includeToday) {
                    // Add 'balanceHistoryNotAvailable' member to account object
                    account.balanceHistoryNotAvailable = true;
                    importTodaysBalance = true;
                  } else {
                    let msg = `Account balance history does not exist for account '${account.name}'. Would you like import today's balance even though it falls outside your specified date range?`;
                    let choice = Browser.msgBox("Account without history detected", msg, Browser.Buttons.YES_NO);
                    importTodaysBalance = (choice === "yes");
                  }
                  Settings.setInternalSetting(Const.IDX_INT_SETTING_CURR_DAY_ACCT_IMPORT, importTodaysBalance);
                }

                if (importTodaysBalance === true) {
                  // We will "fail gracefully" and use today's balance from the account info.
                  if (Debug.enabled) Debug.log("Substituting 'account.bal' for today's balance: %s", account.name, account.bal);

                  const today = new Date(this.getTodayTimestamp());
                  let dateStr = today.toISOString().split('T')[0];
                  data = [{ dateStr, amount: account.value }];
                }
              } else {
                // Trend data is available. Use the date of the first entry to calculate the time
                // offset from UTC. We may need to use this for accounts that do not have any Trend data.
                this.setTodayTimestamp(new Date(data[0].date).getTime());
              }
              // Only include the fields we actually need
              let balances = data.reduce((accum, curr) => {
                accum.push({
                  dateStr: curr.date,
                  // If it's a debt account, make the amount negative
                  amount: (isDebtAcct ? -curr.amount : curr.amount)
                });
                return accum;
              }, []);
              // Append this data to the end of the balanceHistory array
              balanceHistory.push.apply(balanceHistory, balances);
            }
            else
            {
              account.balanceHistoryNotAvailable = true;
              throw new Error(`Account balance retrieval failed. ` + respBody);
            }
          }
          else
          {
            account.balanceHistoryNotAvailable = true;
            throw new Error(`Account balance retrieval failed. HTTP status code ${response.getResponseCode()}, ${response.getContentText()}`);
          }

          iterStartDate = new Date(iterStartDate.getTime() + (41 * Const.ONE_DAY_IN_MILLIS));
          if (Debug.enabled) Debug.log(`iterStartDate=${iterStartDate.toISOString()}`);
          daysToFetch = Math.max(0, (endDateTs - iterStartDate.getTime())/Const.ONE_DAY_IN_MILLIS);
          if (Debug.enabled) Debug.log(`daysToFetch=${daysToFetch}`);

        } // while

        // Add a 'balanceHistory' member to the account [[ not considered good practice, should return a new object ]]
        account.balanceHistory = balanceHistory;
      }
      catch(e)
      {
        if (Debug.enabled) Debug.log(Debug.getExceptionInfo(e));
        // Browser.msgBox(`Error: ${e.toString()}`);
      }

      return account;
    },

    setTodayTimestamp(mintTimestamp) {
      // mint timestamp could be for any date. We will calculate the mint time offset
      // from UTC, then use that to calculate the timestamp for today.

      let mintDate = new Date(mintTimestamp);
      let utcTs = Date.UTC(mintDate.getFullYear(), mintDate.getMonth(), mintDate.getDate());
      let timeOffset = mintDate.getTime() - utcTs;
      if (Debug.enabled) Debug.log("Mint import timestamp offset from UTC: %d hours", timeOffset/(60*60*1000));

      let today = new Date();
      this.mintTimestampForToday = Date.UTC(today.getFullYear(), today.getMonth(), today.getDate()) + timeOffset;
      if (Debug.enabled) Debug.log("Calculated mint timestamp for today: %s", this.mintTimestampForToday);
    },

    getTodayTimestamp() {
      // If no balance history is available for the specified account, then we use the
      // current balance in account.bal and associate it with today's date (adjusted to match Mint).
      // This function calculates the timestamp for today;
      let timestamp = this.mintTimestampForToday;

      if (timestamp === 0) {
        let today = new Date();
        // Parse timezone info from date string to determine if this it is daylight savings time or not
        let todayStr = today.toString();
        let tzStr = todayStr.match(/\([A-Z]+\)/g); // example: "(MDT)" = Moutain Daylight Time
        let dstChar = (tzStr ? String(tzStr).charAt(2): null); // Parse second char. Should be 'D' or 'S'
        if (Debug.enabled) Debug.log("Determining if it is daylight savings time: %s, '%s'", tzStr, dstChar);
        let isDST = (dstChar ? (dstChar === 'D' ? true: false): false);
        // Determine time offset for Pacific Time (that's what Mint uses)
        let mintTimeOffset = (isDST === true ? 7*60*60*1000 /*7 hours*/: 8*60*60*1000 /*8 hours*/);
        if (Debug.enabled) Debug.log("timezone offset: %d hours", mintTimeOffset / (60*60*1000));
        timestamp = Date.UTC(today.getFullYear(), today.getMonth(), today.getDate()) + mintTimeOffset;
        if (Debug.log) Debug.log("Estimated mint timestamp for today: %s", timestamp);
      }
      return timestamp;
    },

  },

  ///////////////////////////////////////////////////////////////////////////////
  // Mint.Categories class

  Categories: {

    // Mint.Categories
    import(interactive)
    {
      try
      {
        if (interactive) toast("Retrieving categories", "Category update");

        let jsonData = Mint.getJsonData(null, 'categories', 'Category', { offset:0, limit:500 });
        // Debug.log(`Categories: ${JSON.stringify(jsonData, null, '  ')}`);

        // Clear cached category map
        Utils.getPrivateCache().remove(Const.CACHE_CATEGORY_MAP);

        if (interactive) toast(`Retrieved ${jsonData.length} categories`, "Category update");

        this.insertDataIntoSheet(jsonData, interactive);
      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        toast("Import failed. Exception encountered: " + e.toString(), "Error", 20);
      }
    },

    // Mint.Categories
    insertDataIntoSheet(categoryData, interactive)
    {
      if (!categoryData || categoryData.length === 0)
        return;
      
      let numCols = Const.IDX_CAT_LAST_COL + 1;
      Debug.log(`Looping through ${categoryData.length} categories`);
      const {IDX_CAT_NAME, IDX_CAT_ID, IDX_CAT_STANDARD, IDX_CAT_PARENT_ID} = Const;

//      Browser.msgBox(categoryData.toSource());
      let mainCategories = [];
      let subCategories = {};
      let catValues = new Array();

      for (let i = 0; i < categoryData.length; ++i) {
        let currRow = categoryData[i];
        let targetRow = new Array(numCols);

        targetRow[IDX_CAT_NAME] = currRow.name;
        targetRow[IDX_CAT_ID] = currRow.id;
        targetRow[IDX_CAT_STANDARD] = !currRow.isCustom;
        targetRow[IDX_CAT_PARENT_ID] = currRow.parentId || '';

        if (!currRow.parentId) {
          mainCategories.push(targetRow);
        }
        else {
          let subArray = subCategories[currRow.parentId];
          if (!subArray) {
            subArray = subCategories[currRow.parentId] = [];
          }
          subArray.push(targetRow);
        }
      }

      mainCategories.sort((a, b) => {
        let catIdA = parseInt(a[IDX_CAT_ID].split('_')[1]);
        let catIdB = parseInt(b[IDX_CAT_ID].split('_')[1]);
        return (catIdA < catIdB ? -1 : 1);
      });
      mainCategories.forEach(c => {
        catValues.push(c);

        let subArray = subCategories[c[IDX_CAT_ID]];
        if (subArray) {
          subArray.sort((a, b) => { return (a[IDX_CAT_ID] < b[IDX_CAT_ID] ? -1 : 1); });
          catValues.push(...subArray);
        }
      });

      Debug.log(`Inserting ${catValues.length} categories`);

      let range = Utils.getCategoryDataRange();
      range.clear(); // Replace existing categories
      let catRange = range.offset(0, 0, catValues.length, numCols);
      catRange.setValues(catValues);

      if (interactive) toast(`Inserted ${catValues.length} categories`, "Category update");
    },
  },
  
  ///////////////////////////////////////////////////////////////////////////////
  // Mint.Tags class

  Tags: {

    // Mint.Tags
    import(interactive)
    {
      try
      {
        if (interactive) toast("Retrieving tags", "Tag update");

        let jsonData = Mint.getJsonData(null, "tags", 'Tag', { offset:0, limit:500} );

        // Clear cached tag map
        Utils.getPrivateCache().remove(Const.CACHE_TAG_MAP);

        //Browser.msgBox(jsonData.toSource());
        if (interactive) toast("Retrieved " + jsonData.length + " tags", "Tag update");

        this.insertDataIntoSheet(jsonData, interactive);
      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        toast("Import failed. Exception encountered: " + e.toString(), "Error", 20);
      }
    },

    // Mint.Tags
    insertDataIntoSheet(tagData, interactive)
    {
      if (!tagData || tagData.length === 0)
        return;
      
      let numRows = tagData.length;
      let numCols = Const.IDX_TAG_LAST_COL + 1;
      Debug.log("Looping through " + numRows + " tags");
      
      let tagValues = new Array();

      for (let i = 0; i < numRows; ++i) {
        let currRow = tagData[i];
        let targetRow = new Array(numCols);
        
        targetRow[Const.IDX_TAG_NAME] = currRow.name;
        targetRow[Const.IDX_TAG_ID] = currRow.id;
        // Add row to tagValues array
        tagValues.push(targetRow);
       
      } // for i

      Debug.log("Inserting " + tagValues.length + " tags");

      let range = Utils.getTagDataRange();
      range.clear(); // Replace existing tags
      let tagRange = range.offset(0, 0, tagValues.length, numCols);
      tagRange.setValues(tagValues);

      if (interactive) toast("Inserted " + tagValues.length + " tags", "Tag update");
    },

  },

  ///////////////////////////////////////////////////////////////////////////////
  // Mint.Session class

  Session: {

    resetSession() {
      Mint.Session.clearHeaders();
    },

    getHeaders(throwIfNone = true)
    {
      let cache = Utils.getPrivateCache();
      let headers = cache.get(Const.CACHE_SESSION_HEADERS);
      // Debug.log(`session headers: ${headers}`);

      if (headers) {
        // Put the same headers back in the cache to reset the expiration
        cache.put(Const.CACHE_SESSION_HEADERS, headers, Const.CACHE_SESSION_EXPIRE_SEC);
      }
      else
      {
        if (Debug.enabled) Debug.log("Mint.Session.getHeaders: No headers in cache");

        // If cookies are not in the cache, prompt the user to provide mint auth data.
        if (throwIfNone) {
          throw new Error("Mint authentication has expired.");
        }
      }

      return JSON.parse(headers);
    },

    clearHeaders()
    {
      Debug.log("Clearing session headers.");
      let cache = Utils.getPrivateCache();
      cache.remove(Const.CACHE_SESSION_HEADERS);
      cache.remove(Const.CACHE_LOGIN_ACCOUNT);
    },

    getCookies(throwIfNone = true)
    {
      let cache = Utils.getPrivateCache();
      let cookies = cache.get(Const.CACHE_LOGIN_COOKIES);
      let mintAccount = Utils.getMintLoginAccount();

      if (cookies) {
        // Put the same cookies and mint account back in the cache to reset the expiration
        cache.put(Const.CACHE_LOGIN_COOKIES, cookies, Const.CACHE_SESSION_EXPIRE_SEC);
        cache.put(Const.CACHE_LOGIN_ACCOUNT, mintAccount, Const.CACHE_SESSION_EXPIRE_SEC);
      }
      else
      {
        if (Debug.enabled) Debug.log("Mint.Session.getCookies: No login cookies in cache");

        // If cookies are not in the cache, prompt the user to provide mint auth data.
        if (throwIfNone) {
          throw new Error("Mint authentication has expired.");
        }
      }

      return cookies;
    },

    clearCookies()
    {
      Debug.log("Clearing login cookies and token from cache.");
      let cache = Utils.getPrivateCache();
      cache.remove(Const.CACHE_LOGIN_COOKIES);
      cache.remove(Const.CACHE_LOGIN_ACCOUNT);
    },

    getCookiesFromResponse(response) {
      let respHeaders = response.getAllHeaders();
      let setCookieArray = respHeaders["Set-Cookie"];
      if (!setCookieArray) {
        setCookieArray = [];
      }
      // Make sure the setCookieArray is actually an array, and not just a single string
      if (typeof setCookieArray === 'string') {
        setCookieArray = [setCookieArray];
      }
      if (Debug.traceEnabled) Debug.trace("Cookies in response: " + setCookieArray.toSource());

      let cookies = {};
      for (let i = 0; i < setCookieArray.length; ++i)
      {
        let cookie = setCookieArray[i];
        let cookieParts = cookie.split('; ');
        //Debug.log('cookieParts: ' + cookieParts.toSource());
        cookie = cookieParts[0].split('=');
        //Debug.log('cookie: ' + cookie.toSource());
        cookies[ cookie[0] ] = cookie[1];
      }

      //Debug.log('******** cookies: ' + cookies.toSource());      
      return cookies;
    },

    showManualMintAuth() {
      try {
        let htmlOutput = HtmlService.createTemplateFromFile('manual_mint_auth.html').evaluate();
        htmlOutput.setHeight(350).setWidth(600).setSandboxMode(HtmlService.SandboxMode.IFRAME);
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Mint Authentication - HTTP headers');
      }
      catch (e) {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox("Error: " + e.toString());
      }
    },

    /**
     * Called from manual_mint_auth.html
     * @param args
     * @returns {boolean}
     */
    verifyManualAuth(args) {
      let success = false;

      try {
        let cache = Utils.getPrivateCache();
        let { cookie, authorization, intuitTid } = args;
        Debug.log("manual auth args: " + JSON.stringify(args, null, '  '));

        let username = Mint.Session.fetchUserName(args);
        success = !!username;
        // let result = Mint.Session.fetchTokenAndUsername(cookies);
        // let retrievedToken = result && result.token;

        // success = (retrievedToken && token === retrievedToken);
        // Cache cookies and authorization
        cache.put(Const.CACHE_SESSION_HEADERS, JSON.stringify(args), Const.CACHE_SESSION_EXPIRE_SEC);
        cache.put(Const.CACHE_LOGIN_COOKIES, cookie, Const.CACHE_SESSION_EXPIRE_SEC);
        cache.put(Const.CACHE_SESSION_TOKEN, authorization, Const.CACHE_SESSION_EXPIRE_SEC);

        // Save username (email)
        if (username) {
          Settings.setSetting(Const.IDX_SETTING_MINT_LOGIN, username);
          cache.put(Const.CACHE_LOGIN_ACCOUNT, username, Const.CACHE_SESSION_EXPIRE_SEC);
        }

        toast('Mint authentication ' + (success ? 'succeeded': 'FAILED'));
      }
      catch (e) {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox("Error: " + e.toString());
      }
      
      return success;
    },

    /**
     * @obsolete This function no longer works.
     */
    loginMintUser(username, password)
    {
      if (!username) {
        return null;
      }

      password = password || '';

      username = username.toLowerCase();
      let isDemoUser = (username == Const.DEMO_MINT_LOGIN);
      
      let msg = Utilities.formatString("Logging in user %s. %s", username, (isDemoUser ? " Note: The demo user can take a while. Be patient.": ""));
      toast(msg, "Mint Login", 60);
      Debug.log("Logging in user " + username);

      let cookies = null;
      
      try
      {
        let podCookie = Mint.Session.getUserPodCookie(username);
        let response = null;
        
        if (isDemoUser)
        {
          // login to mint demo account
          
          let options = {
            "method": "GET",
            "followRedirects": false
          };
          
          // call the demo login page
          response = UrlFetchApp.fetch("https://mint.intuit.com/demoUser.event", options);
        }
        else // normal login
        {
          let headers = {
            "Cookie": podCookie,
            "Accept": "application/json",
            "X-Request-With": "XMLHttpRequest",
            "X-NewRelic-ID": "UA4OVVFWGwEGV1VaBwc=",
            "Referrer": "https://mint.intuit.com/login.event?task=L&messageId=1&country=US&nextPage=overview.event"
          };

          let formData = {
            "username": username,
            "password": password,
            "task": "L",
            "timezone": Utils.getTimezoneOffset(),
            "browser": "Chrome",
            "browserVersion": 39,
            "os": "win"
          };
          
          let options = {
            "method": "POST",
            "headers": headers,
            "payload": formData,
            "followRedirects": false
            //    "muteHttpExceptions": true
          };
          
          // call the login page
          response = UrlFetchApp.fetch("https://mint.intuit.com/loginUserSubmit.xevent", options);
        }

        Debug.log("Response code: " + response.getResponseCode());
        if (response && (response.getResponseCode() == 200 || response.getResponseCode() == 302))
        {
          let respBody = response.getContentText();
          Debug.log("Response: " + respBody);
          let respJson = (respBody ? JSON.parse(respBody): {});
          
          if (respJson.action) {// && respJson.action === 'CHALLENGE') {
            Debug.log('login action: ' + respJson.action);
          }
    
          // get the cookies (including auth info) so we can use them in subsequent json requests
          let respHeaders = response.getAllHeaders();
          if (Debug.traceEnabled) Debug.trace("loginUser: all headers: " + respHeaders.toSource());
          
          let setCookieArray = respHeaders["Set-Cookie"];
          if (!setCookieArray) {
            setCookieArray = [];
          }
          if (Debug.traceEnabled) Debug.trace("loginUser: cookies: " + setCookieArray.toSource());
          let success = (isDemoUser ? true: false);

          // Save all of the cookies returned in the login response
          cookies = "";
          for (let i = 0; i < setCookieArray.length; i++)
          {
            let thisCookie = setCookieArray[i];
            
            // Make sure the login was successful by looking for a specific cookie
            if (!success && thisCookie.indexOf(username) > 0)
            {
              success = true;
            }
            
            cookies += thisCookie + "; ";
          }

          // Add the pod cookie to the login response cookies so we have the full set
          cookies += podCookie;
          
          let cache = Utils.getPrivateCache();
          
          if (success)
          {
            let token = respJson.sUser.token || respJson.CSRFToken;
            if (!token) {
              if (Debug.enabled) Debug.log("Token was not returned in login response. Response text: " + respBody);
            }

            // Login succeeded. Save the cookies and mint account in the cache.
            cache.put(Const.CACHE_LOGIN_COOKIES, cookies, Const.CACHE_SESSION_EXPIRE_SEC);
            cache.put(Const.CACHE_LOGIN_ACCOUNT, username.toLowerCase(), Const.CACHE_SESSION_EXPIRE_SEC);

            Debug.log("Login succeeded");
            toast("Login succeeded");

            if (token) {
              if (Debug.enabled) Debug.log("Saving token in cache: " + token);
              cache.put(Const.CACHE_SESSION_TOKEN, token, Const.CACHE_SESSION_EXPIRE_SEC);
            }
            else {
              // Remove the session token, if any, because this login just changed it.
              if (Debug.enabled) Debug.log("No token was return in login response. Removing existing token (if any) from cache.");
              cache.remove(Const.CACHE_SESSION_TOKEN);
            }

            if (Debug.traceEnabled) Debug.trace("Cookies: %s", cookies);
          }
          else
          {
            // The login wasn't successful. Clear the cookies.
            cache.remove(Const.CACHE_LOGIN_COOKIES);
            cache.remove(Const.CACHE_LOGIN_ACCOUNT);
            cookies = null;
            toast("Login failed.");
            if (Debug.enabled) Debug.log("Login failed: Response: %s", response.getContentText());
          }
        }
        else
        {
          let msg = "Login failed. HTTP status code: " + response.getResponseCode();
          toast(msg);
          if (Debug.enabled) Debug.log(msg);
        }
        
      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        toast("Login failed.");
        Browser.msgBox("Error: " + e.toString());
      }
      
      return cookies;
    },

    getIusSession() {
      let headers = {
      };
      let options = {
          "method": "GET",
          "followRedirects": false
      };

      Debug.log("getIusSession: Starting request");
      let cookies = "";

      let response = UrlFetchApp.fetch("https://accounts.intuit.com/xdr.html", options);
      if (response && response.getResponseCode() === 200)
      {
        cookies = Mint.Session.getCookiesFromResponse(response);
        Debug.log('Cookies from iussession request: ' + cookies.toSource());
      }
      else
      {
        throw new Error("getIusSession failed. " + response.getResponseCode());
      }

      return cookies['ius_session'];
    },
    
    clientSignIn(username, password, iusSession) {

      let headers = {
        'Content-Type': 'application/json',
        'Accept': 'application/json; charset=utf-8',
        'Cookie': 'ius_session=' + iusSession + ';',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.84 Safari/537.36'
      };
      let options = {
        "method": "POST",
        "headers": headers,
        "payload": {'username': username, 'password': password },
        "followRedirects": false,
        "muteHttpExceptions": true
      };

      Debug.log("clientSignIn: Starting request");
      let cookies = "";

      let response = null;

      // From Google Apps Script, the below http requests always fails with a 503...      
      response = UrlFetchApp.fetch("https://accounts.intuit.com/access_client/sign_in", options);
      if (response)
      {
        Debug.log("Response code: " + response.getResponseCode());
        let respBody = response.getContentText();
        Debug.log("Response:  " + respBody);
        if (response.getResponseCode() !== 200) {
          throw new Error("clientSignIn POST failed. code: " + response.getResponseCode());
        }
      }
      else
      {
        throw new Error("clientSignIn POST failed. " + respBody);
      }

      return respBody;
    },
    
    getUserPodCookie(mintAccount)
    {
      let headers = {
          "Cookie": "mintUserName=\"" + mintAccount + "\"; "
      };
      let options = {
          "method": "POST",
          "headers": headers,
          "payload": { "username": mintAccount },
          "followRedirects": false
      };

      Debug.log("getUserPodCookie: Starting request");
      let podCookie = "";

      try
      {
        let response = null;
        
        response = UrlFetchApp.fetch("https://mint.intuit.com/getUserPod.xevent", options);
        if (response && (response.getResponseCode() === 200))
        {
          let respBody = response.getContentText();
          //Debug.log("Response:  " + respBody);
          
          let respCheck = Mint.checkJsonResponse(respBody);
          if (respCheck.success)
          {
            let results = JSON.parse(respBody);
            //Debug.log("Results: " + results.toSource());
            
            Debug.assert(results.mintPN, "results.mintPN != undefined");

            let respHeaders = response.getAllHeaders();
            let setCookieArray = respHeaders["Set-Cookie"] || [];
            if (Debug.traceEnabled) Debug.trace("getUserPodCookie: cookies: " + setCookieArray.toSource());

            for (let i = 0; i < setCookieArray.length; i++)
            {
              let thisCookie = setCookieArray[i];
              if (thisCookie.indexOf("mintPN") == 0) {
                podCookie += thisCookie + "; ";
                break;
              }
            }
            podCookie += "mintUserName=\"" + mintAccount + "\"; ";
            Debug.log("podCookie: %s", podCookie);
          }
          else
          {
            throw new Error("getUserPodCookie failed. " + respBody);
          }
        }
      }
      catch(e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox("Error: " + e.toString());
      }

      return podCookie;
    },

    // Mint.Session
    getSessionToken(cookies, force)
    {
      if (force == undefined)
        force = false;

      let token = '';
      const cache = Utils.getPrivateCache();

      if (!force)
      {
        // Is the token still in the cache?
        Debug.log("Looking up token in cache");
        token = cache.get(Const.CACHE_SESSION_TOKEN);
        if (token)
        {
          Debug.log("Token found in cache: " + token);
          return token;
        }
      }

      if (!cookies)
      {
        cookies = Mint.Session.getCookies();

        if (cookies) {
          token = cache.get(Const.CACHE_SESSION_TOKEN);
          if (token) {
            Debug.log("Token found in cache after re-login: " + token);
            return token;
          }
        }
      }

      Debug.log("getSessionToken: Fetching token from Mint overview page...");
      const result = Mint.Session.fetchTokenAndUsername(cookies);

      return (result ? result.token : null);
    },

    fetchUserName(headers) {
      let userName = '';
      const cache = Utils.getPrivateCache();

      let options = {
        method: "GET",
        headers
      };
      let url = "https://mint.intuit.com/pfm/v1/user";

      try
      {
        let response = UrlFetchApp.fetch(url, options);
        if (response && (response.getResponseCode() === 200))
        {
          let respBody = response.getContentText();
          Debug.log("/v1/user fetch response:  " + respBody);

          let jsonResp = JSON.parse(respBody);
          userName = jsonResp.emailAddress || '';
        }
        else
        {
          toast("Data retrieval failed. HTTP status code: " + response.getResponseCode(), "Error");
          // Failure of this request is likely due to a problem with the login cookies.
          // We'll reset the session so the user will be forced to login again if another attempt is made.
          Mint.Session.resetSession();
        }
      }
      catch(e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox(e.toString());
      }

      return userName;
    },

    fetchTokenAndUsername(cookies) {
      let result = {};
      const cache = Utils.getPrivateCache();

      let headers = {
        "Cookie": cookies
      };
      let options = {
        "method": "GET",
        "headers": headers
      };
      let url = "https://mint.intuit.com/overview";

      try
      {
        let response = UrlFetchApp.fetch(url, options);
        if (response && (response.getResponseCode() === 200))
        {
          let respBody = response.getContentText();
          Debug.log("Overview page fetch response:  " + respBody);

          result = this.parseTokenAndUsernameFromOverviewHtml(respBody);
          if (!result) {
            throw new Error('Invalid headers provided.')
          }
          if (result.token) {
            cache.put(Const.CACHE_SESSION_TOKEN, result.token, Const.CACHE_SESSION_EXPIRE_SEC);
          }
          else {
            Debug.log("Unable to parse token from overview.event html");
          }

          if (result.username) {
            cache.put(Const.CACHE_LOGIN_ACCOUNT, result.username, Const.CACHE_SESSION_EXPIRE_SEC);
          }
          else {
            Debug.log("Unable to parse username from overview.event html");
          }
        }
        else
        {
          toast("Data retrieval failed. HTTP status code: " + response.getResponseCode(), "Error");
          // Failure to get session token is likely due to a problem with the login cookies.
          // We'll reset the session so the user will be forced to login again if another attempt is made.
          Mint.Session.resetSession();
        }
      }
      catch(e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox(e.toString());
      }

      return result;
    },

    parseTokenAndUsernameFromOverviewHtml(html)
    {
      Debug.log('Parsing overview.html for token');
      let parse1 = html.match(/<input\s+type="hidden"\s+id="javascript-user"\s[^>]*value="[^"]*"\s*\/>/gi);
      if (!parse1) {
        Debug.log('Unable to find hidden input with id "javascript-user"');
        return null;
      }
      Logger.log(parse1[0]);

      let parse2 = parse1[0].match(/\{.*\}/gi);
      if (!parse2) {
        Debug.log('Unable to find "javascript-user" JSON value');
        return null;
      }
      Logger.log(parse2[0]);

      let json = parse2[0].replace(/&quot;/g, "\"");
      Debug.log("token json: %s", json);
      let userValues = JSON.parse(json);
      if (!userValues) {
        Debug.log('Unable to parse "javascript-user" value as JSON');
        return null;
      }

      if (!userValues.token) {
        Debug.log('Unable to find "token" property in "javascript-user" JSON');
        return null;
      }

      let token = userValues.token;
      if (Debug.enabled) Debug.log('Retrieved token: %s', token);

      return userValues;
    },
  },

  ///////////////////////////////////////////////////////////////////////////////
  // Helpers

  // Mint
  waitForMintFiRefresh(interactive)
  {
    return true; // not implemented
    if (interactive == undefined)
      interactive = false;

    let isMintDataReady = false;
    
    try
    {
      let maxWaitSeconds = Settings.getSetting(Const.IDX_SETTING_MINT_FI_SYNC_TIMEOUT);
      if (Debug.enabled) Debug.log("Waiting a maxiumum of %s seconds for Mint to refresh data from financial institutions.", maxWaitSeconds);
      let elapsedSeconds = 0;
      while (elapsedSeconds < maxWaitSeconds) {
        if (!this.isMintRefreshingData()) {
          Debug.log("Mint is done refreshing account data from financial institutions");
          isMintDataReady = true;
          break;
        }

        // Sleep 10 seconds
        let waitSec = 10;
        if (interactive) {
          let currentWait = (elapsedSeconds > 0 ? Utils.getHumanFriendlyElapsedTime(elapsedSeconds): "");
          toast(Utilities.formatString("Waiting for Mint to sync data from your financial institutions %s", currentWait), "Mint data update", 10);
        }
        
        Utilities.sleep(waitSec * 1000);
        elapsedSeconds += waitSec;
      }

      if (!isMintDataReady) {
        if (Debug.enabled) Debug.log("Mint took too long to refresh FI data.");
        if (interactive) toast("Mint is taking too long to refresh your account data. Try logging in to the mint.com website to make sure your account information is up to date.", "", 10);
      }

    }
    catch (e)
    {
      if (interactive) Browser.msgBox("Unable to determine if Mint is ready. Error: " + e.toString());
    }

    return isMintDataReady;
  },
  
  // Mint
  isMintRefreshingData()
  {
return false;
    // This function calls Mint's userStatus API to see if Mint is currently sync'ing account data
    // from the financial institutions.
    
    let isRefreshing = true;
    
    let cookies = Mint.Session.getCookies();
    let headers = {
      "Cookie": cookies,
      "Accept": "application/json"
    };
    let options = {
      "method": "GET",
      "headers": headers,
    };
    
    let url = "https://mint.intuit.com/userStatus.xevent";
    let queryParams = Utilities.formatString("?rnd=%d", Date.now());
    
    //Debug.log("userStatus.xevent: Starting fetch");
    //Debug.log("  -- using cookies: %s", cookies);

    try
    {
      let response = UrlFetchApp.fetch(url + queryParams, options);
      if (response && (response.getResponseCode() === 200))
      {
        //Debug.log(response.getResponseCode());
        
        let respBody = response.getContentText();
        let respCheck = Mint.checkJsonResponse(respBody);
        if (respCheck.success)
        {
          let results = JSON.parse(respBody);
          if (Debug.enabled) Debug.log("isMintRefreshingData: Results: " + results.toSource());

          isRefreshing = results.isRefreshing;
        }
        else
        {
          if (Debug.enabled) Debug.log("isMintRefreshingData: Retrieval failed: " + respBody);
          throw respBody;
        }
      }
      else
      {
        Debug.log("isMintRefreshingData: Retrieval failed. HTTP status code: " + response.getResponseCode());
      }
    }
    catch(e)
    {
      Debug.log(Debug.getExceptionInfo(e));
      throw e;
    }
    
    return isRefreshing;
  },

  // Mint
  getJsonData(sessionHeaders, apiName, rootTag, queryParamMap = null)
  {
    if (!sessionHeaders) {
      sessionHeaders = Mint.Session.getHeaders();
    }

    let headers = {...sessionHeaders,
      Accept: "application/json"
    };
    let options = {
      method: "GET",
      headers,
    };

    let url = `https://mint.intuit.com/pfm/v1/${apiName}`;
    let queryParams = '';
    if (queryParamMap) {
      for (let param in queryParamMap) {
        queryParams += `${param}=${queryParamMap[param]}&`;
      }
      // Prepend '?' to front and strip off last '&'
      queryParams = '?' + queryParams.substring(0, queryParams.length - 1);
    }

    let response = null;

    if (Debug.enabled) Debug.log(`getJsonData(): Starting fetch ${url + queryParams}`);
    
    let jsonData = null;
    
    try
    {
      response = UrlFetchApp.fetch(url + queryParams, options);
      if (Debug.enabled) Debug.log(Utilities.formatString("getJsonData(%s): Fetch completed. Response code: %d", apiName, (response ? response.getResponseCode(): "<unknown>")));

      if (response && (response.getResponseCode() === 200))
      {
        //let respHeaders = response.getAllHeaders();
        //Debug.log(respHeaders.toSource());

        let respBody = response.getContentText();
        let respCheck = Mint.checkJsonResponse(respBody);
        if (respCheck.success)
        {
          let results = JSON.parse(respBody);
          //Debug.log(`Results: ${JSON.stringify(results, null, '  '}`);

          Debug.assert(!!results[rootTag], `results.${rootTag} exists`);
          jsonData = results[rootTag];
        }
        else
        {
          throw new Error("Data retrieval failed. " + respBody);
        }
      }
      else
      {
        toast("Data retrieval failed. HTTP status code: " + response.getResponseCode(), "Download error");
      }
    }
    catch(e)
    {
      if (Debug.enabled) Debug.log(Debug.getExceptionInfo(e));
      Browser.msgBox("Error: " + e.toString());
    }
    
    return jsonData;
  },

  // Mint
  checkJsonResponse(json)
  {
    let result = {
      success: (!!json && (json.indexOf('<error>') < 0)),
      sessionExpired: false
    };

    if (!result.success) {
      Debug.log("Request failed");

      if (json.indexOf("Session has expired") > 0)
      {
        result.sessionExpired = true;
        Debug.log("Session expired");
        Mint.Session.resetSession();
      }
    }
    
    return result;
  },

  // Mint
  convertMintDateToDate(mintDate, today)
  {
    let convertedDate = null;

    if (mintDate.indexOf("/") > 0) {
      // mintDate is formatted as month/day/year
      let dateParts = mintDate.split("/");

      Debug.assert(dateParts.length >= 3, "convertMintDateToDate: Invalid date, less than 3 date parts");

      let year = parseInt(dateParts[2], 10);
      let month = parseInt(dateParts[0], 10) - 1;
      let day = parseInt(dateParts[1], 10);

      if (year < 100) {
        let century = Math.round(today.getFullYear() / 100) * 100;
        year = century + year;
      }

      convertedDate = new Date(year, month, day);

    } else if (mintDate.indexOf(" ") > 0) {
      // mintDate is formatted as "Month Day"
      let dateParts = mintDate.split(" ");
      let year = today.getFullYear();
      let month = Const.MONTH_LOOKUP_1[dateParts[0]];
      let day = parseInt(dateParts[1]);

      if (today.getMonth() < month)
        --year;

      Debug.assert(day > 0, "convertMintDateToDate: day is zero");

      convertedDate = new Date(year, month, day);
    }
    else
    {
      convertedDate = new Date(2000, 0, 1);
    }

    return convertedDate;
  },

  // Mint
  getMintAccounts(fiAccount) {
    let txnRange = Utils.getTxnDataRange(true);
    if (!txnRange) {
      return null;
    }

    let mintAccountMap = [];
  
    let mintAcctRange = txnRange.offset(0, Const.IDX_TXN_MINT_ACCOUNT, txnRange.getNumRows(), 1);
    let mintAcctValues = mintAcctRange.getValues();
    let mintAcctValuesLen = mintAcctValues.length;

    let fiAcctValues = null;
    if (fiAccount) {
      let fiAcctRange = txnRange.offset(0, Const.IDX_TXN_ACCOUNT, txnRange.getNumRows(), 1);
      fiAcctValues = fiAcctRange.getValues();
    }

    for (let i = 0; i < mintAcctValuesLen; ++i) {
      if (fiAcctValues && fiAccount !== fiAcctValues[i][0]) {
        continue; // Financial institution account doesn't match specified account. Skip it.
      }
      let mintAcct = mintAcctValues[i][0];
      if (mintAcct && !mintAccountMap[mintAcct]) {
        mintAccountMap[mintAcct] = true;
        if (Debug.enabled) Debug.log("Found mint account: " + mintAcct);
      }
    }

    let mintAccounts = [];
    for (let mintAcct in mintAccountMap) {
      mintAccounts.push(mintAcct);
    }

    if (Debug.enabled) Debug.log("mintAccounts array: " + mintAccounts);
    return mintAccounts;
  },

  _categoryMap: null,  // This variable is only valid while server-side code is executing. It resets each time.
  
  // Mint
  getCategoryMap() {
    if (!this._categoryMap) {
      let cache = Utils.getPrivateCache();
      let catMap = {};
      //cache.remove(CACHE_CATEGORY_MAP);
      let catMapJson = cache.get(Const.CACHE_CATEGORY_MAP);
      if (catMapJson && catMapJson !== "{}")
      {
        Debug.log("Category map found in cache. Parsing JSON.");
        catMap = JSON.parse(catMapJson);
      }
      else
      {
        Debug.log("Rebuilding category map");
        let range = Utils.getCategoryDataRange();
        let catRange = range.offset(0, 0, range.getNumRows(), Const.IDX_CAT_ID + 1);
        let catValues = catRange.getValues();
        let catCount = catValues.length;
        for (let i = 0; i < catCount; ++i) {
          let catName = catValues[i][Const.IDX_CAT_NAME];
          catMap[ catName.toLowerCase() ] = { catId: catValues[i][Const.IDX_CAT_ID], displayName: catName };
        }

        Debug.log("Saving category map in cache");
        catMapJson = JSON.stringify(catMap);
        cache.put(Const.CACHE_CATEGORY_MAP, catMapJson, Const.CACHE_MAP_EXPIRE_SEC);
      }

      this._categoryMap = catMap;
    }

    return this._categoryMap;
  },

  // Mint
  validateCategory(category, interactive)
  {
    if (interactive == undefined)
      interactive = false;

    let validationInfo = { isValid: false, displayName: "", catId: 0 };
    let catInfo = this.lookupCategoryId(category);
    if (catInfo) {
      validationInfo.isValid = true;
      validationInfo.displayName = catInfo.displayName;
      validationInfo.catId = catInfo.catId;
    }

    if (!validationInfo.isValid && interactive) toast(Utilities.formatString("The category \"%s\" is not valid. Please \"undo\" your change.", category));
    return validationInfo;
  },

  // Mint
  lookupCategoryId(category)
  {
    if (!category)
      return null;

    let catInfo = null;

    let catMap = this.getCategoryMap();

    let catLower = category.toLowerCase();
    let lookupVal = catMap[catLower];
    if (lookupVal) {
      catInfo = lookupVal;
  //    if (Debug.enabled) Debug.log(Utilities.formatString("lookupCategoryId: Found catId %d for category \"%s\"", catInfo.catId, category));
    }
    else {
      if (Debug.enabled) Debug.log(Utilities.formatString("lookupCategoryId: Category \"%s\" not found", category));
    }

    return catInfo;
  },

  _tagMap: null,  // This variable is only valid while server-side code is executing. It resets each time.
  
  // Mint
  getTagMap() {
    if (!this._tagMap) {
      let tagMap = {};

      let cache = Utils.getPrivateCache();
      let tagMapJson = cache.get(Const.CACHE_TAG_MAP);
      if (tagMapJson && tagMapJson !== "{}")
      {
//        Debug.log("Tag map found in cache. Parsing JSON.");
        tagMap = JSON.parse(tagMapJson);
      }
      else
      {
        Debug.log("Rebuilding tag map");
        let range = Utils.getTagDataRange();
        let tagValues = range.getValues();
        let tagCount = tagValues.length;
        for (let i = 0; i < tagCount; ++i) {
          let tagName = tagValues[i][Const.IDX_TAG_NAME];
          tagMap[ tagName.toLowerCase() ] = { displayName: tagName, tagId: tagValues[i][Const.IDX_TAG_ID] };
        }

        Debug.log("Saving tag map in cache");
        tagMapJson = JSON.stringify(tagMap);
//        Debug.log(tagMapJson);
        cache.put(Const.CACHE_TAG_MAP, tagMapJson, Const.CACHE_MAP_EXPIRE_SEC);
        //Debug.log(tagMap.toSource());
      }

      this._tagMap = tagMap;
    }
    
    return this._tagMap;
  },

  composeTxnTagArray(tagsVal, clearReconVal) {
    let tagArray = [];

    const tagCleared = Mint.getClearedTag();
    const tagReconciled = Mint.getReconciledTag();

    const reconciled = (clearReconVal.toUpperCase() === "R");
    const cleared = (clearReconVal !== null && clearReconVal !== ""); // "Cleared" is anything other than "R" or empty
    if (reconciled) {
      // Reconciled transactions are also 'cleared', so we'll include both tags
      tagArray.push(tagReconciled);
      tagArray.push(tagCleared);
    }
    else if (cleared) {
      tagArray.push(tagCleared);
    }

    if (tagsVal) {
      tagArray.push.apply(tagArray, tagsVal.split(Const.DELIM));
    }

    return tagArray;
  },

  // Mint
  /**
   * Validate the tags of a transaction row, separating out 'cleared' and 'reconcied' status.
   * @param tags {string[]}
   * @param [interactive] {boolean}
   * @returns {{isValid: boolean, tagNames: string, tagIds: string, cleared: boolean, reconciled: boolean}}
   */
  validateTxnTags(tagArray, interactive)
  {
    let validationInfo = {
      isValid: true,
      tagNames: '',
      tagIds: '',
      cleared: false,
      reconciled: false
    };

    if (!tagArray) {
      return validationInfo;
    }

    if (interactive == undefined)
      interactive = false;

    let tagNames = '';
    let tagIds = '';
    let tagCleared = Mint.getClearedTag();
    let tagReconciled = Mint.getReconciledTag();

    let tagMap = this.getTagMap();
    for (let i = 0; i < tagArray.length; ++i) {
      let tag = tagArray[i].trim();
      if (tag === '')
        continue;

      let lookupVal = tagMap[tag.toLowerCase()];
      if (lookupVal) {
//        Debug.log(lookupVal.toSource());
//        if (Debug.enabled) Debug.log(Utilities.formatString("Found tagId %d for tag \"%s\"", lookupVal.tagId, tag));

        // Build list of tags (using exact case from Mint).
        // Cleared and reconciled tags are handled separately.
        if (lookupVal.displayName === tagCleared) {
          validationInfo.cleared = true;
        }
        else if (lookupVal.displayName === tagReconciled) {
          validationInfo.reconciled = true;
        }
        else {
          tagNames += lookupVal.displayName + Const.DELIM;
        }

        // Build list of tag IDs
        tagIds += lookupVal.tagId + Const.DELIM;
      }
      else {
        if (interactive) { toast(Utilities.formatString("Tag \"%s\" is not valid. You must add this tag using the Mint website or mobile app.", tag)); }
        if (Debug.enabled) Debug.log(Utilities.formatString("No tagId found for tag \"%s\"", tag));
        validationInfo.isValid = false;
        break;
      }
    }

    if (validationInfo.isValid) {
      validationInfo.tagNames = tagNames;
      validationInfo.tagIds = tagIds;
    }

    return validationInfo;
  },

  _clearedTag: undefined, // This variable is only valid while server-side code is executing. It resets each time.

  // Mint
  getClearedTag() {
    if (this._clearedTag === undefined) {
      let tag = Utils.getPrivateCache().get(Const.CACHE_SETTING_CLEARED_TAG);
      if (tag === null) {
        tag = Settings.getSetting(Const.IDX_SETTING_CLEARED_TAG);
        // If setting is empty, store empty string in cache
        Utils.getPrivateCache().put(Const.CACHE_SETTING_CLEARED_TAG, tag || '', 300); // Save the tag in the cache for 5 minutes
  
        if (Debug.enabled) Debug.log("getClearedTag: %s", tag);
      }

      this._clearedTag = tag || null;
    }

    return this._clearedTag;
  },

  _reconciledTag: undefined, // This variable is only valid while server-side code is executing. It resets each time.

  // Mint
  getReconciledTag() {
    if (this._reconciledTag === undefined) {
      let tag = Utils.getPrivateCache().get(Const.CACHE_SETTING_RECONCILED_TAG);
      if (tag === null) {
        tag = Settings.getSetting(Const.IDX_SETTING_RECONCILED_TAG);
        // If setting is empty, store empty string in cache
        Utils.getPrivateCache().put(Const.CACHE_SETTING_RECONCILED_TAG, tag || '', 300); // Save the tag in the cache for 5 minutes
  
        if (Debug.enabled) Debug.log("getReconciledTag: %s", tag);
      }

      this._reconciledTag = tag || null;
    }
    
    return this._reconciledTag;
  }
};
