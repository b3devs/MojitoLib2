'use strict';
/*
 * Copyright (c) 2013-2019 b3devs@gmail.com
 * MIT License: https://spdx.org/licenses/MIT.html
 */

import {Const} from './Constants.js';
import {Debug} from './Debug.js';
import {Utils, Settings, toast} from './Utils.js';


export const Upgrade = {

  autoUpgradeMojitoIfApplicable : function() {
    // Get spreadsheet version from Settings sheet
    const ssVer = String(Settings.getInternalSetting(Const.IDX_INT_SETTING_MOJITO_VERSION));
    // Get MojitoLib version from constants
    const libVer = String(Const.CURRENT_MOJITO_VERSION);

    if (ssVer === libVer) {
      return libVer;
    }

    const newVer = Upgrade.upgradeMojito(ssVer, libVer);
    return newVer;
  },

  upgradeMojito : function(fromVer, toVer) {

    if (Debug.enabled) Debug.log("Attempting to upgrade Mojito from version %s to %s", fromVer, toVer);
    toast(`Upgrading Mojito from version ${fromVer} to version ${toVer}`, 'Mojito upgrade');

    if (this.compareVersions(fromVer, toVer) > 0) {
      // Spreadsheet version is greater than MojitoLib ver. We don't support downgrading.
      Debug.log("upgradeMojito - fromVer (%s) is greater than toVer (%s). Aborting.", fromVer, toVer);
      Browser.msgBox("Internal version mismatch", "Your Mojito spreadsheet version (Settings sheet) is greater than the MojitoLib version it is using. This kind of version mismatch is not supported. Please download a new copy of Mojito to fix this problem.", Browser.Buttons.OK);
      return toVer; // Return the lesser of the two versions
    }

    let newVer = fromVer;
    let upgradeFailed = false;

    try
    {
  
      //TODO: Finish implementing this function
      while (newVer !== toVer && !upgradeFailed) {

        switch (newVer)
        {
        case "1.2.1":
          newVer = this.upgradeFrom_1_2_1(newVer);
          break;

        default:
          throw new Error(`Auto-upgrade from version ${fromVer} to ${toVer} is not supported. Please download the latest version of Mojito instead.`);
        }

      } // while

      if (Debug.enabled) Debug.log("Upgrade to version %s complete", toVer);
    }
    catch(e)
    {
      upgradeFailed = true;
      Debug.log(Debug.getExceptionInfo(e));
      //TODO: Need better message to user
      Browser.msgBox("Mojito upgrade failed", `Mojito upgrade from version ${fromVer} to ${toVer} failed. Error: ${e.toString()}`, Browser.Buttons.OK);
    }

    if (newVer !== fromVer) {
      // Update version in spreadsheet
      Settings.setInternalSetting(Const.IDX_INT_SETTING_MOJITO_VERSION, newVer);
    }

    toast(upgradeFailed ? "Failed." : "Done.", "Mojito upgrade");

    return newVer;
  },

  upgradeFrom_1_2_1: function(fromVer) {
    // Just a code upgrade. Nothing to do.
    return "2.0.0";
  },

  //--------------------------------------------------------------------------
  compareVersions : function(ver1, ver2) {
    if (ver1 === ver2) {
      return 0;
    }

    if (ver1 == null) {
      return -1;
    }
    if (ver2 == null) {
      return 1;
    }

    // Version format should be 1.2.3.4
    var ver1Parts = ver1.split(".");
    var ver2Parts = ver2.split(".");
    var ver1Size = ver1Parts.length;
    var ver2Size = ver2Parts.length;
    for (var i = 0; i < ver1Size && i < ver2Size; ++i) {
      var v1 = parseInt(ver1Parts[i]);
      var v2 = parseInt(ver2Parts[i]);
      if (isNaN(v1) || isNaN(v2))
        break;

      if (v1 === v2)
        continue;

      if (v1 < v2)
        return -1;
      if (v1 > v2)
        return 1;
    }

    // If we finished the loop, then one version must have fewer parts than the other
    if (ver1Size < ver2Size)
      return -1;

    return 1;
  },
};
