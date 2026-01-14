/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { dialogFallback } from "./fallbackauthdialog";
import { callGetUserData } from "./middle-tier-calls";
import { showMessage } from "./message-helper";
import { handleClientSideErrors } from "./error-handler";

/* global OfficeRuntime, Office */

let retryGetMiddletierToken = 0;

async function getSsoToken(options?: any): Promise<string> {
  // Try modern OfficeRuntime API first
  let lastError: any = null;
  if (typeof OfficeRuntime !== 'undefined' && OfficeRuntime.auth && OfficeRuntime.auth.getAccessToken) {
    try {
      return await OfficeRuntime.auth.getAccessToken(options || {});
    } catch (err) {
      // fall through to try older API
      console.warn('OfficeRuntime.auth.getAccessToken failed, trying Office.context.auth...', err);
      lastError = err;
    }
  }

  // Fallback to older callback-based Office.context.auth.getAccessToken
  if (typeof Office !== 'undefined' && Office.context && Office.context.auth && typeof Office.context.auth.getAccessToken === 'function') {
    return new Promise((resolve, reject) => {
      try {
        Office.context.auth.getAccessToken(options || {}, (result: any) => {
          // result may be an object with .value or .value is token
          if (result && (result.value || result.Value)) {
            resolve(result.value || result.Value);
          } else if (result && result.status && result.status.toLowerCase() === 'succeeded' && result.value) {
            resolve(result.value);
          } else {
            reject(result && result.error ? result.error : new Error('Failed to obtain SSO token'));
          }
        });
      } catch (e) {
        reject(e);
      }
    });
  }

  // If no SSO method is available, throw a standardized error with code 13012 so fallback logic can handle it
  if (lastError && (lastError.code || lastError.name)) {
    throw lastError;
  }
  const notSupported: any = new Error('API is not supported in this platform.');
  (notSupported as any).code = 13012;
  (notSupported as any).name = 'API Not Supported';
  throw notSupported;
}

export async function getUserData(callback): Promise<void> {
  try {
    let middletierToken: string = await getSsoToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: false,
    });
    let response: any = await callGetUserData(middletierToken);
    if (!response) {
      throw new Error("Middle tier didn't respond");
    } else if (response.claims) {
      // Microsoft Graph requires an additional form of authentication. Have the Office host
      // get a new token using the Claims string, which tells AAD to prompt the user for all
      // required forms of authentication.
      let mfaMiddletierToken: string = await getSsoToken({
        authChallenge: response.claims,
      });
      response = callGetUserData(mfaMiddletierToken);
    }

    // AAD errors are returned to the client with HTTP code 200, so they do not trigger
    // the catch block below.
    if (response.error) {
      handleAADErrors(response, callback);
    } else {
      callback(response);
    }
  } catch (exception) {
    // if handleClientSideErrors returns true then we will try to authenticate via the fallback
    // dialog rather than simply throw and error
    if (exception.code) {
      if (handleClientSideErrors(exception)) {
        dialogFallback(callback);
      }
    } else {
      showMessage("EXCEPTION: " + JSON.stringify(exception));
      throw exception;
    }
  }
}

function handleAADErrors(response: any, callback: any): void {
  // On rare occasions the middle tier token is unexpired when Office validates it,
  // but expires by the time it is sent to AAD for exchange. AAD will respond
  // with "The provided value for the 'assertion' is not valid. The assertion has expired."
  // Retry the call of getAccessToken (no more than once). This time Office will return a
  // new unexpired middle tier token.

  if (response.error_description.indexOf("AADSTS500133") !== -1 && retryGetMiddletierToken <= 0) {
    retryGetMiddletierToken++;
    getUserData(callback);
  } else {
    dialogFallback(callback);
  }
}

export async function getGraphToken(): Promise<string> {
  // Get SSO token (use OfficeRuntime.auth for modern hosts)
  const ssoToken = await getSsoToken({
    allowSignInPrompt: true,
    allowConsentPrompt: true,
    forMSGraphAccess: false
  });

  // Exchange for Graph token via backend
  const response = await fetch('/getGraphToken', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${ssoToken}`,
      'Content-Type': 'application/json'
    }
  });

  const data = await response.json();
  return data.access_token;
}

export async function getMiddletierToken(): Promise<string> {
  // Returns a SSO token that is intended to be presented to the middle-tier
  const ssoToken = await getSsoToken({
    allowSignInPrompt: true,
    allowConsentPrompt: true,
    forMSGraphAccess: false
  });
  return ssoToken;
}