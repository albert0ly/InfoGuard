/******/ (() => { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ "./src/middle-tier/msgraph-helper.ts"
/*!*******************************************!*\
  !*** ./src/middle-tier/msgraph-helper.ts ***!
  \*******************************************/
(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   getGraphData: () => (/* binding */ getGraphData),
/* harmony export */   getUserData: () => (/* binding */ getUserData)
/* harmony export */ });
/* harmony import */ var https__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! https */ "https");
/* harmony import */ var https__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(https__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _ssoauth_helper__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./ssoauth-helper */ "./src/middle-tier/ssoauth-helper.ts");
/* harmony import */ var http_errors__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! http-errors */ "http-errors");
/* harmony import */ var http_errors__WEBPACK_IMPORTED_MODULE_2___default = /*#__PURE__*/__webpack_require__.n(http_errors__WEBPACK_IMPORTED_MODULE_2__);
function _regenerator() { /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/babel/babel/blob/main/packages/babel-helpers/LICENSE */ var e, t, r = "function" == typeof Symbol ? Symbol : {}, n = r.iterator || "@@iterator", o = r.toStringTag || "@@toStringTag"; function i(r, n, o, i) { var c = n && n.prototype instanceof Generator ? n : Generator, u = Object.create(c.prototype); return _regeneratorDefine2(u, "_invoke", function (r, n, o) { var i, c, u, f = 0, p = o || [], y = !1, G = { p: 0, n: 0, v: e, a: d, f: d.bind(e, 4), d: function d(t, r) { return i = t, c = 0, u = e, G.n = r, a; } }; function d(r, n) { for (c = r, u = n, t = 0; !y && f && !o && t < p.length; t++) { var o, i = p[t], d = G.p, l = i[2]; r > 3 ? (o = l === n) && (u = i[(c = i[4]) ? 5 : (c = 3, 3)], i[4] = i[5] = e) : i[0] <= d && ((o = r < 2 && d < i[1]) ? (c = 0, G.v = n, G.n = i[1]) : d < l && (o = r < 3 || i[0] > n || n > l) && (i[4] = r, i[5] = n, G.n = l, c = 0)); } if (o || r > 1) return a; throw y = !0, n; } return function (o, p, l) { if (f > 1) throw TypeError("Generator is already running"); for (y && 1 === p && d(p, l), c = p, u = l; (t = c < 2 ? e : u) || !y;) { i || (c ? c < 3 ? (c > 1 && (G.n = -1), d(c, u)) : G.n = u : G.v = u); try { if (f = 2, i) { if (c || (o = "next"), t = i[o]) { if (!(t = t.call(i, u))) throw TypeError("iterator result is not an object"); if (!t.done) return t; u = t.value, c < 2 && (c = 0); } else 1 === c && (t = i.return) && t.call(i), c < 2 && (u = TypeError("The iterator does not provide a '" + o + "' method"), c = 1); i = e; } else if ((t = (y = G.n < 0) ? u : r.call(n, G)) !== a) break; } catch (t) { i = e, c = 1, u = t; } finally { f = 1; } } return { value: t, done: y }; }; }(r, o, i), !0), u; } var a = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} t = Object.getPrototypeOf; var c = [][n] ? t(t([][n]())) : (_regeneratorDefine2(t = {}, n, function () { return this; }), t), u = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(c); function f(e) { return Object.setPrototypeOf ? Object.setPrototypeOf(e, GeneratorFunctionPrototype) : (e.__proto__ = GeneratorFunctionPrototype, _regeneratorDefine2(e, o, "GeneratorFunction")), e.prototype = Object.create(u), e; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, _regeneratorDefine2(u, "constructor", GeneratorFunctionPrototype), _regeneratorDefine2(GeneratorFunctionPrototype, "constructor", GeneratorFunction), GeneratorFunction.displayName = "GeneratorFunction", _regeneratorDefine2(GeneratorFunctionPrototype, o, "GeneratorFunction"), _regeneratorDefine2(u), _regeneratorDefine2(u, o, "Generator"), _regeneratorDefine2(u, n, function () { return this; }), _regeneratorDefine2(u, "toString", function () { return "[object Generator]"; }), (_regenerator = function _regenerator() { return { w: i, m: f }; })(); }
function _regeneratorDefine2(e, r, n, t) { var i = Object.defineProperty; try { i({}, "", {}); } catch (e) { i = 0; } _regeneratorDefine2 = function _regeneratorDefine(e, r, n, t) { function o(r, n) { _regeneratorDefine2(e, r, function (e) { return this._invoke(r, n, e); }); } r ? i ? i(e, r, { value: n, enumerable: !t, configurable: !t, writable: !t }) : e[r] = n : (o("next", 0), o("throw", 1), o("return", 2)); }, _regeneratorDefine2(e, r, n, t); }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
/*
    This file provides the provides functionality to get Microsoft Graph data.
*/




/* global process */

var domain = "graph.microsoft.com";
var version = "v1.0";
function getUserData(_x, _x2, _x3) {
  return _getUserData.apply(this, arguments);
}
function _getUserData() {
  _getUserData = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee2(req, res, next) {
    var authorization;
    return _regenerator().w(function (_context2) {
      while (1) switch (_context2.n) {
        case 0:
          authorization = req.get("Authorization");
          _context2.n = 1;
          return (0,_ssoauth_helper__WEBPACK_IMPORTED_MODULE_1__.getAccessToken)(authorization).then(/*#__PURE__*/function () {
            var _ref = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee(graphTokenResponse) {
              var graphToken, graphUrlSegment, graphQueryParamSegment, graphData;
              return _regenerator().w(function (_context) {
                while (1) switch (_context.n) {
                  case 0:
                    if (!(graphTokenResponse && (graphTokenResponse.claims || graphTokenResponse.error))) {
                      _context.n = 1;
                      break;
                    }
                    res.send(graphTokenResponse);
                    _context.n = 3;
                    break;
                  case 1:
                    graphToken = graphTokenResponse.access_token;
                    graphUrlSegment = process.env.GRAPH_URL_SEGMENT || "/me";
                    graphQueryParamSegment = process.env.QUERY_PARAM_SEGMENT || "";
                    _context.n = 2;
                    return getGraphData(graphToken, graphUrlSegment, graphQueryParamSegment);
                  case 2:
                    graphData = _context.v;
                    // If Microsoft Graph returns an error, such as invalid or expired token,
                    // there will be a code property in the returned object set to a HTTP status (e.g. 401).
                    // Relay it to the client. It will caught in the fail callback of `makeGraphApiCall`.
                    if (graphData.code) {
                      next(http_errors__WEBPACK_IMPORTED_MODULE_2__(graphData.code, "Microsoft Graph error " + JSON.stringify(graphData)));
                    } else {
                      res.send(graphData);
                    }
                  case 3:
                    return _context.a(2);
                }
              }, _callee);
            }));
            return function (_x7) {
              return _ref.apply(this, arguments);
            };
          }()).catch(function (err) {
            res.status(401).send(err.message);
            return;
          });
        case 1:
          return _context2.a(2);
      }
    }, _callee2);
  }));
  return _getUserData.apply(this, arguments);
}
function getGraphData(_x4, _x5, _x6) {
  return _getGraphData.apply(this, arguments);
}
function _getGraphData() {
  _getGraphData = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee3(accessToken, apiUrl, queryParams) {
    return _regenerator().w(function (_context3) {
      while (1) switch (_context3.n) {
        case 0:
          return _context3.a(2, new Promise(function (resolve, reject) {
            var options = {
              host: domain,
              path: "/" + version + apiUrl + queryParams,
              method: "GET",
              headers: {
                "Content-Type": "application/json",
                Accept: "application/json",
                Authorization: "Bearer " + accessToken,
                "Cache-Control": "private, no-cache, no-store, must-revalidate",
                Expires: "-1",
                Pragma: "no-cache"
              }
            };
            https__WEBPACK_IMPORTED_MODULE_0__.get(options, function (response) {
              var body = "";
              response.on("data", function (d) {
                body += d;
              });
              response.on("end", function () {
                // The response from the OData endpoint might be an error, say a
                // 401 if the endpoint requires an access token and it was invalid
                // or expired. But a message is not an error in the call of https.get,
                // so the "on('error', reject)" line below isn't triggered.
                // So, the code distinguishes success (200) messages from error
                // messages and sends a JSON object to the caller with either the
                // requested OData or error information.

                var error;
                if (response.statusCode === 200) {
                  var parsedBody = JSON.parse(body);
                  resolve(parsedBody);
                } else {
                  error = new Error();
                  error.code = response.statusCode;
                  error.message = response.statusMessage;

                  // The error body sometimes includes an empty space
                  // before the first character, remove it or it causes an error.
                  body = body.trim();
                  error.bodyCode = JSON.parse(body).error.code;
                  error.bodyMessage = JSON.parse(body).error.message;
                  resolve(error);
                }
              });
            }).on("error", reject);
          }));
      }
    }, _callee3);
  }));
  return _getGraphData.apply(this, arguments);
}

/***/ },

/***/ "./src/middle-tier/ssoauth-helper.ts"
/*!*******************************************!*\
  !*** ./src/middle-tier/ssoauth-helper.ts ***!
  \*******************************************/
(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   getAccessToken: () => (/* binding */ getAccessToken),
/* harmony export */   validateJwt: () => (/* binding */ validateJwt)
/* harmony export */ });
/* harmony import */ var node_fetch__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! node-fetch */ "node-fetch");
/* harmony import */ var node_fetch__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(node_fetch__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var form_urlencoded__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! form-urlencoded */ "form-urlencoded");
/* harmony import */ var form_urlencoded__WEBPACK_IMPORTED_MODULE_1___default = /*#__PURE__*/__webpack_require__.n(form_urlencoded__WEBPACK_IMPORTED_MODULE_1__);
/* harmony import */ var jsonwebtoken__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! jsonwebtoken */ "jsonwebtoken");
/* harmony import */ var jsonwebtoken__WEBPACK_IMPORTED_MODULE_2___default = /*#__PURE__*/__webpack_require__.n(jsonwebtoken__WEBPACK_IMPORTED_MODULE_2__);
/* harmony import */ var jwks_rsa__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! jwks-rsa */ "jwks-rsa");
/* harmony import */ var jwks_rsa__WEBPACK_IMPORTED_MODULE_3___default = /*#__PURE__*/__webpack_require__.n(jwks_rsa__WEBPACK_IMPORTED_MODULE_3__);
function _regenerator() { /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/babel/babel/blob/main/packages/babel-helpers/LICENSE */ var e, t, r = "function" == typeof Symbol ? Symbol : {}, n = r.iterator || "@@iterator", o = r.toStringTag || "@@toStringTag"; function i(r, n, o, i) { var c = n && n.prototype instanceof Generator ? n : Generator, u = Object.create(c.prototype); return _regeneratorDefine2(u, "_invoke", function (r, n, o) { var i, c, u, f = 0, p = o || [], y = !1, G = { p: 0, n: 0, v: e, a: d, f: d.bind(e, 4), d: function d(t, r) { return i = t, c = 0, u = e, G.n = r, a; } }; function d(r, n) { for (c = r, u = n, t = 0; !y && f && !o && t < p.length; t++) { var o, i = p[t], d = G.p, l = i[2]; r > 3 ? (o = l === n) && (u = i[(c = i[4]) ? 5 : (c = 3, 3)], i[4] = i[5] = e) : i[0] <= d && ((o = r < 2 && d < i[1]) ? (c = 0, G.v = n, G.n = i[1]) : d < l && (o = r < 3 || i[0] > n || n > l) && (i[4] = r, i[5] = n, G.n = l, c = 0)); } if (o || r > 1) return a; throw y = !0, n; } return function (o, p, l) { if (f > 1) throw TypeError("Generator is already running"); for (y && 1 === p && d(p, l), c = p, u = l; (t = c < 2 ? e : u) || !y;) { i || (c ? c < 3 ? (c > 1 && (G.n = -1), d(c, u)) : G.n = u : G.v = u); try { if (f = 2, i) { if (c || (o = "next"), t = i[o]) { if (!(t = t.call(i, u))) throw TypeError("iterator result is not an object"); if (!t.done) return t; u = t.value, c < 2 && (c = 0); } else 1 === c && (t = i.return) && t.call(i), c < 2 && (u = TypeError("The iterator does not provide a '" + o + "' method"), c = 1); i = e; } else if ((t = (y = G.n < 0) ? u : r.call(n, G)) !== a) break; } catch (t) { i = e, c = 1, u = t; } finally { f = 1; } } return { value: t, done: y }; }; }(r, o, i), !0), u; } var a = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} t = Object.getPrototypeOf; var c = [][n] ? t(t([][n]())) : (_regeneratorDefine2(t = {}, n, function () { return this; }), t), u = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(c); function f(e) { return Object.setPrototypeOf ? Object.setPrototypeOf(e, GeneratorFunctionPrototype) : (e.__proto__ = GeneratorFunctionPrototype, _regeneratorDefine2(e, o, "GeneratorFunction")), e.prototype = Object.create(u), e; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, _regeneratorDefine2(u, "constructor", GeneratorFunctionPrototype), _regeneratorDefine2(GeneratorFunctionPrototype, "constructor", GeneratorFunction), GeneratorFunction.displayName = "GeneratorFunction", _regeneratorDefine2(GeneratorFunctionPrototype, o, "GeneratorFunction"), _regeneratorDefine2(u), _regeneratorDefine2(u, o, "Generator"), _regeneratorDefine2(u, n, function () { return this; }), _regeneratorDefine2(u, "toString", function () { return "[object Generator]"; }), (_regenerator = function _regenerator() { return { w: i, m: f }; })(); }
function _regeneratorDefine2(e, r, n, t) { var i = Object.defineProperty; try { i({}, "", {}); } catch (e) { i = 0; } _regeneratorDefine2 = function _regeneratorDefine(e, r, n, t) { function o(r, n) { _regeneratorDefine2(e, r, function (e) { return this._invoke(r, n, e); }); } r ? i ? i(e, r, { value: n, enumerable: !t, configurable: !t, writable: !t }) : e[r] = n : (o("next", 0), o("throw", 1), o("return", 2)); }, _regeneratorDefine2(e, r, n, t); }
function _slicedToArray(r, e) { return _arrayWithHoles(r) || _iterableToArrayLimit(r, e) || _unsupportedIterableToArray(r, e) || _nonIterableRest(); }
function _nonIterableRest() { throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return _arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray(r, a) : void 0; } }
function _arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function _iterableToArrayLimit(r, l) { var t = null == r ? null : "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (null != t) { var e, n, i, u, a = [], f = !0, o = !1; try { if (i = (t = t.call(r)).next, 0 === l) { if (Object(t) !== t) return; f = !1; } else for (; !(f = (e = i.call(t)).done) && (a.push(e.value), a.length !== l); f = !0); } catch (r) { o = !0, n = r; } finally { try { if (!f && null != t.return && (u = t.return(), Object(u) !== u)) return; } finally { if (o) throw n; } } return a; } }
function _arrayWithHoles(r) { if (Array.isArray(r)) return r; }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file defines the routes within the authRoute router.
 */






/* global process, console */

var DISCOVERY_KEYS_ENDPOINT = "https://login.microsoftonline.com/common/discovery/v2.0/keys";
function getAccessToken(_x) {
  return _getAccessToken.apply(this, arguments);
}
function _getAccessToken() {
  _getAccessToken = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee(authorization) {
    var error, scopeName, _authorization$split, _authorization$split2, assertion, tokenScopes, accessAsUserScope, formParams, stsDomain, tenant, tokenURLSegment, encodedForm, tokenResponse, json;
    return _regenerator().w(function (_context) {
      while (1) switch (_context.n) {
        case 0:
          if (authorization) {
            _context.n = 1;
            break;
          }
          error = new Error("No Authorization header was found.");
          return _context.a(2, Promise.reject(error));
        case 1:
          scopeName = process.env.SCOPE || "User.Read";
          _authorization$split = authorization.split(" "), _authorization$split2 = _slicedToArray(_authorization$split, 2), /* schema */assertion = _authorization$split2[1];
          tokenScopes = jsonwebtoken__WEBPACK_IMPORTED_MODULE_2___default().decode(assertion).scp.split(" ");
          accessAsUserScope = tokenScopes.find(function (scope) {
            return scope === "access_as_user";
          });
          if (accessAsUserScope) {
            _context.n = 2;
            break;
          }
          throw new Error("Missing access_as_user");
        case 2:
          formParams = {
            client_id: process.env.CLIENT_ID,
            client_secret: process.env.CLIENT_SECRET,
            grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
            assertion: assertion,
            requested_token_use: "on_behalf_of",
            scope: [scopeName].join(" ")
          };
          stsDomain = "https://login.microsoftonline.com";
          tenant = process.env.TENANT_ID || "common";
          tokenURLSegment = "oauth2/v2.0/token";
          encodedForm = form_urlencoded__WEBPACK_IMPORTED_MODULE_1___default()(formParams);
          _context.n = 3;
          return node_fetch__WEBPACK_IMPORTED_MODULE_0___default()("".concat(stsDomain, "/").concat(tenant, "/").concat(tokenURLSegment), {
            method: "POST",
            body: encodedForm,
            headers: {
              Accept: "application/json",
              "Content-Type": "application/x-www-form-urlencoded"
            }
          });
        case 3:
          tokenResponse = _context.v;
          _context.n = 4;
          return tokenResponse.json();
        case 4:
          json = _context.v;
          return _context.a(2, json);
        case 5:
          return _context.a(2);
      }
    }, _callee);
  }));
  return _getAccessToken.apply(this, arguments);
}
function validateJwt(req, res, next) {
  var authHeader = req.headers.authorization;
  if (authHeader) {
    var token = authHeader.split(" ")[1];

    // DECODE (don't verify yet) to see what's inside
    var decoded = jsonwebtoken__WEBPACK_IMPORTED_MODULE_2___default().decode(token, {
      complete: true
    });
    console.log("=== TOKEN PAYLOAD ===");
    console.log(JSON.stringify(decoded.payload, null, 2));
    console.log("Token audience (aud):", decoded.payload.aud);
    console.log("Expected CLIENT_ID:", process.env.CLIENT_ID);
    console.log("===================");
    var validationOptions = {
      audience: [process.env.CLIENT_ID, "api://".concat(process.env.CLIENT_ID), "api://localhost:3000/".concat(process.env.CLIENT_ID)]
    };
    jsonwebtoken__WEBPACK_IMPORTED_MODULE_2___default().verify(token, getSigningKeys, validationOptions, function (err) {
      if (err) {
        console.log(err);
        return res.sendStatus(403);
      }
      next();
    });
  }
}
function getSigningKeys(header, callback) {
  var client = new jwks_rsa__WEBPACK_IMPORTED_MODULE_3__.JwksClient({
    jwksUri: DISCOVERY_KEYS_ENDPOINT
  });
  client.getSigningKey(header.kid, function (err, key) {
    callback(null, key.getPublicKey());
  });
}

/***/ },

/***/ "cookie-parser"
/*!********************************!*\
  !*** external "cookie-parser" ***!
  \********************************/
(module) {

module.exports = require("cookie-parser");

/***/ },

/***/ "dotenv"
/*!*************************!*\
  !*** external "dotenv" ***!
  \*************************/
(module) {

module.exports = require("dotenv");

/***/ },

/***/ "express"
/*!**************************!*\
  !*** external "express" ***!
  \**************************/
(module) {

module.exports = require("express");

/***/ },

/***/ "form-urlencoded"
/*!**********************************!*\
  !*** external "form-urlencoded" ***!
  \**********************************/
(module) {

module.exports = require("form-urlencoded");

/***/ },

/***/ "http-errors"
/*!******************************!*\
  !*** external "http-errors" ***!
  \******************************/
(module) {

module.exports = require("http-errors");

/***/ },

/***/ "https"
/*!************************!*\
  !*** external "https" ***!
  \************************/
(module) {

module.exports = require("https");

/***/ },

/***/ "jsonwebtoken"
/*!*******************************!*\
  !*** external "jsonwebtoken" ***!
  \*******************************/
(module) {

module.exports = require("jsonwebtoken");

/***/ },

/***/ "jwks-rsa"
/*!***************************!*\
  !*** external "jwks-rsa" ***!
  \***************************/
(module) {

module.exports = require("jwks-rsa");

/***/ },

/***/ "morgan"
/*!*************************!*\
  !*** external "morgan" ***!
  \*************************/
(module) {

module.exports = require("morgan");

/***/ },

/***/ "node-fetch"
/*!*****************************!*\
  !*** external "node-fetch" ***!
  \*****************************/
(module) {

module.exports = require("node-fetch");

/***/ },

/***/ "office-addin-dev-certs"
/*!*****************************************!*\
  !*** external "office-addin-dev-certs" ***!
  \*****************************************/
(module) {

module.exports = require("office-addin-dev-certs");

/***/ },

/***/ "path"
/*!***********************!*\
  !*** external "path" ***!
  \***********************/
(module) {

module.exports = require("path");

/***/ }

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Check if module exists (development only)
/******/ 		if (__webpack_modules__[moduleId] === undefined) {
/******/ 			var e = new Error("Cannot find module '" + moduleId + "'");
/******/ 			e.code = 'MODULE_NOT_FOUND';
/******/ 			throw e;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/compat get default export */
/******/ 	(() => {
/******/ 		// getDefaultExport function for compatibility with non-harmony modules
/******/ 		__webpack_require__.n = (module) => {
/******/ 			var getter = module && module.__esModule ?
/******/ 				() => (module['default']) :
/******/ 				() => (module);
/******/ 			__webpack_require__.d(getter, { a: getter });
/******/ 			return getter;
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/define property getters */
/******/ 	(() => {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = (exports, definition) => {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	(() => {
/******/ 		__webpack_require__.o = (obj, prop) => (Object.prototype.hasOwnProperty.call(obj, prop))
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	(() => {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = (exports) => {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	})();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry needs to be wrapped in an IIFE because it needs to be isolated against other modules in the chunk.
(() => {
/*!********************************!*\
  !*** ./src/middle-tier/app.ts ***!
  \********************************/
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var dotenv__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! dotenv */ "dotenv");
/* harmony import */ var dotenv__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(dotenv__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var http_errors__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! http-errors */ "http-errors");
/* harmony import */ var http_errors__WEBPACK_IMPORTED_MODULE_1___default = /*#__PURE__*/__webpack_require__.n(http_errors__WEBPACK_IMPORTED_MODULE_1__);
/* harmony import */ var path__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! path */ "path");
/* harmony import */ var path__WEBPACK_IMPORTED_MODULE_2___default = /*#__PURE__*/__webpack_require__.n(path__WEBPACK_IMPORTED_MODULE_2__);
/* harmony import */ var cookie_parser__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! cookie-parser */ "cookie-parser");
/* harmony import */ var cookie_parser__WEBPACK_IMPORTED_MODULE_3___default = /*#__PURE__*/__webpack_require__.n(cookie_parser__WEBPACK_IMPORTED_MODULE_3__);
/* harmony import */ var morgan__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! morgan */ "morgan");
/* harmony import */ var morgan__WEBPACK_IMPORTED_MODULE_4___default = /*#__PURE__*/__webpack_require__.n(morgan__WEBPACK_IMPORTED_MODULE_4__);
/* harmony import */ var express__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! express */ "express");
/* harmony import */ var express__WEBPACK_IMPORTED_MODULE_5___default = /*#__PURE__*/__webpack_require__.n(express__WEBPACK_IMPORTED_MODULE_5__);
/* harmony import */ var https__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! https */ "https");
/* harmony import */ var https__WEBPACK_IMPORTED_MODULE_6___default = /*#__PURE__*/__webpack_require__.n(https__WEBPACK_IMPORTED_MODULE_6__);
/* harmony import */ var office_addin_dev_certs__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! office-addin-dev-certs */ "office-addin-dev-certs");
/* harmony import */ var office_addin_dev_certs__WEBPACK_IMPORTED_MODULE_7___default = /*#__PURE__*/__webpack_require__.n(office_addin_dev_certs__WEBPACK_IMPORTED_MODULE_7__);
/* harmony import */ var _msgraph_helper__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ./msgraph-helper */ "./src/middle-tier/msgraph-helper.ts");
/* harmony import */ var _ssoauth_helper__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ./ssoauth-helper */ "./src/middle-tier/ssoauth-helper.ts");
function _regenerator() { /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/babel/babel/blob/main/packages/babel-helpers/LICENSE */ var e, t, r = "function" == typeof Symbol ? Symbol : {}, n = r.iterator || "@@iterator", o = r.toStringTag || "@@toStringTag"; function i(r, n, o, i) { var c = n && n.prototype instanceof Generator ? n : Generator, u = Object.create(c.prototype); return _regeneratorDefine2(u, "_invoke", function (r, n, o) { var i, c, u, f = 0, p = o || [], y = !1, G = { p: 0, n: 0, v: e, a: d, f: d.bind(e, 4), d: function d(t, r) { return i = t, c = 0, u = e, G.n = r, a; } }; function d(r, n) { for (c = r, u = n, t = 0; !y && f && !o && t < p.length; t++) { var o, i = p[t], d = G.p, l = i[2]; r > 3 ? (o = l === n) && (u = i[(c = i[4]) ? 5 : (c = 3, 3)], i[4] = i[5] = e) : i[0] <= d && ((o = r < 2 && d < i[1]) ? (c = 0, G.v = n, G.n = i[1]) : d < l && (o = r < 3 || i[0] > n || n > l) && (i[4] = r, i[5] = n, G.n = l, c = 0)); } if (o || r > 1) return a; throw y = !0, n; } return function (o, p, l) { if (f > 1) throw TypeError("Generator is already running"); for (y && 1 === p && d(p, l), c = p, u = l; (t = c < 2 ? e : u) || !y;) { i || (c ? c < 3 ? (c > 1 && (G.n = -1), d(c, u)) : G.n = u : G.v = u); try { if (f = 2, i) { if (c || (o = "next"), t = i[o]) { if (!(t = t.call(i, u))) throw TypeError("iterator result is not an object"); if (!t.done) return t; u = t.value, c < 2 && (c = 0); } else 1 === c && (t = i.return) && t.call(i), c < 2 && (u = TypeError("The iterator does not provide a '" + o + "' method"), c = 1); i = e; } else if ((t = (y = G.n < 0) ? u : r.call(n, G)) !== a) break; } catch (t) { i = e, c = 1, u = t; } finally { f = 1; } } return { value: t, done: y }; }; }(r, o, i), !0), u; } var a = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} t = Object.getPrototypeOf; var c = [][n] ? t(t([][n]())) : (_regeneratorDefine2(t = {}, n, function () { return this; }), t), u = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(c); function f(e) { return Object.setPrototypeOf ? Object.setPrototypeOf(e, GeneratorFunctionPrototype) : (e.__proto__ = GeneratorFunctionPrototype, _regeneratorDefine2(e, o, "GeneratorFunction")), e.prototype = Object.create(u), e; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, _regeneratorDefine2(u, "constructor", GeneratorFunctionPrototype), _regeneratorDefine2(GeneratorFunctionPrototype, "constructor", GeneratorFunction), GeneratorFunction.displayName = "GeneratorFunction", _regeneratorDefine2(GeneratorFunctionPrototype, o, "GeneratorFunction"), _regeneratorDefine2(u), _regeneratorDefine2(u, o, "Generator"), _regeneratorDefine2(u, n, function () { return this; }), _regeneratorDefine2(u, "toString", function () { return "[object Generator]"; }), (_regenerator = function _regenerator() { return { w: i, m: f }; })(); }
function _regeneratorDefine2(e, r, n, t) { var i = Object.defineProperty; try { i({}, "", {}); } catch (e) { i = 0; } _regeneratorDefine2 = function _regeneratorDefine(e, r, n, t) { function o(r, n) { _regeneratorDefine2(e, r, function (e) { return this._invoke(r, n, e); }); } r ? i ? i(e, r, { value: n, enumerable: !t, configurable: !t, writable: !t }) : e[r] = n : (o("next", 0), o("throw", 1), o("return", 2)); }, _regeneratorDefine2(e, r, n, t); }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file is the main Node.js server file that defines the express middleware.
 */

//if (process.env.NODE_ENV !== "production") {
//  require("dotenv").config();
//}

dotenv__WEBPACK_IMPORTED_MODULE_0___default().config();










/* global console, process, require, __dirname */

var app = express__WEBPACK_IMPORTED_MODULE_5___default()();
var port = process.env.API_PORT || "3000";
app.set("port", port);

// view engine setup
app.set("views", path__WEBPACK_IMPORTED_MODULE_2__.join(__dirname, "views"));
app.set("view engine", "pug");
app.use(morgan__WEBPACK_IMPORTED_MODULE_4__("dev"));
app.use(express__WEBPACK_IMPORTED_MODULE_5___default().json());
app.use(express__WEBPACK_IMPORTED_MODULE_5___default().urlencoded({
  extended: false
}));
app.use(cookie_parser__WEBPACK_IMPORTED_MODULE_3__());

/* Turn off caching when developing */
if (true) {
  app.use(express__WEBPACK_IMPORTED_MODULE_5___default()["static"](path__WEBPACK_IMPORTED_MODULE_2__.join(process.cwd(), "dist"), {
    etag: false
  }));
  app.use(function (req, res, next) {
    res.header("Cache-Control", "private, no-cache, no-store, must-revalidate");
    res.header("Expires", "-1");
    res.header("Pragma", "no-cache");
    next();
  });
} else // removed by dead control flow
{}
var indexRouter = express__WEBPACK_IMPORTED_MODULE_5___default().Router();
indexRouter.get("/", function (req, res) {
  res.render("/taskpane.html");
});
app.use("/", indexRouter);

// Middle-tier API calls
// listen for 'ping' to verify service is running
// Un comment for development debugging, but un needed for production deployment
// app.get("/ping", function (req: any, res: any) {
//   res.send(process.platform);
// });

//app.get("/getuserdata", validateJwt, getUserData);
app.get("/getuserdata", _ssoauth_helper__WEBPACK_IMPORTED_MODULE_9__.validateJwt, _msgraph_helper__WEBPACK_IMPORTED_MODULE_8__.getUserData);

// Get the client side task pane files requested
app.get("/taskpane.html", /*#__PURE__*/function () {
  var _ref = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee(req, res) {
    return _regenerator().w(function (_context) {
      while (1) switch (_context.n) {
        case 0:
          return _context.a(2, res.sendfile("taskpane.html"));
      }
    }, _callee);
  }));
  return function (_x, _x2) {
    return _ref.apply(this, arguments);
  };
}());
app.get("/fallbackauthdialog.html", /*#__PURE__*/function () {
  var _ref2 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee2(req, res) {
    return _regenerator().w(function (_context2) {
      while (1) switch (_context2.n) {
        case 0:
          return _context2.a(2, res.sendfile("fallbackauthdialog.html"));
      }
    }, _callee2);
  }));
  return function (_x3, _x4) {
    return _ref2.apply(this, arguments);
  };
}());

// Catch 404 and forward to error handler
app.use(function (req, res, next) {
  next(http_errors__WEBPACK_IMPORTED_MODULE_1__(404));
});

// error handler
app.use(function (err, req, res) {
  // set locals, only providing error in development
  res.locals.message = err.message;
  res.locals.error = req.app.get("env") === "development" ? err : {};

  // render the error page
  res.status(err.status || 500);
  res.render("error");
});
(0,office_addin_dev_certs__WEBPACK_IMPORTED_MODULE_7__.getHttpsServerOptions)().then(function (options) {
  https__WEBPACK_IMPORTED_MODULE_6___default().createServer(options, app).listen(port, function () {
    return console.log("Server running on $3000 in ".concat("development", " mode"));
  });
});
})();

/******/ })()
;
//# sourceMappingURL=middletier.js.map