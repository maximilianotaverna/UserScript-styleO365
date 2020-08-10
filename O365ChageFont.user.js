// ==UserScript==
// @namespace   https://openuserjs.org/users/MilionMax
// @name        Office365 font change
// @description Change the look of office 365
// @match       https://outlook.office.com/*
// @include     https://teams.microsoft.com/*
// @include     https://*.sharepoint.com/*
// @grant       GM_addStyle
// @run-at      document-load 
// @version     1.0
// @license     MIT
// ==/UserScript==

function addGlobalStyle(css) {
    var head, style;
    head = document.getElementsByTagName('head')[0];
    if (!head) { return; }
    style = document.createElement('style');
    style.type = 'text/css';
    style.innerHTML = css;
    head.appendChild(style);
}

GM_addStyle('@import url(https://fonts.googleapis.com/css?family=Roboto|Nunito|Rubik|DM+Sans|Manrope&display=swap);');
//addGlobalStyle('body, h1, h2, h3, span, .app-small-font, .ms-DetailsRow-cell, .ms-DetailsRow-cell .root-73 { font-family: LiHeiPro!important; }');
//addGlobalStyle('body, h1, h2, h3, span, .app-small-font, .ms-DetailsRow-cell, .ms-DetailsRow-cell .root-73 { font-family: Montserrat-Light!important; }');
addGlobalStyle('body, h1, h2, h3, span, .app-small-font, .ms-DetailsRow-cell, .ms-DetailsRow-cell .root-73 { font-family: Manrope!important; }');
addGlobalStyle('.css-47, .css-52, .css-81, .css-82, .css-83, .css-84, .css-85, .css-86, .css-87, .css-88, .css-89, .css-90, .css-255, .css-256, ms-Icon, div#headerButtonsRegionId span { font-family: controlIcons!important; }');
