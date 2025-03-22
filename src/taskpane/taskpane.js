import "office-ui-fabric-react/dist/css/fabric.min.css"
import "./taskpane.css"
import App from "../components/App"
import { initializeIcons } from "@fluentui/react/lib/Icons"
import * as ReactDOM from "react-dom"
import { Office } from "office-js"

/* global document, Office */

initializeIcons()

let isOfficeInitialized = false

// The initialize function must be run each time a new page is loaded
Office.onReady(() => {
  isOfficeInitialized = true
  render(App)
}).catch((error) => {
  console.error("Error initializing Office.js:", error)
  // Still render the app, but with isOfficeInitialized=false
  render(App)
})

// Render application after Office is initialized
function render(Component) {
  ReactDOM.render(<Component isOfficeInitialized={isOfficeInitialized} />, document.getElementById("app"))
}

/* Initial render showing a progress bar */
render(App)

/* global Word console */

export async function insertText(text) {
  // Write text to the document.
  try {
    await Word.run(async (context) => {
      let body = context.document.body;
      body.insertParagraph(text, Word.InsertLocation.end);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}
