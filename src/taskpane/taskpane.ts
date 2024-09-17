/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import mermaid from "mermaid";
import { toBase64 } from "js-base64";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    mermaid.initialize({ startOnLoad: false });
    document.getElementById("insert-image").onclick = run;
  }
});

const getSvgElement = () => {
  const svgElement = document.querySelector("#graph svg")?.cloneNode(true) as HTMLElement;
  svgElement.setAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink");
  return svgElement;
};

const getBase64SVG = (svg?: HTMLElement, width?: number, height?: number): string => {
  if (svg) {
    // Prevents the SVG size of the interface from being changed
    svg = svg.cloneNode(true) as HTMLElement;
  }
  height && svg?.setAttribute("height", `${height}px`);
  width && svg?.setAttribute("width", `${width}px`);
  if (!svg) {
    svg = getSvgElement();
  }
  const svgString = svg.outerHTML
    .replaceAll("<br>", "<br/>")
    .replaceAll(/<img([^>]*)>/g, (m, g: string) => `<img ${g} />`);

  return toBase64(`<?xml version="1.0" encoding="UTF-8"?>
${svgString}`);
};

const getBase64Png = () => {
  const canvas: HTMLCanvasElement = document.createElement("canvas");
  const svg = getSvgElement();
  const box = svg.viewBox.baseVal;
  canvas.width = box.width;
  canvas.height = box.height;
  const context = canvas.getContext("2d");
  // context.fillRect(0, 0, canvas.width, canvas.height);
  return new Promise<string>((resolve) => {
    const image = new Image();
    image.src = `data:image/svg+xml;base64,${getBase64SVG(svg, canvas.width, canvas.height)}`;
    image.addEventListener("load", () => {
      context.drawImage(image, 0, 0, canvas.width, canvas.height);
      resolve(canvas.toDataURL("image/png").replace("data:image/png;base64,", ""));
    });
  });
};

export async function run() {
  return Word.run(async (context) => {
    const wrapperDiv = document.getElementById("graph");
    const { svg } = await mermaid.render("graphDiv", "graph TB\n AA-->BB\n AA-->CC");
    wrapperDiv.innerHTML = svg;
    const png = await getBase64Png();
    context.document.body.insertInlinePictureFromBase64(png, Word.InsertLocation.end);
    await context.sync();
  });
}
