/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import mermaid from "mermaid";
import { toBase64, toUint8Array, fromUint8Array, fromBase64 } from "js-base64";
import { metadataPNG } from "./metadataPNG";

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("errors").innerText = "";
    mermaid.initialize({ startOnLoad: false });
    document.getElementById("insert-image").onclick = GenerateImage;
    await ExtractMermaidFromSelection();
  }
});

const GenerateImage = async () => {
  return Word.run(async (context) => {
    const wrapperDiv = document.getElementById("graph");
    const inputSyntax = document.getElementById("mermaid-input").value.trim();
    document.getElementById("errors").innerText = "";
    try {
      const { svg } = await mermaid.render("graphDiv", inputSyntax);
      wrapperDiv.innerHTML = svg;
      const pngBase64 = await getBase64Png();
      const encoded = toBase64(inputSyntax);
      const newImage = metadataPNG.savetEXt(toUint8Array(pngBase64), `mermaid:${encoded}`);
      const newBase64 = fromUint8Array(newImage);
      context.document.getSelection().insertInlinePictureFromBase64(newBase64, Word.InsertLocation.replace);
    } catch (error) {
      document.getElementById("errors").innerText = error.toString();
      console.log(error);
    }
    await context.sync();
  });
};

const ExtractMermaidFromSelection = async () => {
  return Word.run(async (context) => {
    const selection = context.document.getSelection().load();
    await context.sync();
    const pictures = selection.inlinePictures.load();
    await context.sync();
    for (let i = 0; i < pictures.items.length; i++) {
      const base64Src = pictures.items[i].getBase64ImageSrc();
      await context.sync();
      const imageBytes = toUint8Array(base64Src.value);
      if (metadataPNG.isPNG(imageBytes)) {
        const textChunk = metadataPNG.gettEXt(imageBytes);
        if (textChunk === undefined || !textChunk.startsWith("mermaid:")) {
          continue;
        }
        const mermaidCodeBase64 = textChunk.replace("mermaid:", "");
        const mermaidCode = fromBase64(mermaidCodeBase64);
        document.getElementById("mermaid-input").value = mermaidCode;
        pictures.items[i].select();
        break;
      }
    }
  });
};

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
