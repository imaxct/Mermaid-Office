/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import mermaid from "mermaid";
import { Canvg, presets } from "canvg";

const preset = presets.offscreen();

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("insert-image").onclick = run;
    mermaid.initialize({ startOnLoad: false });
  }
});

export async function run() {
  return Word.run(async (context) => {
    const wrapperDiv = document.getElementById("graph");
    const { svg } = await mermaid.render("graphDiv", "graph TB\n AA-->BB\n AA-->CC");
    console.log(svg);
    wrapperDiv.innerHTML = svg;
    const canvas = new OffscreenCanvas(300, 300);
    const ctx = canvas.getContext("2d");
    const v = await Canvg.from(ctx, svg, preset);
    await v.render();
    const blob = await canvas.convertToBlob();
    context.document.body.insertInlinePictureFromBase64();
    await context.sync();
  });
}
