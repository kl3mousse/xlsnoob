/// <reference types="office-js" />

import { openContainingFolder, openInfo } from "./file";

Object.assign(globalThis, {
  openContainingFolder,
  openInfo,
});
