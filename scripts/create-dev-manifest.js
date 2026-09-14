const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const source = path.join(root, "manifest.xml");
const outputDirectory = path.join(root, ".dev");
const output = path.join(outputDirectory, "manifest.xml");

const productionId = "4af92278-7f2b-4a3c-bd96-4cacaf7f63d3";
const developmentId = "7af0ae9d-6054-4726-a230-e83c7ca10cff";
const productionOrigin = "https://kl3mousse.github.io/xlsnoob";
const developmentOrigin = "https://localhost:3000";

const manifest = fs.readFileSync(source, "utf8")
  .replaceAll(productionId, developmentId)
  .replaceAll(productionOrigin, developmentOrigin)
  .replace('DefaultValue="xlsNoob"', 'DefaultValue="xlsNoob Dev"')
  .replace('id="Tab.Label" DefaultValue="xlsNoob"', 'id="Tab.Label" DefaultValue="xlsNoob Dev"')
  .replace('<AppDomain>https://kl3mousse.github.io</AppDomain>', '<AppDomain>https://localhost:3000</AppDomain>');

if (!manifest.includes(developmentId) || !manifest.includes(developmentOrigin)) {
  throw new Error("Could not generate the development manifest from manifest.xml.");
}

fs.mkdirSync(outputDirectory, { recursive: true });
fs.writeFileSync(output, manifest);
console.log(`Generated ${path.relative(root, output)}`);
