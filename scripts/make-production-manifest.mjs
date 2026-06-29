import { readFileSync, writeFileSync } from "node:fs";

const baseUrl = process.env.ADDIN_BASE_URL;

if (!baseUrl) {
  console.error("ADDIN_BASE_URL is required. Example: ADDIN_BASE_URL=https://example.com npm run build:production-manifest");
  process.exit(1);
}

const normalizedBaseUrl = baseUrl.replace(/\/+$/, "");
const manifest = readFileSync("manifest.xml", "utf8");
const productionManifest = manifest.replaceAll("https://localhost:3000", normalizedBaseUrl);

writeFileSync("manifest.production.xml", productionManifest);
console.log(`Wrote manifest.production.xml for ${normalizedBaseUrl}`);
