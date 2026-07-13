import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const responsePath = process.argv[2];
if (!responsePath) {
  throw new Error("Pagey response path is required");
}

const response = JSON.parse(await readFile(responsePath, "utf8"));
const site = response.site;
if (!site?.id || !site?.subdomain) {
  throw new Error(response.error || "Pagey deployment failed");
}

const scriptsDir = path.dirname(fileURLToPath(import.meta.url));
const outputPath = path.join(path.dirname(scriptsDir), "data", "deployment.json");
const publicMetadata = {
  provider: "Pagey",
  siteId: site.id,
  url: `https://${site.subdomain}.pagey.site/`,
  deployedAt: new Date().toISOString(),
  expiresAt: site.expiresAt ?? null,
};

await writeFile(outputPath, `${JSON.stringify(publicMetadata, null, 2)}\n`, "utf8");
console.log(publicMetadata.url);
