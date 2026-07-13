import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const scriptsDir = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.dirname(scriptsDir);
const outputDir = path.join(rootDir, "artifacts");

const [html, css, engineSource, appSource] = await Promise.all([
  readFile(path.join(rootDir, "index.html"), "utf8"),
  readFile(path.join(rootDir, "styles.css"), "utf8"),
  readFile(path.join(rootDir, "engine.js"), "utf8"),
  readFile(path.join(rootDir, "app.js"), "utf8"),
]);

const standalone = html
  .replace(
    '<link rel="stylesheet" href="./styles.css" />',
    `<style>\n${css}\n</style>`,
  )
  .replace(
    '<script src="./app.js" type="module"></script>',
    `<script type="module">\n${engineSource}\n${appSource.replace(
      'import { createExcelEngine } from "./engine.js";\n',
      "",
    )}\n</script>`,
  );

if (standalone.includes('src="./app.js"') || standalone.includes('href="./styles.css"')) {
  throw new Error("Standalone asset replacement failed");
}

await mkdir(outputDir, { recursive: true });
const outputPath = path.join(outputDir, "univ-admission-standalone.html");
await writeFile(outputPath, standalone, "utf8");
console.log(outputPath);

