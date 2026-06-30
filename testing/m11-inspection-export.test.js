const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const projectRoot = path.resolve(__dirname, "..");
const outputDir = path.join(projectRoot, "test-output");
const inputPpt = "test-output/m03-text-basic.pptx";

function runCommand(command) {
  try {
    return execSync(command, {
      cwd: projectRoot,
      encoding: "utf8",
      stdio: ["pipe", "pipe", "pipe"],
    });
  } catch (error) {
    return `${error.stdout || ""}${error.stderr || ""}${error.message || ""}`;
  }
}

describe("PPT Tool - M11 Presentation Inspection & Export", () => {
  test("M11-TC-001: Inspect valid PPT successfully", () => {
    const output = runCommand(`node src/index.js inspect ${inputPpt}`);

    expect(output).toContain("Slide count: 1");
    expect(output).toContain("QA Text Test");
  });

  test("M11-TC-002: Export PPT to JSON successfully", () => {
    const outputFile = path.join(outputDir, "m11-auto-export.json");

    const output = runCommand(
      `node src/index.js export --input ${inputPpt} --format json --output test-output/m11-auto-export.json`
    );

    expect(output).toContain("Export successful");
    expect(fs.existsSync(outputFile)).toBe(true);

    const exportedJson = JSON.parse(fs.readFileSync(outputFile, "utf8"));
    expect(exportedJson.slides.length).toBeGreaterThan(0);
  });

  test("M11-TC-003: Export PPT to XML successfully", () => {
    const outputFile = path.join(outputDir, "m11-auto-export.xml");

    const output = runCommand(
      `node src/index.js export --input ${inputPpt} --format xml --output test-output/m11-auto-export.xml`
    );

    expect(output).toContain("Export successful");
    expect(fs.existsSync(outputFile)).toBe(true);
  });

  test("M11-TC-004: Handle unsupported export format validation", () => {
    const output = runCommand(
      `node src/index.js export --input ${inputPpt} --format csv --output test-output/m11-auto-export.csv`
    );

    expect(output).toContain("Unsupported format");
  });

  test("M11-TC-005: Handle missing input validation", () => {
    const output = runCommand(`node src/index.js export`);

    expect(output).toContain("Missing input file");
  });
});
