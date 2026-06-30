const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const projectRoot = path.resolve(__dirname, "..");
const outputDir = path.join(projectRoot, "test-output");
const dataDir = path.join(outputDir, "template-data");

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

describe("PPT Tool - M09 Template Components", () => {

  test("M09-TC-001: Generate Executive Summary template", () => {
    const outputFile = path.join(outputDir, "m09-auto-executive-summary.pptx");

    runCommand(
      `node src/index.js template executive-summary --data ${path.join(dataDir, "executive-summary.json")} --output test-output/m09-auto-executive-summary.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M09-TC-002: Generate Project Status template", () => {
    const outputFile = path.join(outputDir, "m09-auto-project-status.pptx");

    runCommand(
      `node src/index.js template project-status --data ${path.join(dataDir, "project-status.json")} --output test-output/m09-auto-project-status.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M09-TC-003: Generate Quarterly Business Review template", () => {
    const outputFile = path.join(outputDir, "m09-auto-quarterly-business-review.pptx");

    runCommand(
      `node src/index.js template quarterly-business-review --data ${path.join(dataDir, "quarterly-business-review.json")} --output test-output/m09-auto-quarterly-business-review.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M09-TC-004: Inspect generated Executive Summary presentation", () => {
    const output = runCommand(
      `node src/index.js inspect test-output/m09-auto-executive-summary.pptx`
    );

    expect(output).toContain("Slide count");
    expect(output).toContain("Slide titles");
  });

  test("M09-TC-005: Verify generated template presentation file", () => {
    const outputFile = path.join(outputDir, "m09-auto-project-status.pptx");

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

});
