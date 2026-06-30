const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const projectRoot = path.resolve(__dirname, "..");
const outputDir = path.join(projectRoot, "test-output");

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

describe("PPT Tool - M04 Table Components", () => {
  test("M04-TC-001: Create basic table PPT successfully", () => {
    const outputFile = path.join(outputDir, "m04-auto-table-basic.pptx");

    runCommand(
      `node src/index.js create table --rows '[["Name","Role"],["Diya","QA"],["Ayush","Developer"]]' --output test-output/m04-auto-table-basic.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M04-TC-002: Inspect basic table PPT successfully", () => {
    const output = runCommand(
      `node src/index.js inspect test-output/m04-auto-table-basic.pptx`
    );

    expect(output).toContain("Slide count: 1");
    expect(output).toContain("Name");
  });

  test("M04-TC-003: Create table with multiple rows successfully", () => {
    const outputFile = path.join(outputDir, "m04-auto-table-rows.pptx");

    runCommand(
      `node src/index.js create table --rows '[["Module","Status"],["Text","Pass"],["Table","Testing"],["Chart","Pending"]]' --output test-output/m04-auto-table-rows.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M04-TC-004: Create table with custom position and width successfully", () => {
    const outputFile = path.join(outputDir, "m04-auto-table-position.pptx");

    runCommand(
      `node src/index.js create table --rows '[["Module","Status"],["Text","Pass"],["Table","Testing"]]' --x 1 --y 1 --w 7 --output test-output/m04-auto-table-position.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M04-TC-005: Inspect positioned table PPT successfully", () => {
    const output = runCommand(
      `node src/index.js inspect test-output/m04-auto-table-position.pptx`
    );

    expect(output).toContain("Slide count: 1");
    expect(output).toContain("Module");
  });
});