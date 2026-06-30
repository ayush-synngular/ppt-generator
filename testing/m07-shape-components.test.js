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

describe("PPT Tool - M07 Shape Components", () => {

  test("M07-TC-001: Create rectangle successfully", () => {
    const outputFile = path.join(outputDir, "m07-auto-rect.pptx");

    runCommand(
      `node src/index.js create shape --type rect --output test-output/m07-auto-rect.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M07-TC-002: Create circle successfully", () => {
    const outputFile = path.join(outputDir, "m07-auto-circle.pptx");

    runCommand(
      `node src/index.js create shape --type circle --output test-output/m07-auto-circle.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M07-TC-003: Create line successfully", () => {
    const outputFile = path.join(outputDir, "m07-auto-line.pptx");

    runCommand(
      `node src/index.js create shape --type line --output test-output/m07-auto-line.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M07-TC-004: Create arrow successfully", () => {
    const outputFile = path.join(outputDir, "m07-auto-arrow.pptx");

    runCommand(
      `node src/index.js create shape --type arrow --output test-output/m07-auto-arrow.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M07-TC-005: Create custom rectangle successfully", () => {
    const outputFile = path.join(outputDir, "m07-auto-custom.pptx");

    runCommand(
      `node src/index.js create shape --type rect --color FF0000 --x 2 --y 1 --w 4 --h 2 --output test-output/m07-auto-custom.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

});
