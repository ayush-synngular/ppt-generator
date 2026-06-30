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

describe("PPT Tool - M05 Chart Components", () => {
  test("M05-TC-001: Create bar chart PPT successfully", () => {
    const outputFile = path.join(outputDir, "m05-auto-chart-bar.pptx");

    runCommand(
      `node src/index.js create chart --type bar --title "Monthly Sales" --data '[["Jan",10],["Feb",20],["Mar",30]]' --output test-output/m05-auto-chart-bar.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M05-TC-002: Create line chart PPT successfully", () => {
    const outputFile = path.join(outputDir, "m05-auto-chart-line.pptx");

    runCommand(
      `node src/index.js create chart --type line --title "Growth Trend" --data '[["Q1",15],["Q2",25],["Q3",40]]' --output test-output/m05-auto-chart-line.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M05-TC-003: Create pie chart PPT successfully", () => {
    const outputFile = path.join(outputDir, "m05-auto-chart-pie.pptx");

    runCommand(
      `node src/index.js create chart --type pie --title "Market Share" --data '[["Product A",50],["Product B",30],["Product C",20]]' --output test-output/m05-auto-chart-pie.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M05-TC-004: Create chart with custom dimensions successfully", () => {
    const outputFile = path.join(outputDir, "m05-auto-chart-custom.pptx");

    runCommand(
      `node src/index.js create chart --type bar --title "Custom Size Chart" --data '[["A",5],["B",10]]' --x 1 --y 1 --w 6 --h 3 --output test-output/m05-auto-chart-custom.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M05-TC-005: Inspect generated chart PPT successfully", () => {
    const output = runCommand(
      `node src/index.js inspect test-output/m05-auto-chart-bar.pptx`
    );

    expect(output).toContain("Slide count: 1");
    expect(output).toContain("Slide titles");
  });
});
