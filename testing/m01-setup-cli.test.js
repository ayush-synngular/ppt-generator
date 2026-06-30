const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const projectRoot = path.resolve(__dirname, "..");
const testPpt = path.join(projectRoot, "test-output.pptx");

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

describe("PPT Tool - M01 Setup & CLI Validation", () => {
  test("M01-TC-001: Verify package.json exists", () => {
    expect(fs.existsSync(path.join(projectRoot, "package.json"))).toBe(true);
  });

  test("M01-TC-002: Verify project build executes successfully", () => {
    runCommand("npm run build");
    expect(fs.existsSync(path.join(projectRoot, "dist", "index.js"))).toBe(true);
  });

  test("M01-TC-003: Verify help command displays CLI directory", () => {
    const output = runCommand("node src/index.js help");
    expect(output).toContain("PPT GENERATOR");
    expect(output).toContain("Available Commands");
    expect(output).toContain("Component Types");
  });

  test("M01-TC-004: Verify version command displays application version", () => {
    const output = runCommand("node src/index.js version");
    expect(output.trim()).toMatch(/1\.0\.0/);
  });

  test("M01-TC-005: Verify inspect command displays slide information", () => {
    expect(fs.existsSync(testPpt)).toBe(true);
    const output = runCommand("node src/index.js inspect test-output.pptx");
    expect(output).toContain("Slide count");
    expect(output).toContain("Slide titles");
  });
});