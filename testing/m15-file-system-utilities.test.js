const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const projectRoot = path.resolve(__dirname, "..");
const testDir = "test-output/m15-auto";
const testFile = `${testDir}/sample.txt`;

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

describe("PPT Tool - M15 File System Utilities", () => {

  test("M15-TC-001: Create directory successfully", () => {
    const output = runCommand(
      `node src/index.js create-dir --path ${testDir}`
    );

    expect(output).toContain("Directory created");
    expect(fs.existsSync(path.join(projectRoot, testDir))).toBe(true);
  });

  test("M15-TC-002: Write file successfully", () => {
    const output = runCommand(
      `node src/index.js write-file --path ${testFile} --content "Hello PPT Tool"`
    );

    expect(output).toContain("Written to");
    expect(fs.existsSync(path.join(projectRoot, testFile))).toBe(true);
  });

  test("M15-TC-003: Read file successfully", () => {
    const output = runCommand(
      `node src/index.js read-file --path ${testFile}`
    );

    expect(output).toContain("Hello PPT Tool");
  });

  test("M15-TC-004: Edit file successfully", () => {
    const output = runCommand(
      `node src/index.js edit-file --path ${testFile} --search "PPT" --replace "Presentation"`
    );

    expect(output).toContain("File edited");

    const verify = runCommand(
      `node src/index.js read-file --path ${testFile}`
    );

    expect(verify).toContain("Hello Presentation Tool");
  });

  test("M15-TC-005: Grep text successfully", () => {
    const output = runCommand(
      `node src/index.js grep --pattern "Presentation" --path ${testDir}`
    );

    expect(output).toContain("Presentation");
  });

  test("M15-TC-006: Glob files successfully", () => {
    const output = runCommand(
      `node src/index.js glob --pattern "*.txt" --path ${testDir}`
    );

    expect(output).toContain("sample.txt");
  });

  test("M15-TC-007: List directory successfully", () => {
    const output = runCommand(
      `node src/index.js list-dir --path ${testDir}`
    );

    expect(output).toContain("sample.txt");
  });

});