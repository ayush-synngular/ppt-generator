const fs = require("fs");
const path = require("path");
const { Presentation } = require("../src/index");

const projectRoot = path.resolve(__dirname, "..");
const outputDir = path.join(projectRoot, "test-output", "automation");
const outputPpt = path.join(outputDir, "m02-core-presentation.pptx");

beforeEach(() => {
  fs.mkdirSync(outputDir, { recursive: true });
  if (fs.existsSync(outputPpt)) {
    fs.unlinkSync(outputPpt);
  }
});

describe("PPT Tool - M02 Core Presentation Operations", () => {
  test("M02-TC-001: Create presentation object successfully", () => {
    const ppt = new Presentation("Automation Core Test");

    expect(ppt).toBeDefined();
    expect(ppt.title).toContain("Automation Core Test");
  });

  test("M02-TC-002: Update presentation title successfully", () => {
    const ppt = new Presentation("Initial Title");

    ppt.setTitle("Updated Title");

    expect(ppt.title).toContain("Updated Title");
  });

  test("M02-TC-003: Update presentation subject successfully", () => {
    const ppt = new Presentation("Subject Test");

    ppt.setSubject("Automation Subject");

    expect(ppt.pptx.subject).toContain("Automation Subject");
  });

  test("M02-TC-004: Add slide successfully", () => {
    const ppt = new Presentation("Slide Test");

    ppt.addSlide();

    expect(ppt.slides.length).toBe(1);
  });

  test("M02-TC-005: Add multiple slides successfully", () => {
    const ppt = new Presentation("Multiple Slides Test");

    ppt.addSlide();
    ppt.addSlide();
    ppt.addSlide();

    expect(ppt.slides.length).toBe(3);
  });

  test("M02-TC-006: Save presentation successfully", async () => {
    const ppt = new Presentation("Save Test");

    ppt.addSlide();
    await ppt.save(outputPpt);

    expect(fs.existsSync(outputPpt)).toBe(true);
    expect(fs.statSync(outputPpt).size).toBeGreaterThan(0);
  });
});