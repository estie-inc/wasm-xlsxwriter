import { Image } from "../web";
import { describe, test, beforeAll, beforeEach, afterEach, expect, vi } from "vitest";
import { initWasModule } from "./common";

beforeAll(async () => {
  await initWasModule();
});

describe("xlsx-wasm test", () => {
  // spy console.error to test panic hook
  let consoleErrorSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    consoleErrorSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    consoleErrorSpy.mockRestore();
  });

  test("panic hook test", async () => {
    // Arrange
    // Create an empty image buffer to trigger panic hook
    const imageBuf = Buffer.from([]);

    // Act
    try {
      new Image(imageBuf);
    } catch (e) {
      // ignore
    }

    // Assert
    expect(consoleErrorSpy).toHaveBeenCalledOnce();
    expect(consoleErrorSpy.mock.calls[0]?.[0]).toContain("out of range for slice of length 0");
  });
});
