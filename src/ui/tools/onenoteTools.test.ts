import { beforeEach, describe, expect, it, vi } from "vitest";
import { navigateToPage } from "./navigateToPage";
import { setNoteSelection } from "./setNoteSelection";
import { setPageTitle } from "./setPageTitle";

const oneNoteRunMock = vi.fn();

(globalThis as any).OneNote = {
  run: oneNoteRunMock,
};

describe("onenote tools", () => {
  beforeEach(() => {
    oneNoteRunMock.mockReset();
    (globalThis as any).Office = {
      context: {
        requirements: {
          isSetSupported: vi.fn(() => true),
        },
      },
    };
  });

  it("validates navigate_to_page inputs before host calls", async () => {
    const result = await navigateToPage.handler({});
    expect(typeof result).toBe("object");
    expect((result as any).resultType).toBe("failure");
  });

  it("validates clientUrl navigation by checking the active page", async () => {
    oneNoteRunMock.mockImplementation(async (callback: (context: any) => Promise<unknown>) => {
      const requestedPage = {
        id: "requested-page",
        title: "Requested",
        clientUrl: "https://example.invalid/requested",
        load: vi.fn(),
      };
      const activePage = {
        id: "different-page",
        title: "Different",
        clientUrl: "https://example.invalid/different",
        load: vi.fn(),
      };
      const context = {
        application: {
          navigateToPageWithClientUrl: vi.fn(() => requestedPage),
          getActivePage: vi.fn(() => activePage),
        },
        sync: vi.fn(async () => undefined),
      };

      return await callback(context);
    });

    const result = await navigateToPage.handler({ clientUrl: "https://example.invalid/requested" });
    expect((result as any).resultType).toBe("failure");
    expect((result as any).error).toMatch(/did not navigate/);
  });

  it("rejects empty OneNote selection updates", async () => {
    const result = await setNoteSelection.handler({ content: "   " });
    expect((result as any).resultType).toBe("failure");
  });

  it("rejects empty page titles", async () => {
    const result = await setPageTitle.handler({ title: "   " });
    expect((result as any).resultType).toBe("failure");
  });
});
