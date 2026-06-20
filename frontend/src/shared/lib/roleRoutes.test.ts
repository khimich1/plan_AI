import { describe, expect, it } from "vitest";

import { defaultRouteForRole } from "./roleRoutes";

describe("defaultRouteForRole", () => {
  it("sends production users to /production", () => {
    expect(defaultRouteForRole("production")).toBe("/production");
  });

  it("sends admin and manager users to /new", () => {
    expect(defaultRouteForRole("admin")).toBe("/new");
    expect(defaultRouteForRole("manager")).toBe("/new");
  });

  it("falls back to /new for unknown roles", () => {
    expect(defaultRouteForRole(undefined)).toBe("/new");
    expect(defaultRouteForRole("guest")).toBe("/new");
  });
});
