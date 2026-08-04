import { describe, expect, it } from "vitest";

import { canAccessRoute, defaultRouteForRole } from "./roleRoutes";

describe("defaultRouteForRole", () => {
  it("sends production users to /production", () => {
    expect(defaultRouteForRole("production")).toBe("/production");
  });

  it("sends admin and manager users to /new", () => {
    expect(defaultRouteForRole("admin")).toBe("/new");
    expect(defaultRouteForRole("manager")).toBe("/new");
  });

  it("sends logistics users to /logistics", () => {
    expect(defaultRouteForRole("logistics")).toBe("/logistics");
  });

  it("falls back to /new for unknown roles", () => {
    expect(defaultRouteForRole(undefined)).toBe("/new");
    expect(defaultRouteForRole("guest")).toBe("/new");
  });
});

describe("canAccessRoute", () => {
  it("allows admin to access all guarded routes", () => {
    expect(canAccessRoute("admin", "/new")).toBe(true);
    expect(canAccessRoute("admin", "/archive")).toBe(true);
    expect(canAccessRoute("admin", "/production")).toBe(true);
  });

  it("allows manager only on commercial routes", () => {
    expect(canAccessRoute("manager", "/new")).toBe(true);
    expect(canAccessRoute("manager", "/archive")).toBe(true);
    expect(canAccessRoute("manager", "/production")).toBe(false);
  });

  it("allows production only on production route", () => {
    expect(canAccessRoute("production", "/new")).toBe(false);
    expect(canAccessRoute("production", "/archive")).toBe(false);
    expect(canAccessRoute("production", "/production")).toBe(true);
  });

  it("denies unknown and missing roles on guarded routes", () => {
    expect(canAccessRoute(undefined, "/new")).toBe(false);
    expect(canAccessRoute("guest", "/archive")).toBe(false);
    expect(canAccessRoute("guest", "/production")).toBe(false);
  });

  it("allows admin and logistics on logistics routes", () => {
    expect(canAccessRoute("admin", "/logistics")).toBe(true);
    expect(canAccessRoute("logistics", "/logistics")).toBe(true);
    expect(canAccessRoute("logistics", "/logistics/carriers")).toBe(true);
  });

  it("denies manager and production on logistics routes", () => {
    expect(canAccessRoute("manager", "/logistics")).toBe(false);
    expect(canAccessRoute("production", "/logistics")).toBe(false);
    expect(canAccessRoute("manager", "/logistics/carriers")).toBe(false);
    expect(canAccessRoute(undefined, "/logistics")).toBe(false);
  });

  it("normalizes paths without a leading slash", () => {
    expect(canAccessRoute("manager", "new")).toBe(true);
    expect(canAccessRoute("production", "production")).toBe(true);
    expect(canAccessRoute("manager", "production")).toBe(false);
  });

  it("allows any role on unguarded paths", () => {
    expect(canAccessRoute("production", "/login")).toBe(true);
    expect(canAccessRoute(undefined, "/unknown")).toBe(true);
  });
});
