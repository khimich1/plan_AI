import { NavLink } from "react-router";

const LINKS = [
  { to: "/logistics", label: "Реестр рейсов", end: true },
  { to: "/logistics/carriers", label: "Перевозчики", end: false },
];

export const LogisticsTabs = () => (
  <nav
    aria-label="Разделы логистики"
    style={{
      display: "inline-flex",
      flexWrap: "wrap",
      padding: 4,
      gap: 4,
      borderRadius: 14,
      background: "#eef2ff",
      border: "1px solid #d6defa",
      alignSelf: "flex-start",
    }}
  >
    {LINKS.map((link) => (
      <NavLink
        key={link.to}
        to={link.to}
        end={link.end}
        style={({ isActive }) => ({
          borderRadius: 10,
          padding: "0.55rem 0.9rem",
          textDecoration: "none",
          background: isActive ? "#ffffff" : "transparent",
          color: isActive ? "#23366f" : "#475467",
          fontWeight: 600,
          boxShadow: isActive ? "0 4px 12px rgba(15, 23, 42, 0.08)" : undefined,
        })}
      >
        {link.label}
      </NavLink>
    ))}
  </nav>
);
