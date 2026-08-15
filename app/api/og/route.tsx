import { ImageResponse } from "next/og";
import { listPublishedCompanies } from "@/lib/db/queries";

export const runtime = "nodejs";

export async function GET(request: Request) {
  const { searchParams } = new URL(request.url);
  const slugs = (searchParams.get("c") ?? "").split(",").filter(Boolean).slice(0, 4);

  const all = await listPublishedCompanies();
  const names = slugs.map((slug) => all.find((c) => c.slug === slug)?.name ?? slug);

  return new ImageResponse(
    (
      <div
        style={{
          width: "100%",
          height: "100%",
          display: "flex",
          flexDirection: "column",
          justifyContent: "center",
          alignItems: "center",
          background: "linear-gradient(to bottom right, #eef2ff, #ffffff)",
          padding: 64,
        }}
      >
        <div style={{ fontSize: 32, color: "#4338ca", fontWeight: 700, marginBottom: 24 }}>PerkStack</div>
        <div style={{ fontSize: 48, fontWeight: 700, color: "#111827", textAlign: "center", display: "flex", flexWrap: "wrap", justifyContent: "center", gap: 12 }}>
          {names.map((name, i) => (
            <span key={i} style={{ display: "flex" }}>
              {name}
              {i < names.length - 1 ? <span style={{ color: "#9ca3af", marginLeft: 12 }}>vs</span> : null}
            </span>
          ))}
        </div>
        <div style={{ fontSize: 24, color: "#6b7280", marginTop: 24 }}>Compare sourced company benefits</div>
      </div>
    ),
    { width: 1200, height: 630 }
  );
}
