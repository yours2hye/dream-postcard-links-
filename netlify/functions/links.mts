import type { Context, Config } from "@netlify/functions";
import * as XLSX from "xlsx";

const SHEET_ID = Netlify.env.get("SHEET_ID") || "1UJdKmCo94XOlFSIqcFFLY3SMBOa322t92GkCJg08X6g";
const EXPORT_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=xlsx`;

const CATEGORY_MAP: Record<string, { icon: string; label: string }> = {
  "운영": { icon: "🗂️", label: "운영" },
  "기획": { icon: "📝", label: "기획" },
  "이벤트": { icon: "🎉", label: "이벤트" },
  "심사": { icon: "🏆", label: "심사·시상" },
  "심사영상": { icon: "🏆", label: "심사·시상" },
  "시상": { icon: "🏆", label: "심사·시상" },
  "실적": { icon: "📈", label: "실적" },
  "고도화": { icon: "🖥️", label: "시스템·고도화" },
  "시스템": { icon: "🖥️", label: "시스템·고도화" },
  "운영 시스템": { icon: "🖥️", label: "시스템·고도화" },
  "운영 만족팀": { icon: "☎️", label: "만족팀" },
  "랜딩": { icon: "🎯", label: "랜딩" },
  "랜딩대시보드": { icon: "🎯", label: "랜딩" },
  "학교DB": { icon: "🏫", label: "학교 DB" },
  "에이전시": { icon: "🤝", label: "에이전시 관리" },
  "에이전시 관리": { icon: "🤝", label: "에이전시 관리" },
  "GIK": { icon: "🤝", label: "GIK·타부서협업" },
  "지사본": { icon: "🏢", label: "지사본" },
};
const FALLBACK_SECTION = { icon: "📌", label: "기타" };
const ENV_SECTION = { icon: "🔐", label: "운영·개발 환경" };

type Cell = { text: string; url: string | null; hasOwnLink: boolean };
type Item = { title: string; url: string; subtitle: string | null; extra: { label: string; url: string } | null; favorite: boolean };
type Section = { key: string; icon: string; label: string; items: Item[] };

let cache: { at: number; data: unknown } | null = null;
const TTL_MS = 5 * 60 * 1000;

function normalizeCategory(text: string) {
  return text.trim().replace(/_/g, " ").replace(/\s+/g, " ");
}

function cellInfo(raw: XLSX.CellObject | undefined): Cell {
  const text = (raw?.v ?? "").toString().trim();
  const link = (raw as any)?.l?.Target as string | undefined;
  if (link) return { text, url: link, hasOwnLink: true };
  if (/^https?:\/\//i.test(text)) return { text, url: text, hasOwnLink: false };
  return { text, url: null, hasOwnLink: false };
}

async function loadData() {
  const res = await fetch(EXPORT_URL);
  if (!res.ok) throw new Error(`sheet export failed: ${res.status}`);
  const buf = Buffer.from(await res.arrayBuffer());
  const wb = XLSX.read(buf, { type: "buffer" });

  const hiddenFlags = wb.Workbook?.Sheets ?? [];
  const visibleIndex = wb.SheetNames.findIndex((_, i) => !hiddenFlags[i]?.Hidden);
  const sheetName = wb.SheetNames[visibleIndex >= 0 ? visibleIndex : 0];
  const ws = wb.Sheets[sheetName];
  const range = XLSX.utils.decode_range(ws["!ref"] || "A1");

  const sections = new Map<string, Section>();
  const favorites: Item[] = [];
  let pendingSection: { icon: string; label: string } | null = null;
  let pendingEnvTag: string | null = null;

  for (let r = range.s.r; r <= range.e.r; r++) {
    const cols = [0, 1, 2, 3].map((c) => cellInfo(ws[XLSX.utils.encode_cell({ r, c })]));
    const [a] = cols;

    // The sheet is hand-maintained and not strictly schema'd: most rows are
    // "category | title | subtitle", some rows put a raw URL directly in a
    // cell instead of a label, and a few rows are bare section headers
    // ("1. 운영환경") followed by dash-prefixed sub-items with no category.
    // Scan B/C/D/A in that priority order for the first usable link.
    const candidates = [cols[1], cols[2], cols[3], cols[0]];
    const primary = candidates.find((cell) => cell.url);

    if (!primary) {
      const label = a.text;
      if (!label) continue; // blank spacer row
      const envMatch = label.match(/^\d+\.\s*(운영|개발)/);
      if (envMatch) {
        pendingSection = ENV_SECTION;
        pendingEnvTag = envMatch[1];
      }
      continue; // header-only row, nothing to link
    }

    const catInfo = CATEGORY_MAP[normalizeCategory(a.text)];
    const pool = cols.filter((cell) => cell !== primary && !(catInfo && cell === a));

    let title: string;
    if (primary.hasOwnLink) {
      title = primary.text;
    } else {
      const titleCell = pool.find((cell) => cell.text && !cell.url);
      title = titleCell ? titleCell.text : new URL(primary.url!).hostname;
      if (titleCell) pool.splice(pool.indexOf(titleCell), 1);
    }
    // clean up leftover "- 어드민: " style dashes/colons from env sub-item rows
    title = title.replace(/^-\s*/, "").replace(/:\s*$/, "").trim() || title;

    const subtitleCell = pool.find((cell) => cell.text && !cell.url);
    let subtitle = subtitleCell ? subtitleCell.text : null;
    if (subtitleCell) pool.splice(pool.indexOf(subtitleCell), 1);

    const extraCell = pool.find((cell) => cell.url);
    const extra = extraCell ? { label: extraCell.hasOwnLink ? extraCell.text : "관련 링크", url: extraCell.url! } : null;

    const favorite = title.startsWith("★");
    if (favorite) title = title.replace(/^★\s*/, "");

    const sectionMeta = catInfo ?? pendingSection ?? FALLBACK_SECTION;
    if (sectionMeta === ENV_SECTION && pendingEnvTag && !subtitle) {
      subtitle = `${pendingEnvTag} 환경`;
    }

    const item: Item = { title, url: primary.url!, subtitle, extra, favorite };
    if (favorite) favorites.push(item);

    const sectionKey = sectionMeta.label;
    if (!sections.has(sectionKey)) {
      sections.set(sectionKey, { key: sectionKey, icon: sectionMeta.icon, label: sectionMeta.label, items: [] });
    }
    sections.get(sectionKey)!.items.push(item);
  }

  return {
    updatedAt: new Date().toISOString(),
    favorites,
    sections: Array.from(sections.values()),
  };
}

export default async (req: Request, context: Context) => {
  try {
    if (!cache || Date.now() - (cache.at as number) > TTL_MS) {
      cache = { at: Date.now(), data: await loadData() };
    }
    return new Response(JSON.stringify(cache.data), {
      headers: {
        "content-type": "application/json; charset=utf-8",
        "cache-control": "public, max-age=60",
      },
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: String(err) }), {
      status: 500,
      headers: { "content-type": "application/json; charset=utf-8" },
    });
  }
};

export const config: Config = {
  path: "/api/links",
};
