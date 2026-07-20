import type { Context, Config } from "@netlify/functions";
import * as XLSX from "xlsx";
import { unzipSync, strFromU8 } from "fflate";

const SHEET_ID = Netlify.env.get("SHEET_ID") || "1UJdKmCo94XOlFSIqcFFLY3SMBOa322t92GkCJg08X6g";
const EXPORT_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=xlsx`;

const CATEGORY_MAP: Record<string, { icon: string; label: string }> = {
    "운영": { icon: "🗂️", label: "운영" },
  "운영 25년 기존": { icon: "🗂️", label: "운영" },
  "운영 참고": { icon: "🗂️", label: "운영" },
  "운영 25년 참고": { icon: "🗂️", label: "운영" },
  "운영 업로드": { icon: "🗂️", label: "운영" },
  "공유 지사본": { icon: "🗂️", label: "운영" },
  "기획": { icon: "📝", label: "기획" },
  "기획 이벤트": { icon: "📝", label: "기획" },
  "기획 시상": { icon: "📝", label: "기획" },
  "기획 준비": { icon: "📝", label: "기획" },
  "기획 사업본부": { icon: "📝", label: "기획" },
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
  "민원": { icon: "📮", label: "민원" },
};
const FALLBACK_SECTION = { icon: "📌", label: "기타" };
const ENV_SECTION = { icon: "🔐", label: "운영·개발 환경" };

type Cell = { text: string; url: string | null; hasOwnLink: boolean; bold: boolean };
type Item = { title: string; url: string; subtitle: string | null; extra: { label: string; url: string } | null; favorite: boolean; icon: string | null };
type Section = { key: string; icon: string; label: string; items: Item[] };

let cache: { at: number; data: unknown } | null = null;
const TTL_MS = 5 * 60 * 1000;

function normalizeCategory(text: string) {
  return text.trim().replace(/_/g, " ").replace(/\s+/g, " ");
}

const looksLikeUrl = (t: string) => /^https?:\/\//i.test(t);

// sheetjs resolves hyperlink targets from the raw .rels XML without decoding
// entities, so "&amp;" in a target survives verbatim unless unescaped here.
function unescapeXml(s: string) {
  return s
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function cellInfo(raw: XLSX.CellObject | undefined, bold: boolean): Cell {
  const text = (raw?.v ?? "").toString().trim();
  const link = (raw as any)?.l?.Target as string | undefined;
  if (link) return { text, url: unescapeXml(link), hasOwnLink: true, bold };
  if (looksLikeUrl(text)) return { text, url: text, hasOwnLink: false, bold };
  return { text, url: null, hasOwnLink: false, bold };
}

// SheetJS's public build doesn't surface font weight when reading styles, so
// bold-cell detection (used to mark "favorite" rows) is done by hand against
// the raw XLSX XML instead of relying on ws[ref].s.
function getBoldRefs(buf: Buffer): Set<string> {
  try {
    const files = unzipSync(new Uint8Array(buf));
    const dec = (name: string) => (files[name] ? strFromU8(files[name]) : "");

    const stylesXml = dec("xl/styles.xml");
    const fonts = stylesXml.match(/<font>[\s\S]*?<\/font>/g) || [];
    const boldFontIds = new Set<number>();
    fonts.forEach((f, i) => {
      if (/<b\s*\/?>/.test(f)) boldFontIds.add(i);
    });

    const cellXfsBlock = stylesXml.match(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/);
    const xfs = cellXfsBlock ? cellXfsBlock[1].match(/<xf\b[^>]*\/?>/g) || [] : [];
    const boldXfIds = new Set<number>();
    xfs.forEach((xf, i) => {
      const m = xf.match(/fontId="(\d+)"/);
      if (m && boldFontIds.has(Number(m[1]))) boldXfIds.add(i);
    });

    const workbookXml = dec("xl/workbook.xml");
    const sheetEls = workbookXml.match(/<sheet\b[^>]*\/>/g) || [];
    let targetRid: string | null = null;
    for (const s of sheetEls) {
      if (!/state="hidden"/.test(s)) {
        const m = s.match(/r:id="(rId\d+)"/);
        if (m) {
          targetRid = m[1];
          break;
        }
      }
    }
    if (!targetRid) return new Set();

    const workbookRels = dec("xl/_rels/workbook.xml.rels");
    const relMatch = workbookRels.match(new RegExp(`<Relationship[^>]*Id="${targetRid}"[^>]*Target="([^"]+)"`));
    if (!relMatch) return new Set();
    const sheetXml = dec(`xl/${relMatch[1]}`);
    if (!sheetXml) return new Set();

    const boldRefs = new Set<string>();
    const cellOpenTags = sheetXml.match(/<c\s[^>]*>/g) || [];
    for (const tag of cellOpenTags) {
      const refM = tag.match(/\br="([A-Z]+\d+)"/);
      const sM = tag.match(/\bs="(\d+)"/);
      if (refM && sM && boldXfIds.has(Number(sM[1]))) boldRefs.add(refM[1]);
    }
    return boldRefs;
  } catch {
    return new Set();
  }
}

async function loadData() {
  const res = await fetch(EXPORT_URL);
  if (!res.ok) throw new Error(`sheet export failed: ${res.status}`);
  const buf = Buffer.from(await res.arrayBuffer());
  const wb = XLSX.read(buf, { type: "buffer" });
  const boldRefs = getBoldRefs(buf);

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
    const cols = [0, 1, 2, 3].map((c) => {
      const ref = XLSX.utils.encode_cell({ r, c });
      return cellInfo(ws[ref], boldRefs.has(ref));
    });
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
    let titleBold = false;
    // some rows have a real hyperlink whose display text is just the raw
    // URL itself (pasted link auto-linked to itself) - that's not a usable
    // label, so fall through to the pool search below in that case too.
    if (primary.hasOwnLink && !looksLikeUrl(primary.text)) {
      title = primary.text;
      titleBold = primary.bold;
    } else {
      const titleCell = pool.find((cell) => cell.text && !cell.url);
      title = titleCell ? titleCell.text : new URL(primary.url!).hostname;
      titleBold = titleCell ? titleCell.bold : false;
      if (titleCell) pool.splice(pool.indexOf(titleCell), 1);
    }
    // clean up leftover "- 어드민: " style dashes/colons from env sub-item rows
    title = title.replace(/^-\s*/, "").replace(/:\s*$/, "").trim() || title;

    // a leading emoji typed into the sheet becomes the card's custom icon
    const emojiMatch = title.match(/^(\p{Extended_Pictographic}️?)\s*/u);
    let icon: string | null = null;
    if (emojiMatch) {
      icon = emojiMatch[1];
      title = title.slice(emojiMatch[0].length).trim();
    }

    const subtitleCell = pool.find((cell) => cell.text && !cell.url);
    let subtitle = subtitleCell ? subtitleCell.text : null;
    if (subtitleCell) pool.splice(pool.indexOf(subtitleCell), 1);

    const extraCell = pool.find((cell) => cell.url);
    const extra = extraCell
      ? { label: extraCell.hasOwnLink && !looksLikeUrl(extraCell.text) ? extraCell.text : "관련 링크", url: extraCell.url! }
      : null;

    // a bolded title cell marks the row as a favorite; a leading "★" is
    // kept as a manual override for rows that can't easily be bolded
    const favorite = titleBold || title.startsWith("★");
    if (favorite) title = title.replace(/^★\s*/, "");

    const sectionMeta = catInfo ?? pendingSection ?? FALLBACK_SECTION;
    if (sectionMeta === ENV_SECTION && pendingEnvTag && !subtitle) {
      subtitle = `${pendingEnvTag} 환경`;
    }

    const item: Item = { title, url: primary.url!, subtitle, extra, favorite, icon };
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
