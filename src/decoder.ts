// Core decoding logic for Microsoft 365 Copilot / Word deep links.
//
// This is the inverse of the link generator. A generated link is one of:
//   1. M365 chat:  https://m365.cloud.microsoft/chat/entity1-<entityId>/<base64Payload>?<query>
//   2. Word handoff: https://word.cloud.microsoft/handoff/growth/?handinPayload=<base64Payload>&<query>
//   3. M365 base:  https://m365.cloud.microsoft/?<query>            (no payload)
//   4. Create / store / researcher links: mostly query-only.
//
// The base64 payload is JSON. Its `id` field carries the substrate identifier,
// which itself may be:
//   - url-encoded base64 of "MicrosoftV2|https://substrate.office.com|<uuid>|<promptId>|<locale>"
//   - base64 of "MicrosoftHandOffV5|<promptId>|<locale>"           (Researcher-with-prompt)
//   - a raw "<promptId>_<locale>" string                            (Word v1.2 handoff)

export interface KeyValue {
  key: string;
  value: string;
}

export interface SubstrateIdentifier {
  /** Raw decoded identifier string. */
  raw: string;
  /** Encoding scheme, e.g. "MicrosoftV2", "MicrosoftHandOffV5", or "v1.2 (promptId_locale)". */
  encoding: string;
  endpoint?: string;
  /** Random correlation GUID present only in MicrosoftV2 identifiers. */
  linkGuid?: string;
  promptId?: string;
  locale?: string;
}

export interface DecodedLink {
  /** The kind of link detected. */
  linkKind: string;
  host: string;
  path: string;
  /** Host family: "M365", "Word", "UDL", "App Store", or "Other". */
  surface: string;
  /** UDL delivery mode ("app" or "mweb"), when this is a Unified Deep Link. */
  udlMode?: string;
  /** UDL target app (e.g. "copilot"), when this is a Unified Deep Link. */
  udlApp?: string;
  /** Entity GUID from the /chat/entity1-<guid>/ route, when present. */
  entityId?: string;
  /** Parsed substrate identifier from payload.id. */
  substrate?: SubstrateIdentifier;
  /** Pretty-printed JSON of the payload. */
  payloadJson?: string;
  /** Query-string parameters. */
  queryParams: KeyValue[];
  /** Non-fatal notes surfaced to the user. */
  notes: string[];
}

/** Decode a base64 string (standard or URL-safe, padded or not) to UTF-8. */
function base64ToUtf8(input: string): string {
  let b64 = input.replace(/-/g, "+").replace(/_/g, "/");
  const pad = b64.length % 4;
  if (pad) b64 += "=".repeat(4 - pad);
  const binary = atob(b64);
  const bytes = Uint8Array.from(binary, (c) => c.charCodeAt(0));
  return new TextDecoder().decode(bytes);
}

/** Heuristic: does this string look like base64 (vs. a raw promptId_locale)? */
function looksLikeBase64(s: string): boolean {
  return /^[A-Za-z0-9+/\-_]+={0,2}$/.test(s) && s.length % 4 <= 3 && s.length >= 8 && !s.includes("_");
}

/** Parse a substrate identifier string into its parts. */
function parseSubstrateIdentifier(idField: string): SubstrateIdentifier {
  // The generator stores payload.id as encodeURIComponent(base64(...)) for V2/HandOff,
  // or as a raw "<promptId>_<locale>" for Word v1.2.
  let decodedOnce = idField;
  try {
    decodedOnce = decodeURIComponent(idField);
  } catch {
    // Not url-encoded; use as-is.
  }

  // Word v1.2 raw form: "<promptId>_<locale>" (contains an underscore, not base64).
  if (decodedOnce.includes("_") && !decodedOnce.includes("|") && !looksLikeBase64(decodedOnce)) {
    const idx = decodedOnce.lastIndexOf("_");
    return {
      raw: decodedOnce,
      encoding: "v1.2 (promptId_locale)",
      promptId: decodedOnce.slice(0, idx),
      locale: decodedOnce.slice(idx + 1),
    };
  }

  // Otherwise it should be base64 of a pipe-delimited identifier.
  let inner = decodedOnce;
  try {
    inner = base64ToUtf8(decodedOnce);
  } catch {
    // Leave inner as-is; parsing below will just report the raw value.
  }

  const parts = inner.split("|");
  const encoding = parts[0] || "unknown";

  if (encoding === "MicrosoftV2") {
    // MicrosoftV2 | endpoint | <uuid> | <promptId> | <locale>
    return {
      raw: inner,
      encoding,
      endpoint: parts[1],
      linkGuid: parts[2],
      promptId: parts[3],
      locale: parts[4],
    };
  }

  if (encoding === "MicrosoftHandOffV5") {
    // MicrosoftHandOffV5 | <promptId> | <locale>
    return {
      raw: inner,
      encoding,
      promptId: parts[1],
      locale: parts[2],
    };
  }

  return { raw: inner, encoding };
}

/** Parse a query string (without the leading '?') into ordered key/value pairs. */
function parseQuery(search: string): KeyValue[] {
  const out: KeyValue[] = [];
  if (!search) return out;
  const params = new URLSearchParams(search);
  params.forEach((value, key) => out.push({ key, value }));
  return out;
}

const UDL_HOST = "unifiedlink.svc.cloud.microsoft";
const M365_HOST = "m365.cloud.microsoft";
const WORD_HOST = "word.cloud.microsoft";

interface HostInfo {
  surface: string;
  /** Path with any UDL "/{mode}/{app}" prefix stripped, so downstream matching is host-agnostic. */
  effectivePath: string;
  udlMode?: string;
  udlApp?: string;
}

/**
 * Classify the host and, for UDL links, peel off the "/{mode}/{app}" prefix so
 * the shared chat/handoff/base matchers work across M365, Word and UDL.
 */
function classifyHost(url: URL): HostInfo {
  const host = url.host;
  if (host.includes(UDL_HOST)) {
    // UDL path shape: /{mode}/{app}/<rest>, e.g. /app/copilot/chat/...
    const m = url.pathname.match(/^\/(app|mweb)\/([^/]+)(\/.*)?$/);
    if (m) {
      return { surface: "UDL", udlMode: m[1], udlApp: m[2], effectivePath: m[3] || "/" };
    }
    return { surface: "UDL", effectivePath: url.pathname };
  }
  if (host.includes(WORD_HOST)) return { surface: "Word", effectivePath: url.pathname };
  if (host.includes(M365_HOST)) return { surface: "M365", effectivePath: url.pathname };
  if (host.includes("apps.apple.com")) return { surface: "App Store", effectivePath: url.pathname };
  if (host.includes("play.google.com")) return { surface: "App Store", effectivePath: url.pathname };
  return { surface: "Other", effectivePath: url.pathname };
}

/**
 * Decode a Microsoft 365 / Word / UDL deep link. Throws on input that is not a URL.
 */
export function decodeDeepLink(rawUrl: string): DecodedLink {
  const trimmed = rawUrl.trim();
  if (!trimmed) {
    throw new Error("Enter a link to decode.");
  }

  let url: URL;
  try {
    url = new URL(trimmed);
  } catch {
    throw new Error("That doesn't look like a valid URL.");
  }

  const notes: string[] = [];
  const queryParams = parseQuery(url.search.replace(/^\?/, ""));
  const { surface, effectivePath, udlMode, udlApp } = classifyHost(url);
  const result: DecodedLink = {
    linkKind: "Unknown",
    host: url.host,
    path: url.pathname,
    surface,
    udlMode,
    udlApp,
    queryParams,
    notes,
  };
  // Prefix used in link-kind labels so the surface is obvious.
  const label = surface === "UDL" ? `UDL${udlApp ? ` (${udlApp})` : ""}` : surface;

  // Word handoff link — payload is in the handinPayload query param.
  if (effectivePath.includes("/handoff/growth")) {
    result.linkKind = `${label} handoff (growth)`;
    // The handinPayload param is the payload itself; show it decoded, not raw.
    result.queryParams = queryParams.filter((kv) => kv.key !== "handinPayload");
    const handin = url.searchParams.get("handinPayload");
    if (handin) {
      try {
        const json = base64ToUtf8(decodeURIComponent(handin));
        const payload = JSON.parse(json) as Record<string, unknown>;
        result.payloadJson = JSON.stringify(payload, null, 2);
        if (typeof payload.id === "string") {
          result.substrate = parseSubstrateIdentifier(payload.id);
        }
      } catch {
        notes.push("Could not decode the handinPayload parameter as base64 JSON.");
      }
    } else {
      notes.push("No handinPayload parameter found on this handoff link.");
    }
    return result;
  }

  // Chat link with an encoded payload in the path.
  const chatMatch = effectivePath.match(/\/chat\/entity1-([0-9a-fA-F-]+)\/([^/]+)$/);
  if (chatMatch) {
    result.linkKind = `${label} Copilot chat (prompt)`;
    result.entityId = chatMatch[1];
    const encodedPayload = chatMatch[2];
    try {
      const json = base64ToUtf8(decodeURIComponent(encodedPayload));
      const payload = JSON.parse(json) as Record<string, unknown>;
      result.payloadJson = JSON.stringify(payload, null, 2);
      if (payload.gptInfo && typeof payload.gptInfo === "object") {
        result.linkKind = `${label} Copilot chat — Researcher with prompt`;
      }
      if (typeof payload.id === "string") {
        result.substrate = parseSubstrateIdentifier(payload.id);
      }
    } catch {
      notes.push("Could not decode the path payload as base64 JSON.");
    }
    return result;
  }

  // Researcher link (no payload, titleId in query).
  if (effectivePath.endsWith("/chat") || effectivePath.endsWith("/chat/")) {
    result.linkKind = `${label} Copilot chat (Researcher entry point)`;
    return result;
  }

  if (effectivePath.endsWith("/create")) {
    result.linkKind = `${label} Copilot create`;
    return result;
  }

  if (surface === "App Store") {
    result.linkKind = url.host.includes("apps.apple.com")
      ? "iOS App Store link"
      : "Android Play Store link";
    return result;
  }

  // Base / marketing link — everything is in the query string.
  if (surface === "M365" || surface === "UDL") {
    result.linkKind = `${label} base link`;
    // The referrer param nests utm_* pairs; surface them.
    const referrer = url.searchParams.get("referrer");
    if (referrer) {
      parseQuery(referrer).forEach((kv) =>
        notes.push(`referrer → ${kv.key} = ${kv.value}`),
      );
    }
    return result;
  }

  return result;
}
