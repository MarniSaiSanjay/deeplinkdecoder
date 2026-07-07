import { useState } from "react";
import { decodeDeepLink, type DecodedLink, type KeyValue } from "./decoder";
import { theme, surfaceColors } from "./theme";

const SAMPLE =
  "https://m365.cloud.microsoft/chat/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429/eyJpZCI6IlRXbGpjbTl6YjJaMFZqSjhhSFIwY0hNNkx5OXpkV0p6ZEhKaGRHVXViMlptYVdObExtTnZiWHcxT0RBeU9EVTRNaTB6WlRabExUUmpZVE10T0RFeE1TMHhNR00yTkRKaE0yVmxaV0o4TlRsaVlqaGlabVF0TWpnM1lpMDBaVGMwTFdFNFpHSXRZelprWVdRM1lUWXlPREF5ZkdWdUxYVnoiLCJzY2VuYXJpbyI6InNoYXJlTGluayIsImNoYXRUeXBlIjoid2ViIiwicHJvcGVydGllcyI6eyJwcm9tcHRTb3VyY2UiOiJtaWNyb3NvZnQiLCJjbGlja1RpbWVzdGFtcCI6IjIwMjYtMDctMDFUMTU6NTE6NDkuMzQwWiJ9LCJyZWZlcnJhbCI6eyJjbW1pZCI6ImNtbXhoeHA3YmE3In0sImNvcnJlbGF0aW9uSWQiOiJlNDhkOTRlMi1iNDcwLTQ3NDYtYmE0Mi04NDE0NTI2MDc1MjUiLCJ2ZXJzaW9uIjoxLjF9?fromcode=cmmxhxp7ba7&ct=starter_jun_mca&ocid=starter_jun_mca&utm_campaign=starter_jun_mca&utm_source=starter_jun_mca";

function Card(props: { title: string; children: React.ReactNode; accent?: string; wide?: boolean; action?: React.ReactNode }) {
  return (
    <section
      className="dl-card"
      style={{
        background: theme.cardBg,
        border: `1px solid ${theme.cardBorder}`,
        borderRadius: 14,
        boxShadow: theme.cardShadow,
        padding: "18px 20px",
        marginBottom: 0,
        gridColumn: props.wide ? "1 / -1" : undefined,
        position: "relative",
        overflow: "hidden",
      }}
    >
      <span
        aria-hidden
        style={{
          position: "absolute",
          left: 0,
          top: 0,
          bottom: 0,
          width: 4,
          background: props.accent ?? theme.accentGradient,
        }}
      />
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, margin: "0 0 12px" }}>
        <h2
          style={{
            margin: 0,
            fontSize: 12.5,
            fontWeight: 700,
            letterSpacing: 0.6,
            textTransform: "uppercase",
            color: theme.textSoft,
          }}
        >
          {props.title}
        </h2>
        {props.action}
      </div>
      {props.children}
    </section>
  );
}

function Badge(props: { surface: string }) {
  const c = surfaceColors[props.surface] ?? surfaceColors.Other;
  return (
    <span
      style={{
        display: "inline-flex",
        alignItems: "center",
        gap: 6,
        background: c.bg,
        color: c.fg,
        borderRadius: 999,
        padding: "4px 11px",
        fontSize: 12,
        fontWeight: 700,
        letterSpacing: 0.2,
      }}
    >
      <span style={{ width: 7, height: 7, borderRadius: 999, background: c.dot }} />
      {props.surface}
    </span>
  );
}

function CopyBtn(props: { value: string; solid?: boolean }) {
  const [copied, setCopied] = useState(false);
  return (
    <button
      type="button"
      className={props.solid ? "dl-copy-solid" : "dl-copy"}
      title="Copy"
      onClick={() => {
        navigator.clipboard?.writeText(props.value).then(
          () => {
            setCopied(true);
            setTimeout(() => setCopied(false), 1100);
          },
          () => undefined,
        );
      }}
      style={{
        border: props.solid ? `1px solid ${theme.fieldBorder}` : "none",
        background: props.solid ? theme.pillBg : "transparent",
        color: copied ? theme.primary : props.solid ? theme.pillText : theme.subtle,
        cursor: "pointer",
        fontSize: 11,
        fontWeight: 600,
        padding: props.solid ? "4px 10px" : "2px 6px",
        borderRadius: 6,
        flexShrink: 0,
      }}
    >
      {copied ? "Copied" : "Copy"}
    </button>
  );
}

function Row(props: { label: string; value?: string; mono?: boolean }) {
  if (props.value === undefined || props.value === "") return null;
  return (
    <div
      className="dl-row"
      style={{
        display: "grid",
        gridTemplateColumns: "120px minmax(0, 1fr) auto",
        gap: 12,
        padding: "7px 8px",
        margin: "0 -8px",
        borderRadius: 8,
        borderBottom: `1px solid ${theme.cardBorderLight}`,
        alignItems: "baseline",
      }}
    >
      <div style={{ color: theme.textSoft, fontSize: 13, fontWeight: 600 }}>{props.label}</div>
      <div
        style={{
          fontSize: 13,
          wordBreak: "break-all",
          fontFamily: props.mono ? theme.monoStack : theme.fontStack,
          color: theme.text,
        }}
      >
        {props.value}
      </div>
      {props.mono ? <CopyBtn value={props.value} /> : <span />}
    </div>
  );
}

function KeyValueTable(props: { rows: KeyValue[] }) {
  if (props.rows.length === 0) {
    return <div style={{ color: theme.subtle, fontSize: 13 }}>None.</div>;
  }
  return (
    <div>
      {props.rows.map((kv, i) => (
        <Row key={`${kv.key}-${i}`} label={kv.key} value={kv.value} mono />
      ))}
    </div>
  );
}

function Results(props: { decoded: DecodedLink }) {
  const d = props.decoded;
  const summary = d.substrate
    ? [d.substrate.promptId && `prompt ${d.substrate.promptId}`, d.substrate.locale && `locale ${d.substrate.locale}`]
        .filter(Boolean)
        .join(" · ")
    : "";
  return (
    <div style={{ animation: "cardIn 0.35s ease both" }}>
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 12,
          flexWrap: "wrap",
          background: theme.cardBg,
          border: `1px solid ${theme.cardBorder}`,
          borderRadius: 14,
          boxShadow: theme.cardShadow,
          padding: "16px 20px",
          marginBottom: 16,
        }}
      >
        <Badge surface={d.surface} />
        <div style={{ fontSize: 16, fontWeight: 700, color: theme.text }}>{d.linkKind}</div>
        {summary && (
          <div style={{ marginLeft: "auto", fontSize: 12.5, color: theme.textSoft, fontFamily: theme.monoStack }}>
            {summary}
          </div>
        )}
      </div>

      <div className="results-grid">
        <Card title="Overview" accent={surfaceColors[d.surface]?.dot} wide>
          <Row label="Link type" value={d.linkKind} />
          <Row label="Surface" value={d.surface} />
          <Row label="UDL mode" value={d.udlMode} mono />
          <Row label="UDL app" value={d.udlApp} mono />
          <Row label="Host" value={d.host} mono />
          <Row label="Path" value={d.path} mono />
          <Row label="Entity ID" value={d.entityId} mono />
        </Card>

        {d.substrate && (
          <Card title="Substrate identifier (payload.id)">
            <Row label="Encoding" value={d.substrate.encoding} />
            <Row label="Endpoint" value={d.substrate.endpoint} mono />
            <Row label="Prompt ID" value={d.substrate.promptId} mono />
            <Row label="Locale" value={d.substrate.locale} mono />
            <Row label="Link GUID" value={d.substrate.linkGuid} mono />
            <Row label="Raw" value={d.substrate.raw} mono />
          </Card>
        )}

        <Card title="Query parameters">
          <KeyValueTable rows={d.queryParams} />
        </Card>

        {d.payloadJson && (
          <Card
            title="Decoded payload (JSON)"
            wide
            action={<CopyBtn value={d.payloadJson} solid />}
          >
            <pre
              style={{
                margin: 0,
                padding: 16,
                borderRadius: 10,
                background: theme.codeBg,
                color: theme.codeText,
                fontFamily: theme.monoStack,
                fontSize: 12.5,
                lineHeight: 1.55,
                overflowX: "auto",
              }}
            >
              {d.payloadJson}
            </pre>
          </Card>
        )}

        {d.notes.length > 0 && (
          <Card title="Notes" accent="#e6a700">
            <ul style={{ margin: 0, paddingLeft: 18, color: theme.textSoft, fontSize: 13, lineHeight: 1.7 }}>
              {d.notes.map((n, i) => (
                <li key={i}>{n}</li>
              ))}
            </ul>
          </Card>
        )}
      </div>
    </div>
  );
}

export default function App() {
  const [input, setInput] = useState("");
  const [decoded, setDecoded] = useState<DecodedLink | null>(null);
  const [error, setError] = useState<string | null>(null);

  function handleDecode(value: string) {
    setError(null);
    setDecoded(null);
    try {
      setDecoded(decodeDeepLink(value));
    } catch (e) {
      setError(e instanceof Error ? e.message : "Failed to decode the link.");
    }
  }

  const hasResults = !!(decoded || error);

  const btnStyle: React.CSSProperties = {
    padding: "10px 20px",
    borderRadius: 8,
    border: "none",
    fontSize: 14,
    fontWeight: 600,
    cursor: "pointer",
  };

  return (
    <div style={{ minHeight: "100vh" }}>
      <style>{`
        .decoder-layout {
          display: grid;
          grid-template-columns: minmax(360px, 460px) minmax(0, 1fr);
          gap: 24px;
          align-items: start;
          transition: grid-template-columns 0.3s ease;
        }
        .decoder-layout[data-mode="solo"] {
          display: block;
          max-width: 660px;
          margin: 0 auto;
        }
        .decoder-layout[data-mode="solo"] .decoder-input { position: static; }
        .decoder-layout > * { min-width: 0; }
        .decoder-input { position: sticky; top: 20px; }
        .results-grid {
          display: grid;
          grid-template-columns: repeat(2, minmax(0, 1fr));
          gap: 16px;
          align-items: start;
        }
        .results-grid > section { max-width: 100%; }

        .dl-card { transition: box-shadow 0.18s ease, transform 0.18s ease; }
        .dl-card:hover { box-shadow: ${theme.cardShadowLg}; transform: translateY(-2px); }

        .dl-row { transition: background 0.12s ease; }
        .dl-row:hover { background: ${theme.fieldNestBg}; }
        .dl-copy { opacity: 0; transition: opacity 0.12s ease; }
        .dl-row:hover .dl-copy { opacity: 1; }
        .dl-copy:hover { background: ${theme.pillBg} !important; }

        .dl-textarea:focus {
          outline: none;
          border-color: ${theme.primary} !important;
          box-shadow: 0 0 0 3px rgba(0,120,212,0.15);
        }

        .dl-btn { transition: transform 0.1s ease, box-shadow 0.15s ease, filter 0.15s ease; }
        .dl-btn:hover { filter: brightness(1.05); }
        .dl-btn:active { transform: translateY(1px); }
        .dl-btn-primary:hover { box-shadow: 0 4px 14px rgba(0,120,212,0.35); }

        @media (max-width: 1080px) {
          .decoder-layout { grid-template-columns: 1fr; }
          .decoder-input { position: static; }
          .results-grid { grid-template-columns: 1fr; }
        }
      `}</style>
      <header
        style={{
          background: theme.heroGradient,
          color: "#fff",
          padding: "44px 24px 52px",
          position: "relative",
          overflow: "hidden",
        }}
      >
        <div style={{ maxWidth: 900, margin: "0 auto", position: "relative", padding: "0 32px", textAlign: "center" }}>
          <div
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: 8,
              background: "rgba(255,255,255,0.16)",
              border: "1px solid rgba(255,255,255,0.25)",
              borderRadius: 999,
              padding: "5px 13px",
              fontSize: 12,
              fontWeight: 600,
              marginBottom: 14,
            }}
          >
            <span style={{ width: 7, height: 7, borderRadius: 999, background: "#7ee0a1" }} />
            100% client-side · nothing leaves your browser
          </div>
          <h1
            style={{
              margin: 0,
              fontSize: 34,
              fontWeight: 800,
              letterSpacing: -0.5,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              gap: 14,
            }}
          >
            <img
              src={`${import.meta.env.BASE_URL}icon.png`}
              alt="Deep Link Decoder logo"
              style={{ height: 52, width: "auto", filter: "drop-shadow(0 2px 6px rgba(0,0,0,0.25))" }}
            />
            Deep Link Decoder
          </h1>
          <p style={{ margin: "10px auto 0", fontSize: 15.5, opacity: 0.94, maxWidth: 640, lineHeight: 1.55 }}>
            Paste a Microsoft 365 Copilot, Word or Unified Deep Link (UDL) to break it down into its
            payload, encoding, prompt ID, locale and query parameters.
          </p>
        </div>
      </header>

      <main style={{ maxWidth: 1520, margin: "0 auto", padding: "24px 32px 64px", marginTop: -20 }}>
        <div className="decoder-layout" data-mode={hasResults ? "split" : "solo"}>
          <div className="decoder-input">
            <Card title="Link">
              <textarea
                value={input}
                onChange={(e) => setInput(e.target.value)}
                placeholder="https://m365.cloud.microsoft/chat/entity1-.../<encoded payload>?..."
                spellCheck={false}
                rows={8}
                className="dl-textarea"
                style={{
                  width: "100%",
                  boxSizing: "border-box",
                  resize: "vertical",
                  padding: 12,
                  borderRadius: 8,
                  border: `1px solid ${theme.fieldBorder}`,
                  fontFamily: theme.monoStack,
                  fontSize: 12.5,
                  lineHeight: 1.5,
                  color: theme.text,
                  background: theme.fieldNestBg,
                  transition: "border-color 0.15s ease, box-shadow 0.15s ease",
                }}
              />
              <div style={{ display: "flex", gap: 10, marginTop: 12, flexWrap: "wrap" }}>
                <button
                  type="button"
                  className="dl-btn dl-btn-primary"
                  onClick={() => handleDecode(input)}
                  style={{ ...btnStyle, background: theme.accentGradient, color: "#fff" }}
                >
                  Decode
                </button>
                <button
                  type="button"
                  className="dl-btn"
                  onClick={() => {
                    setInput(SAMPLE);
                    handleDecode(SAMPLE);
                  }}
                  style={{ ...btnStyle, background: theme.pillBg, color: theme.pillText }}
                >
                  Try a sample
                </button>
                <button
                  type="button"
                  className="dl-btn"
                  onClick={() => {
                    setInput("");
                    setDecoded(null);
                    setError(null);
                  }}
                  style={{ ...btnStyle, background: "transparent", color: theme.textSoft, border: `1px solid ${theme.fieldBorder}` }}
                >
                  Clear
                </button>
              </div>
            </Card>
          </div>

          {hasResults && (
            <div>
              {error && (
                <div
                  role="alert"
                  style={{
                    background: theme.dangerBg,
                    color: theme.danger,
                    border: `1px solid ${theme.danger}33`,
                    borderRadius: 10,
                    padding: "12px 16px",
                    fontSize: 13.5,
                    marginBottom: 16,
                  }}
                >
                  {error}
                </div>
              )}

              {decoded && <Results decoded={decoded} />}
            </div>
          )}
        </div>
      </main>

      <footer style={{ textAlign: "center", padding: "0 24px 32px", color: theme.subtle, fontSize: 12 }}>
        Decoding runs entirely in your browser — nothing is sent to any server.
      </footer>
    </div>
  );
}
