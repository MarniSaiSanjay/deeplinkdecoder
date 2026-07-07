import { useState } from "react";
import { decodeDeepLink, type DecodedLink, type KeyValue } from "./decoder";
import { theme } from "./theme";

const SAMPLE =
  "https://m365.cloud.microsoft/chat/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429/eyJpZCI6IlRXbGpjbTl6YjJaMFZqSjhhSFIwY0hNNkx5OXpkV0p6ZEhKaGRHVXViMlptYVdObExtTnZiWHcxT0RBeU9EVTRNaTB6WlRabExUUmpZVE10T0RFeE1TMHhNR00yTkRKaE0yVmxaV0o4TlRsaVlqaGlabVF0TWpnM1lpMDBaVGMwTFdFNFpHSXRZelprWVdRM1lUWXlPREF5ZkdWdUxYVnoiLCJzY2VuYXJpbyI6InNoYXJlTGluayIsImNoYXRUeXBlIjoid2ViIiwicHJvcGVydGllcyI6eyJwcm9tcHRTb3VyY2UiOiJtaWNyb3NvZnQiLCJjbGlja1RpbWVzdGFtcCI6IjIwMjYtMDctMDFUMTU6NTE6NDkuMzQwWiJ9LCJyZWZlcnJhbCI6eyJjbW1pZCI6ImNtbXhoeHA3YmE3In0sImNvcnJlbGF0aW9uSWQiOiJlNDhkOTRlMi1iNDcwLTQ3NDYtYmE0Mi04NDE0NTI2MDc1MjUiLCJ2ZXJzaW9uIjoxLjF9?fromcode=cmmxhxp7ba7&ct=starter_jun_mca&ocid=starter_jun_mca&utm_campaign=starter_jun_mca&utm_source=starter_jun_mca";

function Card(props: { title: string; children: React.ReactNode }) {
  return (
    <section
      style={{
        background: theme.cardBg,
        border: `1px solid ${theme.cardBorder}`,
        borderRadius: 12,
        boxShadow: theme.cardShadow,
        padding: "18px 20px",
        marginBottom: 16,
      }}
    >
      <h2
        style={{
          margin: "0 0 12px",
          fontSize: 14,
          fontWeight: 700,
          letterSpacing: 0.3,
          textTransform: "uppercase",
          color: theme.textSoft,
        }}
      >
        {props.title}
      </h2>
      {props.children}
    </section>
  );
}

function Row(props: { label: string; value?: string; mono?: boolean }) {
  if (props.value === undefined || props.value === "") return null;
  return (
    <div
      style={{
        display: "grid",
        gridTemplateColumns: "160px 1fr",
        gap: 12,
        padding: "7px 0",
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
  return (
    <>
      <Card title="Overview">
        <Row label="Link type" value={d.linkKind} />
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
        <Card title="Decoded payload (JSON)">
          <pre
            style={{
              margin: 0,
              padding: 16,
              borderRadius: 8,
              background: theme.codeBg,
              color: theme.codeText,
              fontFamily: theme.monoStack,
              fontSize: 12.5,
              lineHeight: 1.5,
              overflowX: "auto",
            }}
          >
            {d.payloadJson}
          </pre>
        </Card>
      )}

      {d.notes.length > 0 && (
        <Card title="Notes">
          <ul style={{ margin: 0, paddingLeft: 18, color: theme.textSoft, fontSize: 13, lineHeight: 1.7 }}>
            {d.notes.map((n, i) => (
              <li key={i}>{n}</li>
            ))}
          </ul>
        </Card>
      )}
    </>
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
      <header
        style={{
          background: theme.heroGradient,
          color: "#fff",
          padding: "40px 24px 44px",
        }}
      >
        <div style={{ maxWidth: 900, margin: "0 auto" }}>
          <h1 style={{ margin: 0, fontSize: 30, fontWeight: 700 }}>Deep Link Decoder</h1>
          <p style={{ margin: "8px 0 0", fontSize: 15, opacity: 0.92, maxWidth: 640 }}>
            Paste a Microsoft 365 Copilot or Word deep link to break it down into its payload,
            encoding, prompt ID, locale and query parameters.
          </p>
        </div>
      </header>

      <main style={{ maxWidth: 900, margin: "0 auto", padding: "24px 24px 64px", marginTop: -20 }}>
        <Card title="Link">
          <textarea
            value={input}
            onChange={(e) => setInput(e.target.value)}
            placeholder="https://m365.cloud.microsoft/chat/entity1-.../<encoded payload>?..."
            spellCheck={false}
            rows={5}
            style={{
              width: "100%",
              resize: "vertical",
              padding: 12,
              borderRadius: 8,
              border: `1px solid ${theme.fieldBorder}`,
              fontFamily: theme.monoStack,
              fontSize: 12.5,
              lineHeight: 1.5,
              color: theme.text,
              background: theme.fieldNestBg,
            }}
          />
          <div style={{ display: "flex", gap: 10, marginTop: 12, flexWrap: "wrap" }}>
            <button
              type="button"
              onClick={() => handleDecode(input)}
              style={{ ...btnStyle, background: theme.primary, color: "#fff" }}
            >
              Decode
            </button>
            <button
              type="button"
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
      </main>

      <footer style={{ textAlign: "center", padding: "0 24px 32px", color: theme.subtle, fontSize: 12 }}>
        Decoding runs entirely in your browser — nothing is sent to any server.
      </footer>
    </div>
  );
}
