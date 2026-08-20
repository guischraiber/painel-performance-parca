import React, { useMemo, useState, useCallback } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  ResponsiveContainer,
  LineChart,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ReferenceLine,
} from "recharts";

// ---------------------------------------------------------------------------
// Paleta padrão (ver skill gestao-parca-rca)
// ---------------------------------------------------------------------------
const C = {
  laranja: "#F97316",
  laranjaLight: "#FED7AA",
  verde: "#16A34A",
  verdeLight: "#BBF7D0",
  vermelho: "#DC2626",
  vermelhoLight: "#FEE2E2",
  amarelo: "#CA8A04",
  amareloLight: "#FEF08A",
  azul: "#2563EB",
  azulLight: "#DBEAFE",
  cinzaFundo: "#F8F7F4",
  cinzaCard: "#FFFFFF",
  cinzaBorda: "#E5E3DF",
  cinzaTexto: "#6B7280",
  texto: "#1C1917",
};

const META_REVERSA = 0.18;
const META_DEVOLUCAO = 0.505;
const TODOS = "TODOS";

// Grupos de nome_armazem que representam o mesmo parça, mas aparecem com nomes diferentes
// na base (confirmado em conversa em 20/08/2026). A chave é o nome canônico usado no painel.
const CONSOLIDACAO_ARMAZEM = {
  "LOGME CWB": "LOGME",
  LOGME: "LOGME",
  REAL: "GRM MÓVEIS (Real)",
  "GRM MOV": "GRM MÓVEIS (Real)",
};

// nome_armazem que não representa um parça de fato — fica fora da análise, mas sinalizado.
const ARMAZEM_EXCLUIDO = "BOM/FORA DE LINHA";

// Bucket de devolução sem separação por parça (sem estoque próprio na devolução) — tratado
// como um "parça" à parte, conforme decisão do usuário em 20/08/2026.
const ARMAZEM_DEVOLUCAO_GENERICA = "DEVOLUÇÃO (41)";
const NOME_DEVOLUCAO_GENERICA = "Devolução 41";

// Extrai o nome do parça a partir da coluna nome_armazem, removendo o sufixo de tipo de
// operação (REVERSA/DEVOLUÇÃO, com ou sem hífen) e aplicando as consolidações acima.
function extrairParceiroDoArmazem(nomeArmazem) {
  const raw = String(nomeArmazem || "").trim();
  if (!raw) return { parceiro: null, excluido: true, motivo: "vazio" };

  const upper = raw.toUpperCase();
  if (upper.includes(ARMAZEM_EXCLUIDO)) {
    return { parceiro: null, excluido: true, motivo: ARMAZEM_EXCLUIDO };
  }
  if (upper.includes(ARMAZEM_DEVOLUCAO_GENERICA)) {
    return { parceiro: NOME_DEVOLUCAO_GENERICA, excluido: false };
  }

  const base = raw.replace(/\s*-?\s*(REVERSA|DEVOLU[ÇC][ÃA]O)\s*$/i, "").trim();
  const baseUpper = base.toUpperCase();
  const parceiro = CONSOLIDACAO_ARMAZEM[baseUpper] || base;
  return { parceiro, excluido: false };
}

// Nome final exibido no painel: padroniza para "<Nome> - Reversa" / "<Nome> - Devolução",
// a partir do nome base extraído do armazém e do tipo de operação (vindo de origem, mais
// confiável que o sufixo do próprio nome_armazem). O bucket genérico "Devolução 41" fica sem
// sufixo — decisão do usuário em 20/08/2026, já que não existe uma versão Reversa dele para
// desambiguar.
function nomeFinalParceiro(base, tipoOperacao) {
  if (base === NOME_DEVOLUCAO_GENERICA) return base;
  return `${base} - ${tipoOperacao}`;
}

// ---------------------------------------------------------------------------
// Helpers de formatação e parsing numérico
// ---------------------------------------------------------------------------
function toNum(v) {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return v;
  const s = String(v).trim().replace(/\./g, "").replace(",", ".");
  const n = parseFloat(s);
  return Number.isNaN(n) ? null : n;
}

// Detecta automaticamente o formato numérico: exportações do Looker vêm com ponto decimal
// "cru" (ex.: "467.95999999999998"), enquanto planilhas reformatadas em pt-BR vêm com vírgula
// decimal (ex.: "1006,42"). Sem essa detecção, tratar tudo como pt-BR gera valores absurdos.
function toNumSmart(v) {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return v;
  const s = String(v).trim();
  if (s.includes(",")) return toNum(s);
  const n = parseFloat(s.replace(",", "."));
  return Number.isNaN(n) ? null : n;
}

function fmtPct(x, digits = 1) {
  if (x === null || x === undefined || Number.isNaN(x)) return "—";
  return `${(x * 100).toFixed(digits)}%`;
}

function fmtMoney(x) {
  if (x === null || x === undefined || Number.isNaN(x)) return "—";
  return x.toLocaleString("pt-BR", { style: "currency", currency: "BRL", maximumFractionDigits: 0 });
}

// Lê texto de arquivo detectando csv (Papa) ou xlsx (SheetJS), devolvendo array de objetos.
function parseTabular(text, isXlsxWorkbook) {
  if (isXlsxWorkbook) {
    const ws = isXlsxWorkbook.Sheets[isXlsxWorkbook.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
  }
  const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
  return parsed.data;
}

function normalizeKey(s) {
  return String(s)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") // remove acentos
    .trim()
    .toLowerCase();
}

// Casa por nome exato primeiro; se não achar, tenta por versão normalizada (sem acento/caixa),
// para tolerar pequenas variações de cabeçalho (ex.: "Código" vs "codigo").
function getField(row, names) {
  for (const n of names) {
    if (row[n] !== undefined && row[n] !== null && row[n] !== "") return row[n];
  }
  const normalizedTargets = names.map(normalizeKey);
  for (const key of Object.keys(row)) {
    if (normalizedTargets.includes(normalizeKey(key)) && row[key] !== null && row[key] !== "") {
      return row[key];
    }
  }
  return null;
}

// ---------------------------------------------------------------------------
// Parsing da base Venda Salvados (Looker) — nota a nota
// Receita e custo já vêm faturados no arquivo (colunas total_gmv / total_custo) — não há
// taxa nem canal para aplicar nesta versão simplificada. O nome do parça vem da coluna
// nome_armazem (não depende mais de planilha de cadastro/código de cliente).
// ---------------------------------------------------------------------------
function parseVendaSalvados(rows) {
  const notas = [];
  const excluidosPorMotivo = {}; // motivo -> { qtd, receita }

  rows.forEach((r) => {
    const origem = String(getField(r, ["origem"]) || "");
    let tipo_operacao = null;
    if (origem.includes("Reversa")) tipo_operacao = "Reversa";
    else if (origem.includes("Devolução") || origem.includes("Devolucao")) tipo_operacao = "Devolução";
    else return;

    const receita = toNumSmart(getField(r, ["total_gmv"]));
    const custo = toNumSmart(getField(r, ["total_custo"]));
    if (receita === null || custo === null) return;

    const ano = getField(r, ["ano"]);
    const mes = getField(r, ["mes"]);
    if (!ano || !mes) return;
    const competencia = `${ano}-${String(mes).padStart(2, "0")}`;

    const nomeArmazem = getField(r, ["nome_armazem"]);
    const { parceiro: base, excluido, motivo } = extrairParceiroDoArmazem(nomeArmazem);
    if (excluido) {
      const key = motivo || "desconhecido";
      if (!excluidosPorMotivo[key]) excluidosPorMotivo[key] = { qtd: 0, receita: 0 };
      excluidosPorMotivo[key].qtd += 1;
      excluidosPorMotivo[key].receita += receita;
      return;
    }

    const parceiro = nomeFinalParceiro(base, tipo_operacao);
    notas.push({ parceiro, tipo_operacao, competencia, receita, custo });
  });

  const excluidos = Object.entries(excluidosPorMotivo)
    .map(([motivo, v]) => ({ motivo, qtd: v.qtd, receita: v.receita }))
    .sort((a, b) => b.receita - a.receita);

  return { notas, excluidos };
}

// ---------------------------------------------------------------------------
// Agregação: soma receita/custo por parceiro + tipo_operacao + competência
// ---------------------------------------------------------------------------
function agregarPorGrupo(notas) {
  const groups = {};
  notas.forEach((n) => {
    const key = `${n.parceiro}||${n.tipo_operacao}||${n.competencia}`;
    if (!groups[key]) {
      groups[key] = { parceiro: n.parceiro, tipo_operacao: n.tipo_operacao, competencia: n.competencia, receita: 0, custo: 0, qtd: 0 };
    }
    groups[key].receita += n.receita;
    groups[key].custo += n.custo;
    groups[key].qtd += 1;
  });
  return Object.values(groups);
}

// Soma um conjunto de grupos (já filtrados) num único total { receita, custo, qtd } por tipo_operacao.
function somarPorTipo(grupos) {
  const total = {
    Reversa: { receita: 0, custo: 0, qtd: 0 },
    Devolução: { receita: 0, custo: 0, qtd: 0 },
  };
  grupos.forEach((g) => {
    const bucket = total[g.tipo_operacao];
    if (!bucket) return;
    bucket.receita += g.receita;
    bucket.custo += g.custo;
    bucket.qtd += g.qtd;
  });
  return total;
}

function rentab(bucket) {
  return bucket && bucket.custo > 0 ? bucket.receita / bucket.custo : null;
}

// ---------------------------------------------------------------------------
// Componentes de UI
// ---------------------------------------------------------------------------
function Card({ children, style }) {
  return (
    <div style={{ background: C.cinzaCard, border: `1px solid ${C.cinzaBorda}`, borderRadius: 12, padding: 20, ...style }}>
      {children}
    </div>
  );
}

function KpiCard({ label, value, meta, good }) {
  const color = good === null ? C.cinzaTexto : good ? C.verde : C.vermelho;
  const bg = good === null ? C.cinzaFundo : good ? C.verdeLight : C.vermelhoLight;
  return (
    <Card style={{ flex: 1, minWidth: 220 }}>
      <div style={{ fontSize: 13, color: C.cinzaTexto, marginBottom: 6 }}>{label}</div>
      <div style={{ fontSize: 28, fontWeight: 700, color: C.texto }}>{value}</div>
      {meta && (
        <div style={{ marginTop: 8, display: "inline-block", fontSize: 12, color, background: bg, borderRadius: 6, padding: "2px 8px" }}>
          {meta}
        </div>
      )}
    </Card>
  );
}

function UploadBox({ label, hint, accept, onFile, done }) {
  const [dragging, setDragging] = useState(false);
  return (
    <div
      onDragOver={(e) => {
        e.preventDefault();
        setDragging(true);
      }}
      onDragLeave={() => setDragging(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragging(false);
        const f = e.dataTransfer.files?.[0];
        if (f) onFile(f);
      }}
      style={{
        border: `2px dashed ${dragging ? C.laranja : C.cinzaBorda}`,
        borderRadius: 12,
        padding: 24,
        textAlign: "center",
        background: done ? C.verdeLight : C.cinzaCard,
        flex: 1,
        minWidth: 260,
      }}
    >
      <div style={{ fontWeight: 600, marginBottom: 6 }}>{done ? `✅ ${label}` : label}</div>
      <div style={{ fontSize: 13, color: C.cinzaTexto, marginBottom: 12 }}>{hint}</div>
      <label style={{ display: "inline-block", padding: "8px 16px", background: C.laranja, color: "#fff", borderRadius: 8, cursor: "pointer", fontSize: 13 }}>
        {done ? "Trocar arquivo" : "Escolher arquivo"}
        <input
          type="file"
          accept={accept}
          style={{ display: "none" }}
          onChange={(e) => {
            const f = e.target.files?.[0];
            if (f) onFile(f);
          }}
        />
      </label>
    </div>
  );
}

const thStyle = {
  textAlign: "left",
  padding: "8px 10px",
  borderBottom: `2px solid ${C.cinzaBorda}`,
  color: C.cinzaTexto,
  fontWeight: 600,
  whiteSpace: "nowrap",
  position: "sticky",
  top: 0,
  background: C.cinzaCard,
};
const tdStyle = { padding: "6px 10px", borderBottom: `1px solid ${C.cinzaBorda}`, whiteSpace: "nowrap" };
const tdStyleLabel = { ...tdStyle, fontWeight: 500 };

// ---------------------------------------------------------------------------
// App principal
// ---------------------------------------------------------------------------
export default function App() {
  const [notasData, setNotasData] = useState(null); // { notas, excluidos }
  const [error, setError] = useState(null);

  const [filtroParceiro, setFiltroParceiro] = useState(TODOS);
  const [filtroGranularidade, setFiltroGranularidade] = useState("mes"); // "mes" | "ano" | "todos"
  const [filtroCompetencia, setFiltroCompetencia] = useState(TODOS);
  const [filtroAno, setFiltroAno] = useState(TODOS);

  const readFile = (file, cb) => {
    const isXlsx = /\.xlsx?$/i.test(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        if (isXlsx) {
          const wb = XLSX.read(e.target.result, { type: "array" });
          cb(parseTabular(null, wb));
        } else {
          // Exportações de planilha no Windows costumam vir em Windows-1252/Latin-1, não UTF-8.
          // Ler como UTF-8 direto corrompe caracteres acentuados. Detecta o encoding certo antes
          // de parsear.
          const bytes = new Uint8Array(e.target.result);
          let text;
          try {
            text = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
          } catch {
            text = new TextDecoder("windows-1252").decode(bytes);
          }
          cb(parseTabular(text, null));
        }
      } catch (err) {
        setError("Erro ao ler arquivo: " + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleVendaSalvadosFile = useCallback((file) => {
    readFile(file, (rows) => {
      const parsed = parseVendaSalvados(rows);
      if (parsed.notas.length === 0) {
        setError("Nenhuma nota válida foi encontrada. Verifique as colunas nome_armazem, origem, ano, mes, total_gmv, total_custo.");
        return;
      }
      setNotasData(parsed);
      setError(null);
    });
  }, []);

  const dataReady = notasData && notasData.notas.length > 0;


  // Grupos base: parceiro + tipo_operacao + competência, com receita/custo somados.
  const grupos = useMemo(() => (notasData ? agregarPorGrupo(notasData.notas) : []), [notasData]);

  const parceirosDisponiveis = useMemo(() => [...new Set(grupos.map((g) => g.parceiro))].sort((a, b) => a.localeCompare(b)), [grupos]);
  const mesesDisponiveis = useMemo(() => [...new Set(grupos.map((g) => g.competencia))].sort(), [grupos]);
  const anosDisponiveis = useMemo(() => [...new Set(grupos.map((g) => g.competencia.slice(0, 4)))].sort(), [grupos]);

  // Grupos filtrados por parça (aplicado antes de qualquer agregação temporal)
  const gruposDoParceiro = useMemo(() => {
    if (filtroParceiro === TODOS) return grupos;
    return grupos.filter((g) => g.parceiro === filtroParceiro);
  }, [grupos, filtroParceiro]);

  // Série mensal (para o gráfico) — sempre por mês, respeitando o filtro de parça e de ano (se houver)
  const serieMensal = useMemo(() => {
    const porMes = {};
    gruposDoParceiro.forEach((g) => {
      if (filtroGranularidade === "ano" && filtroAno !== TODOS && !g.competencia.startsWith(filtroAno)) return;
      if (!porMes[g.competencia]) porMes[g.competencia] = [];
      porMes[g.competencia].push(g);
    });
    return Object.entries(porMes)
      .map(([competencia, gs]) => {
        const tot = somarPorTipo(gs);
        return {
          competencia,
          rentReversa: rentab(tot.Reversa),
          rentDevolucao: rentab(tot.Devolução),
        };
      })
      .sort((a, b) => (a.competencia < b.competencia ? -1 : 1));
  }, [gruposDoParceiro, filtroGranularidade, filtroAno]);

  // Grupos do período selecionado, conforme a granularidade — usado nos KPIs e na tabela por parça.
  const gruposDoPeriodo = useMemo(() => {
    if (filtroGranularidade === "mes") {
      if (filtroCompetencia === TODOS) return gruposDoParceiro;
      return gruposDoParceiro.filter((g) => g.competencia === filtroCompetencia);
    }
    if (filtroGranularidade === "ano") {
      if (filtroAno === TODOS) return gruposDoParceiro;
      return gruposDoParceiro.filter((g) => g.competencia.startsWith(filtroAno));
    }
    return gruposDoParceiro; // "todos"
  }, [gruposDoParceiro, filtroGranularidade, filtroCompetencia, filtroAno]);

  const totalPeriodo = useMemo(() => somarPorTipo(gruposDoPeriodo), [gruposDoPeriodo]);
  const rentReversaAtual = rentab(totalPeriodo.Reversa);
  const rentDevolucaoAtual = rentab(totalPeriodo.Devolução);

  // Tabela por parça — 1 linha por parça+tipo (o nome do parça já inclui o tipo, ex.: "SAFARI - Reversa")
  const tabelaPorParceiro = useMemo(() => {
    const porParceiro = {};
    gruposDoPeriodo.forEach((g) => {
      if (!porParceiro[g.parceiro]) {
        porParceiro[g.parceiro] = { parceiro: g.parceiro, tipo_operacao: g.tipo_operacao, receita: 0, custo: 0 };
      }
      porParceiro[g.parceiro].receita += g.receita;
      porParceiro[g.parceiro].custo += g.custo;
    });
    return Object.values(porParceiro)
      .map((v) => ({
        parceiro: v.parceiro,
        tipo_operacao: v.tipo_operacao,
        rentabilidade: v.custo > 0 ? v.receita / v.custo : null,
        meta: v.tipo_operacao === "Reversa" ? META_REVERSA : META_DEVOLUCAO,
        receita: v.receita,
      }))
      .sort((a, b) => b.receita - a.receita);
  }, [gruposDoPeriodo]);

  const periodoLabel = useMemo(() => {
    if (filtroGranularidade === "mes") return filtroCompetencia === TODOS ? "todos os meses" : filtroCompetencia;
    if (filtroGranularidade === "ano") return filtroAno === TODOS ? "todos os anos" : filtroAno;
    return "todo o período";
  }, [filtroGranularidade, filtroCompetencia, filtroAno]);

  return (
    <div style={{ minHeight: "100vh", background: C.cinzaFundo, padding: "24px 32px" }}>
      <header style={{ marginBottom: 24 }}>
        <h1 style={{ fontSize: 24, margin: 0, color: C.texto }}>Painel de Performance de Faturamento Parça</h1>
        <p style={{ color: C.cinzaTexto, marginTop: 4, fontSize: 14 }}>
          Rentabilidade de reversa e devolução, por parça, mês e ano — Gestão Parça, MadeiraMadeira.
        </p>
      </header>

      {!dataReady && (
        <Card style={{ marginBottom: 24 }}>
          <div style={{ fontWeight: 600, marginBottom: 12 }}>Carregar dados</div>
          <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
            <UploadBox
              label="Venda Salvados (Looker)"
              hint="CSV nota a nota — receita e custo já faturados, parça identificado por nome_armazem"
              accept=".csv,.xlsx"
              onFile={handleVendaSalvadosFile}
              done={!!notasData}
            />
          </div>
          {error && <div style={{ marginTop: 12, color: C.vermelho, fontSize: 13 }}>{error}</div>}
        </Card>
      )}

      {dataReady && (
        <>
          {notasData.excluidos.length > 0 && (
            <Card style={{ marginBottom: 24, background: C.amareloLight, border: `1px solid ${C.amarelo}` }}>
              <div style={{ fontWeight: 600, marginBottom: 6, color: C.texto }}>
                ⚠️ Ponto de atenção — notas fora da análise
              </div>
              <div style={{ fontSize: 13, color: C.texto, marginBottom: 6 }}>
                As notas abaixo não têm um parça identificável em nome_armazem e ficaram fora dos cálculos:
              </div>
              <ul style={{ margin: 0, paddingLeft: 20, fontSize: 13 }}>
                {notasData.excluidos.map((s) => (
                  <li key={s.motivo}>
                    <strong>{s.motivo}</strong> — {s.qtd} nota(s), {fmtMoney(s.receita)} em receita
                  </li>
                ))}
              </ul>
            </Card>
          )}


          <Card style={{ marginBottom: 24 }}>
            <div style={{ display: "flex", gap: 16, flexWrap: "wrap", alignItems: "flex-end" }}>
              <div>
                <div style={{ fontSize: 12, color: C.cinzaTexto, marginBottom: 4 }}>Parça</div>
                <select
                  value={filtroParceiro}
                  onChange={(e) => setFiltroParceiro(e.target.value)}
                  style={{ padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.cinzaBorda}`, fontSize: 13, minWidth: 220 }}
                >
                  <option value={TODOS}>Todos os parças (consolidado)</option>
                  {parceirosDisponiveis.map((p) => (
                    <option key={p} value={p}>
                      {p}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <div style={{ fontSize: 12, color: C.cinzaTexto, marginBottom: 4 }}>Granularidade</div>
                <select
                  value={filtroGranularidade}
                  onChange={(e) => setFiltroGranularidade(e.target.value)}
                  style={{ padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.cinzaBorda}`, fontSize: 13, minWidth: 160 }}
                >
                  <option value="mes">Por mês</option>
                  <option value="ano">Por ano</option>
                  <option value="todos">Todo o período</option>
                </select>
              </div>

              {filtroGranularidade === "mes" && (
                <div>
                  <div style={{ fontSize: 12, color: C.cinzaTexto, marginBottom: 4 }}>Mês</div>
                  <select
                    value={filtroCompetencia}
                    onChange={(e) => setFiltroCompetencia(e.target.value)}
                    style={{ padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.cinzaBorda}`, fontSize: 13, minWidth: 180 }}
                  >
                    <option value={TODOS}>Todos os meses (consolidado)</option>
                    {mesesDisponiveis.map((m) => (
                      <option key={m} value={m}>
                        {m}
                      </option>
                    ))}
                  </select>
                </div>
              )}

              {filtroGranularidade === "ano" && (
                <div>
                  <div style={{ fontSize: 12, color: C.cinzaTexto, marginBottom: 4 }}>Ano</div>
                  <select
                    value={filtroAno}
                    onChange={(e) => setFiltroAno(e.target.value)}
                    style={{ padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.cinzaBorda}`, fontSize: 13, minWidth: 160 }}
                  >
                    <option value={TODOS}>Todos os anos</option>
                    {anosDisponiveis.map((a) => (
                      <option key={a} value={a}>
                        {a}
                      </option>
                    ))}
                  </select>
                </div>
              )}
            </div>
          </Card>

          <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 24 }}>
            <KpiCard
              label={`Rentabilidade de Reversa (${periodoLabel})`}
              value={fmtPct(rentReversaAtual)}
              meta={rentReversaAtual === null ? null : rentReversaAtual >= META_REVERSA ? `✓ acima da meta (${fmtPct(META_REVERSA)})` : `abaixo da meta (${fmtPct(META_REVERSA)})`}
              good={rentReversaAtual === null ? null : rentReversaAtual >= META_REVERSA}
            />
            <KpiCard
              label={`Rentabilidade de Devolução (${periodoLabel})`}
              value={fmtPct(rentDevolucaoAtual)}
              meta={rentDevolucaoAtual === null ? null : rentDevolucaoAtual >= META_DEVOLUCAO ? `✓ acima da meta (${fmtPct(META_DEVOLUCAO)})` : `abaixo da meta (${fmtPct(META_DEVOLUCAO)})`}
              good={rentDevolucaoAtual === null ? null : rentDevolucaoAtual >= META_DEVOLUCAO}
            />
          </div>

          <Card style={{ marginBottom: 24 }}>
            <div style={{ fontWeight: 600, marginBottom: 12 }}>Rentabilidade real x meta, por mês</div>
            <ResponsiveContainer width="100%" height={320}>
              <LineChart data={serieMensal}>
                <CartesianGrid strokeDasharray="3 3" stroke={C.cinzaBorda} />
                <XAxis dataKey="competencia" tick={{ fontSize: 12 }} />
                <YAxis tickFormatter={(v) => `${(v * 100).toFixed(0)}%`} tick={{ fontSize: 12 }} />
                <Tooltip formatter={(v) => fmtPct(v)} />
                <Legend />
                <ReferenceLine y={META_REVERSA} stroke={C.azul} strokeDasharray="4 4" label={{ value: "Meta Reversa 18%", fontSize: 11, fill: C.azul }} />
                <ReferenceLine y={META_DEVOLUCAO} stroke={C.laranja} strokeDasharray="4 4" label={{ value: "Meta Devolução 50,5%", fontSize: 11, fill: C.laranja }} />
                <Line type="monotone" dataKey="rentReversa" name="Rentabilidade Reversa" stroke={C.azul} strokeWidth={2} connectNulls dot={{ r: 3 }} />
                <Line type="monotone" dataKey="rentDevolucao" name="Rentabilidade Devolução" stroke={C.laranja} strokeWidth={2} connectNulls dot={{ r: 3 }} />
              </LineChart>
            </ResponsiveContainer>
          </Card>

          {filtroParceiro === TODOS && (
            <Card style={{ marginBottom: 24 }}>
              <div style={{ fontWeight: 600, marginBottom: 12 }}>Rentabilidade por parça — {periodoLabel}</div>
              <div style={{ maxHeight: 480, overflow: "auto" }}>
                <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 13 }}>
                  <thead>
                    <tr>
                      <th style={thStyle}>Parça</th>
                      <th style={thStyle}>Tipo</th>
                      <th style={thStyle}>Rentabilidade</th>
                      <th style={thStyle}>Meta</th>
                    </tr>
                  </thead>
                  <tbody>
                    {tabelaPorParceiro.map((r) => (
                      <tr key={r.parceiro}>
                        <td style={tdStyleLabel}>{r.parceiro}</td>
                        <td style={tdStyle}>{r.tipo_operacao}</td>
                        <td style={{ ...tdStyle, color: r.rentabilidade === null ? C.cinzaTexto : r.rentabilidade >= r.meta ? C.verde : C.vermelho }}>
                          {fmtPct(r.rentabilidade)}
                        </td>
                        <td style={tdStyle}>{fmtPct(r.meta)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Card>
          )}

          <div style={{ marginTop: 8, fontSize: 12, color: C.cinzaTexto, lineHeight: 1.6 }}>
            <strong>Sobre o cálculo desta versão simplificada:</strong> Receita e Custo vêm diretamente do arquivo
            Venda Salvados (colunas total_gmv e total_custo), já faturados — sem taxa, canal ou score aplicados.
            Rentabilidade = Receita ÷ Custo. O parça é identificado pela coluna nome_armazem (removendo o sufixo
            Reversa/Devolução); LOGME CWB e LOGME são somados como um único parça, assim como REAL e GRM MOV
            (GRM MÓVEIS (Real)). Notas de devolução sem separação de estoque por parça aparecem agrupadas em
            "Devolução 41". Sem quebra de canal (Reversa/Coleta Bulky/Recebimento Terceira) ou categoria de
            produto (Padrão/Assistência) nesta versão.
          </div>

          <div style={{ marginTop: 24 }}>
            <button
              onClick={() => {
                setNotasData(null);
              }}
              style={{ background: "transparent", border: `1px solid ${C.cinzaBorda}`, borderRadius: 8, padding: "8px 14px", cursor: "pointer", color: C.cinzaTexto, fontSize: 13 }}
            >
              Carregar outros arquivos
            </button>
          </div>
        </>
      )}
    </div>
  );
}
