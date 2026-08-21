import React, { useMemo, useState, useCallback, useEffect } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  ResponsiveContainer,
  LineChart,
  Line,
  BarChart,
  Bar,
  Cell,
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
// Variação em pontos percentuais, sempre com sinal explícito — usada na ponte de variação e
// nos chips de comparação, onde "+1,2 pp" e "1,2%" significam coisas diferentes.
function fmtPP(x, digits = 1) {
  if (x === null || x === undefined || Number.isNaN(x)) return "—";
  const v = x * 100;
  const sinal = v > 0 ? "+" : v < 0 ? "−" : "";
  return `${sinal}${Math.abs(v).toFixed(digits)} pp`;
}
function fmtMoney(x) {
  if (x === null || x === undefined || Number.isNaN(x)) return "—";
  return x.toLocaleString("pt-BR", { style: "currency", currency: "BRL", maximumFractionDigits: 0 });
}
function fmtMoneySigned(x) {
  if (x === null || x === undefined || Number.isNaN(x)) return "—";
  const sinal = x > 0 ? "+" : x < 0 ? "−" : "";
  return `${sinal}${fmtMoney(Math.abs(x))}`;
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
  const notasExcluidas = []; // mesma estrutura das notas, com `motivo` em vez de parceiro
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
    const semanaIso = getField(r, ["semana_iso"]);
    const semana = semanaIso ? `${ano}-W${String(semanaIso).padStart(2, "0")}` : null;
    const nomeArmazem = getField(r, ["nome_armazem"]);
    const { parceiro: base, excluido, motivo } = extrairParceiroDoArmazem(nomeArmazem);
    if (excluido) {
      // Guarda a nota excluída com a mesma dimensionalidade das incluídas, para que o
      // indicador de cobertura possa respeitar os filtros de período e operação em vez de
      // reportar só um total global da base.
      notasExcluidas.push({
        motivo: motivo || "desconhecido",
        tipo_operacao,
        competencia,
        semana,
        ano: String(ano),
        receita,
        custo,
      });
      return;
    }
    const parceiro = nomeFinalParceiro(base, tipo_operacao);
    notas.push({ parceiro, tipo_operacao, competencia, semana, ano: String(ano), receita, custo });
  });
  return { notas, notasExcluidas };
}
// ---------------------------------------------------------------------------
// Agregação: soma receita/custo por parceiro + tipo_operacao + semana + competência (mês)
// Mantém semana e mês juntos no mesmo grupo (cada nota já traz os dois direto da base), para
// permitir montar a série temporal em qualquer granularidade (semana, mês ou ano) a partir do
// mesmo conjunto agregado, sem precisar reagrupar as notas originais a cada troca de filtro.
// ---------------------------------------------------------------------------
function agregarPorGrupo(notas, campoChave = "parceiro") {
  const groups = {};
  notas.forEach((n) => {
    const rotulo = n[campoChave];
    const key = `${rotulo}||${n.tipo_operacao}||${n.competencia}||${n.semana}`;
    if (!groups[key]) {
      groups[key] = {
        parceiro: rotulo,
        tipo_operacao: n.tipo_operacao,
        competencia: n.competencia,
        semana: n.semana,
        ano: n.ano,
        receita: 0,
        custo: 0,
        qtd: 0,
      };
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
function metaDoTipo(tipo) {
  return tipo === "Reversa" ? META_REVERSA : META_DEVOLUCAO;
}
// ---------------------------------------------------------------------------
// Rentabilidade consolidada (Reversa + Devolução)
//
// É a soma das receitas dividida pela soma dos custos — NUNCA a média das duas
// rentabilidades. A média ignora que as duas operações têm tamanhos muito diferentes.
// ---------------------------------------------------------------------------
function consolidado(tot) {
  const receita = tot.Reversa.receita + tot.Devolução.receita;
  const custo = tot.Reversa.custo + tot.Devolução.custo;
  return custo > 0 ? receita / custo : null;
}
function custoTotal(tot) {
  return tot.Reversa.custo + tot.Devolução.custo;
}
// Meta mix-ajustada: média das duas metas (18% e 50,5%) ponderada pelo custo de cada
// operação no período. Sem isso, o consolidado não teria meta comparável — ele sobe só
// porque entrou mais devolução no mix (meta 50,5%) e desce só porque entrou mais reversa
// (meta 18%), sem ninguém ter performado melhor ou pior.
function metaConsolidada(tot) {
  const ct = custoTotal(tot);
  if (ct <= 0) return null;
  return (META_REVERSA * tot.Reversa.custo + META_DEVOLUCAO * tot.Devolução.custo) / ct;
}
// Decomposição exata da variação do consolidado entre dois períodos, em pontos percentuais.
//
// Com r = wR·rR + wD·rD (w = participação da operação no custo), a variação se abre em:
//   efeitoMix      = (wR1 − wR0)·rR0 + (wD1 − wD0)·rD0
//   efeitoReversa  = wR1·(rR1 − rR0)
//   efeitoDevolucao= wD1·(rD1 − rD0)
// A soma dos três é exatamente r1 − r0 (sem resíduo) — o que permite usar isso como ponte
// de variação (waterfall) fechada.
function decomporVariacao(tot0, tot1) {
  const ct0 = custoTotal(tot0);
  const ct1 = custoTotal(tot1);
  if (ct0 <= 0 || ct1 <= 0) return null;
  const r0 = consolidado(tot0);
  const r1 = consolidado(tot1);
  const wR0 = tot0.Reversa.custo / ct0;
  const wD0 = tot0.Devolução.custo / ct0;
  const wR1 = tot1.Reversa.custo / ct1;
  const wD1 = tot1.Devolução.custo / ct1;
  // Quando uma operação não existe num dos períodos, a rentabilidade dela é nula: usa 0
  // como referência para não propagar NaN pela ponte (o peso já é 0 nesse caso).
  const rR0 = rentab(tot0.Reversa) ?? 0;
  const rD0 = rentab(tot0.Devolução) ?? 0;
  const rR1 = rentab(tot1.Reversa) ?? 0;
  const rD1 = rentab(tot1.Devolução) ?? 0;
  const efeitoMix = (wR1 - wR0) * rR0 + (wD1 - wD0) * rD0;
  const efeitoReversa = wR1 * (rR1 - rR0);
  const efeitoDevolucao = wD1 * (rD1 - rD0);
  return {
    r0,
    r1,
    total: r1 - r0,
    efeitoMix,
    efeitoReversa,
    efeitoDevolucao,
    mixDevolucao0: wD0,
    mixDevolucao1: wD1,
  };
}
// Resultado vs. meta em R$: quanto de receita o parça entregou acima (ou abaixo) do que a
// meta da operação dele exigiria, dado o custo. Diferente da rentabilidade em %, essa
// medida é aditiva — pode ser somada entre parças, o que a torna a base correta para
// ranquear quem mais explica a variação do consolidado.
function gapVsMeta(receita, custo, tipo) {
  return receita - metaDoTipo(tipo) * custo;
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
function Chip({ children, tone = "neutro" }) {
  const map = {
    neutro: { color: C.cinzaTexto, background: C.cinzaFundo },
    bom: { color: C.verde, background: C.verdeLight },
    ruim: { color: C.vermelho, background: C.vermelhoLight },
    info: { color: C.azul, background: C.azulLight },
    atencao: { color: C.amarelo, background: C.amareloLight },
  };
  const s = map[tone] || map.neutro;
  return (
    <span style={{ display: "inline-block", fontSize: 12, borderRadius: 6, padding: "2px 8px", ...s }}>{children}</span>
  );
}
function KpiCard({ label, value, meta, good, chips, destaque }) {
  const color = good === null || good === undefined ? C.cinzaTexto : good ? C.verde : C.vermelho;
  const bg = good === null || good === undefined ? C.cinzaFundo : good ? C.verdeLight : C.vermelhoLight;
  return (
    <Card
      style={{
        flex: 1,
        minWidth: 240,
        ...(destaque ? { borderColor: C.laranja, borderWidth: 2, boxShadow: `0 1px 3px ${C.laranjaLight}` } : {}),
      }}
    >
      <div style={{ fontSize: 13, color: C.cinzaTexto, marginBottom: 6 }}>{label}</div>
      <div style={{ fontSize: 28, fontWeight: 700, color: C.texto }}>{value}</div>
      {meta && (
        <div style={{ marginTop: 8, display: "inline-block", fontSize: 12, color, background: bg, borderRadius: 6, padding: "2px 8px" }}>
          {meta}
        </div>
      )}
      {chips && chips.length > 0 && (
        <div style={{ marginTop: 8, display: "flex", gap: 6, flexWrap: "wrap" }}>{chips}</div>
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
  const [notasData, setNotasData] = useState(null); // { notas, notasExcluidas }
  const [error, setError] = useState(null);
  const [origemDados, setOrigemDados] = useState(null); // "automatico" | "manual" | null
  const [atualizadoEm, setAtualizadoEm] = useState(null);
  const [buscandoAutomatico, setBuscandoAutomatico] = useState(true);
  const [filtroParceiro, setFiltroParceiro] = useState(TODOS);
  const [filtroTipoOperacao, setFiltroTipoOperacao] = useState(TODOS); // "TODOS" (consolidado) | "Reversa" | "Devolução"
  const [filtroGranularidade, setFiltroGranularidade] = useState("mes"); // "semana" | "mes" | "ano" | "todos"
  const [filtroSemana, setFiltroSemana] = useState(TODOS);
  const [filtroCompetencia, setFiltroCompetencia] = useState(TODOS);
  const [filtroAno, setFiltroAno] = useState(TODOS);
  const [visaoGrafico, setVisaoGrafico] = useState("consolidado"); // "consolidado" | "separado"
  // Ao abrir o painel, tenta buscar automaticamente o último arquivo recebido via
  // /api/dados-atuais (alimentado pela rotina semanal do Google Apps Script). Se não houver
  // nada disponível ainda (404) ou a rota não existir (deploy sem backend configurado), o
  // painel simplesmente cai no upload manual, sem erro visível para o usuário.
  useEffect(() => {
    let cancelado = false;
    (async () => {
      try {
        const resp = await fetch("/api/dados-atuais");
        if (!resp.ok) {
          if (!cancelado) setBuscandoAutomatico(false);
          return;
        }
        const text = await resp.text();
        const atualizadoHeader = resp.headers.get("X-Atualizado-Em");
        const rows = parseTabular(text, null);
        const parsed = parseVendaSalvados(rows);
        if (cancelado) return;
        if (parsed.notas.length > 0) {
          setNotasData(parsed);
          setOrigemDados("automatico");
          setAtualizadoEm(atualizadoHeader || null);
        }
      } catch {
        // sem conexão com /api ou ambiente sem backend — segue para upload manual
      } finally {
        if (!cancelado) setBuscandoAutomatico(false);
      }
    })();
    return () => {
      cancelado = true;
    };
  }, []);
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
      setOrigemDados("manual");
      setAtualizadoEm(null);
      setError(null);
    });
  }, []);
  const dataReady = notasData && notasData.notas.length > 0;
  // Grupos base: parceiro + tipo_operacao + competência, com receita/custo somados.
  const grupos = useMemo(() => (notasData ? agregarPorGrupo(notasData.notas) : []), [notasData]);
  // Mesma agregação para as notas que ficaram fora da análise (rótulo = motivo da exclusão),
  // usada no indicador de cobertura e no banner de atenção, ambos respeitando os filtros.
  const gruposExcluidos = useMemo(
    () => (notasData ? agregarPorGrupo(notasData.notasExcluidas, "motivo") : []),
    [notasData]
  );
  const parceirosDisponiveis = useMemo(() => [...new Set(grupos.map((g) => g.parceiro))].sort((a, b) => a.localeCompare(b)), [grupos]);
  const semanasDisponiveis = useMemo(() => [...new Set(grupos.map((g) => g.semana).filter(Boolean))].sort(), [grupos]);
  const mesesDisponiveis = useMemo(() => [...new Set(grupos.map((g) => g.competencia))].sort(), [grupos]);
  const anosDisponiveis = useMemo(() => [...new Set(grupos.map((g) => g.ano))].sort(), [grupos]);
  // Grupos filtrados por parça (aplicado antes de qualquer agregação temporal)
  const gruposDoParceiro = useMemo(() => {
    if (filtroParceiro === TODOS) return grupos;
    return grupos.filter((g) => g.parceiro === filtroParceiro);
  }, [grupos, filtroParceiro]);
  // Filtro de tipo de operação (Consolidado / Reversa / Devolução) — aplicado antes de
  // qualquer agregação temporal, junto com o filtro de parça.
  const gruposFiltrados = useMemo(() => {
    if (filtroTipoOperacao === TODOS) return gruposDoParceiro;
    return gruposDoParceiro.filter((g) => g.tipo_operacao === filtroTipoOperacao);
  }, [gruposDoParceiro, filtroTipoOperacao]);
  // Chave temporal de um grupo conforme a granularidade escolhida — usada tanto na série do
  // gráfico quanto na resolução dos períodos atual/anterior da comparação.
  const chaveDe = useCallback(
    (g) => {
      if (filtroGranularidade === "semana") return g.semana;
      if (filtroGranularidade === "ano") return g.ano;
      return g.competencia; // "mes" e "todos" usam a resolução mensal
    },
    [filtroGranularidade]
  );
  // Série temporal (para o gráfico) — resolução (semana/mês/ano) acompanha a granularidade
  // selecionada, sempre mostrando a linha do tempo inteira disponível nessa resolução.
  const serieTemporal = useMemo(() => {
    const porChave = {};
    gruposFiltrados.forEach((g) => {
      const chave = chaveDe(g);
      if (!chave) return; // ex.: semana ausente porque semana_iso não veio no arquivo
      if (!porChave[chave]) porChave[chave] = [];
      porChave[chave].push(g);
    });
    return Object.entries(porChave)
      .map(([chave, gs]) => {
        const tot = somarPorTipo(gs);
        return {
          chave,
          tot,
          rentReversa: rentab(tot.Reversa),
          rentDevolucao: rentab(tot.Devolução),
          rentConsolidada: consolidado(tot),
          metaMix: metaConsolidada(tot),
        };
      })
      .sort((a, b) => (a.chave < b.chave ? -1 : 1));
  }, [gruposFiltrados, chaveDe]);
  // Grupos do período selecionado, conforme a granularidade — usado nos KPIs e na tabela por parça.
  const gruposDoPeriodo = useMemo(() => {
    if (filtroGranularidade === "semana") {
      if (filtroSemana === TODOS) return gruposFiltrados;
      return gruposFiltrados.filter((g) => g.semana === filtroSemana);
    }
    if (filtroGranularidade === "mes") {
      if (filtroCompetencia === TODOS) return gruposFiltrados;
      return gruposFiltrados.filter((g) => g.competencia === filtroCompetencia);
    }
    if (filtroGranularidade === "ano") {
      if (filtroAno === TODOS) return gruposFiltrados;
      return gruposFiltrados.filter((g) => g.ano === filtroAno);
    }
    return gruposFiltrados; // "todos"
  }, [gruposFiltrados, filtroGranularidade, filtroSemana, filtroCompetencia, filtroAno]);
  const totalPeriodo = useMemo(() => somarPorTipo(gruposDoPeriodo), [gruposDoPeriodo]);
  const rentReversaAtual = rentab(totalPeriodo.Reversa);
  const rentDevolucaoAtual = rentab(totalPeriodo.Devolução);
  const rentConsolidadaAtual = consolidado(totalPeriodo);
  const metaMixAtual = metaConsolidada(totalPeriodo);
  // -------------------------------------------------------------------------
  // Cobertura de dados: quanto do custo da base entrou de fato no cálculo.
  // Um mês com cobertura baixa não é comparável com um mês de cobertura alta — por isso
  // isso vira indicador permanente, e não apenas um aviso.
  // -------------------------------------------------------------------------
  const excluidosDoPeriodo = useMemo(() => {
    let base = gruposExcluidos;
    if (filtroTipoOperacao !== TODOS) base = base.filter((g) => g.tipo_operacao === filtroTipoOperacao);
    if (filtroGranularidade === "semana" && filtroSemana !== TODOS) base = base.filter((g) => g.semana === filtroSemana);
    if (filtroGranularidade === "mes" && filtroCompetencia !== TODOS) base = base.filter((g) => g.competencia === filtroCompetencia);
    if (filtroGranularidade === "ano" && filtroAno !== TODOS) base = base.filter((g) => g.ano === filtroAno);
    const porMotivo = {};
    base.forEach((g) => {
      if (!porMotivo[g.parceiro]) porMotivo[g.parceiro] = { motivo: g.parceiro, qtd: 0, receita: 0, custo: 0 };
      porMotivo[g.parceiro].qtd += g.qtd;
      porMotivo[g.parceiro].receita += g.receita;
      porMotivo[g.parceiro].custo += g.custo;
    });
    const itens = Object.values(porMotivo).sort((a, b) => b.custo - a.custo);
    const custoFora = itens.reduce((s, i) => s + i.custo, 0);
    const receitaFora = itens.reduce((s, i) => s + i.receita, 0);
    const custoDentro = custoTotal(totalPeriodo);
    const cobertura = custoDentro + custoFora > 0 ? custoDentro / (custoDentro + custoFora) : null;
    return { itens, custoFora, receitaFora, cobertura };
  }, [gruposExcluidos, totalPeriodo, filtroTipoOperacao, filtroGranularidade, filtroSemana, filtroCompetencia, filtroAno]);
  // -------------------------------------------------------------------------
  // Comparação entre períodos: resolve o período atual e o imediatamente anterior na
  // granularidade escolhida. Quando o filtro está em "todos", compara os dois últimos
  // períodos disponíveis — a ponte de variação e o ranking de movimentação sempre deixam
  // explícito qual par está sendo comparado.
  // -------------------------------------------------------------------------
  const chavesOrdenadas = useMemo(() => serieTemporal.map((s) => s.chave), [serieTemporal]);
  const parComparacao = useMemo(() => {
    let chaveAtual = null;
    if (filtroGranularidade === "semana" && filtroSemana !== TODOS) chaveAtual = filtroSemana;
    else if (filtroGranularidade === "mes" && filtroCompetencia !== TODOS) chaveAtual = filtroCompetencia;
    else if (filtroGranularidade === "ano" && filtroAno !== TODOS) chaveAtual = filtroAno;
    else chaveAtual = chavesOrdenadas[chavesOrdenadas.length - 1] ?? null;
    const idx = chavesOrdenadas.indexOf(chaveAtual);
    const chaveAnterior = idx > 0 ? chavesOrdenadas[idx - 1] : null;
    return { chaveAtual, chaveAnterior };
  }, [filtroGranularidade, filtroSemana, filtroCompetencia, filtroAno, chavesOrdenadas]);
  const gruposAtualComparacao = useMemo(
    () => (parComparacao.chaveAtual ? gruposFiltrados.filter((g) => chaveDe(g) === parComparacao.chaveAtual) : []),
    [gruposFiltrados, chaveDe, parComparacao.chaveAtual]
  );
  const gruposAnteriorComparacao = useMemo(
    () => (parComparacao.chaveAnterior ? gruposFiltrados.filter((g) => chaveDe(g) === parComparacao.chaveAnterior) : []),
    [gruposFiltrados, chaveDe, parComparacao.chaveAnterior]
  );
  const ponte = useMemo(() => {
    if (!parComparacao.chaveAnterior) return null;
    return decomporVariacao(somarPorTipo(gruposAnteriorComparacao), somarPorTipo(gruposAtualComparacao));
  }, [gruposAnteriorComparacao, gruposAtualComparacao, parComparacao.chaveAnterior]);
  // Dados da ponte de variação em formato de waterfall: cada barra "flutua" a partir da
  // anterior usando uma barra invisível empilhada embaixo (padrão para waterfall no Recharts).
  const dadosWaterfall = useMemo(() => {
    if (!ponte) return [];
    const passos = [
      { label: parComparacao.chaveAnterior, valor: ponte.r0, tipo: "total" },
      { label: "Mix", valor: ponte.efeitoMix, tipo: "delta" },
      { label: "Perf. Reversa", valor: ponte.efeitoReversa, tipo: "delta" },
      { label: "Perf. Devolução", valor: ponte.efeitoDevolucao, tipo: "delta" },
      { label: parComparacao.chaveAtual, valor: ponte.r1, tipo: "total" },
    ];
    let acumulado = 0;
    return passos.map((p) => {
      if (p.tipo === "total") {
        acumulado = p.valor;
        return { label: p.label, invisivel: 0, barra: p.valor, valor: p.valor, tipo: "total" };
      }
      const inicio = acumulado;
      const fim = acumulado + p.valor;
      acumulado = fim;
      return {
        label: p.label,
        invisivel: Math.min(inicio, fim),
        barra: Math.abs(p.valor),
        valor: p.valor,
        tipo: p.valor >= 0 ? "positivo" : "negativo",
      };
    });
  }, [ponte, parComparacao]);
  // -------------------------------------------------------------------------
  // Top movers: os parças que mais explicam a variação do período, medidos pelo resultado
  // vs. meta em R$ (aditivo entre parças) e não pela rentabilidade em % (que não é somável
  // e faz um parça minúsculo parecer o maior problema da carteira).
  // -------------------------------------------------------------------------
  const topMovers = useMemo(() => {
    if (!parComparacao.chaveAnterior || filtroParceiro !== TODOS) return [];
    const acumula = (grupos) => {
      const m = {};
      grupos.forEach((g) => {
        if (!m[g.parceiro]) m[g.parceiro] = { receita: 0, custo: 0, tipo_operacao: g.tipo_operacao };
        m[g.parceiro].receita += g.receita;
        m[g.parceiro].custo += g.custo;
      });
      return m;
    };
    const atual = acumula(gruposAtualComparacao);
    const anterior = acumula(gruposAnteriorComparacao);
    const nomes = [...new Set([...Object.keys(atual), ...Object.keys(anterior)])];
    return nomes
      .map((nome) => {
        const a = atual[nome];
        const b = anterior[nome];
        const tipo = (a || b).tipo_operacao;
        const gapAtual = a ? gapVsMeta(a.receita, a.custo, tipo) : 0;
        const gapAnterior = b ? gapVsMeta(b.receita, b.custo, tipo) : 0;
        return {
          parceiro: nome,
          tipo_operacao: tipo,
          deltaGap: gapAtual - gapAnterior,
          gapAtual,
          rentAtual: a && a.custo > 0 ? a.receita / a.custo : null,
          rentAnterior: b && b.custo > 0 ? b.receita / b.custo : null,
          custoAtual: a ? a.custo : 0,
          entrouOuSaiu: !a ? "saiu" : !b ? "entrou" : null,
        };
      })
      .sort((x, y) => Math.abs(y.deltaGap) - Math.abs(x.deltaGap))
      .slice(0, 5);
  }, [gruposAtualComparacao, gruposAnteriorComparacao, parComparacao.chaveAnterior, filtroParceiro]);
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
        meta: metaDoTipo(v.tipo_operacao),
        receita: v.receita,
        custo: v.custo,
        gap: gapVsMeta(v.receita, v.custo, v.tipo_operacao),
      }))
      .sort((a, b) => b.receita - a.receita);
  }, [gruposDoPeriodo]);
  const periodoLabel = useMemo(() => {
    if (filtroGranularidade === "semana") return filtroSemana === TODOS ? "todas as semanas" : filtroSemana;
    if (filtroGranularidade === "mes") return filtroCompetencia === TODOS ? "todos os meses" : filtroCompetencia;
    if (filtroGranularidade === "ano") return filtroAno === TODOS ? "todos os anos" : filtroAno;
    return "todo o período";
  }, [filtroGranularidade, filtroSemana, filtroCompetencia, filtroAno]);
  const mostrarConsolidado = filtroTipoOperacao === TODOS;
  const graficoConsolidado = mostrarConsolidado && visaoGrafico === "consolidado";
  return (
    <div style={{ minHeight: "100vh", background: C.cinzaFundo, padding: "24px 32px" }}>
      <header style={{ marginBottom: 24 }}>
        <h1 style={{ fontSize: 24, margin: 0, color: C.texto }}>Painel de Performance de Faturamento Parça</h1>
        <p style={{ color: C.cinzaTexto, marginTop: 4, fontSize: 14 }}>
          Rentabilidade de reversa e devolução, por parça, mês e ano — Gestão Parça, MadeiraMadeira.
        </p>
      </header>
      {!dataReady && buscandoAutomatico && (
        <Card style={{ marginBottom: 24 }}>
          <div style={{ fontSize: 13, color: C.cinzaTexto }}>Buscando dados atualizados automaticamente...</div>
        </Card>
      )}
      {!dataReady && !buscandoAutomatico && (
        <Card style={{ marginBottom: 24 }}>
          <div style={{ fontWeight: 600, marginBottom: 12 }}>Carregar dados</div>
          <div style={{ fontSize: 13, color: C.cinzaTexto, marginBottom: 12 }}>
            Nenhum arquivo automático disponível ainda (a rotina semanal roda toda segunda). Envie manualmente:
          </div>
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
          <Card style={{ marginBottom: 24, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12 }}>
            <div style={{ fontSize: 13, color: C.cinzaTexto }}>
              {origemDados === "automatico" ? (
                <>
                  ✅ Dados carregados automaticamente
                  {atualizadoEm && ` — atualizado em ${new Date(atualizadoEm).toLocaleString("pt-BR")}`}
                </>
              ) : (
                "📄 Dados carregados manualmente (upload local, não fica salvo entre sessões)"
              )}
            </div>
            <label style={{ display: "inline-block", padding: "6px 12px", background: "transparent", border: `1px solid ${C.cinzaBorda}`, color: C.cinzaTexto, borderRadius: 8, cursor: "pointer", fontSize: 12 }}>
              Atualizar manualmente
              <input
                type="file"
                accept=".csv,.xlsx"
                style={{ display: "none" }}
                onChange={(e) => {
                  const f = e.target.files?.[0];
                  if (f) handleVendaSalvadosFile(f);
                }}
              />
            </label>
          </Card>
          {error && <div style={{ marginBottom: 24, color: C.vermelho, fontSize: 13 }}>{error}</div>}
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
                <div style={{ fontSize: 12, color: C.cinzaTexto, marginBottom: 4 }}>Operação</div>
                <select
                  value={filtroTipoOperacao}
                  onChange={(e) => setFiltroTipoOperacao(e.target.value)}
                  style={{ padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.cinzaBorda}`, fontSize: 13, minWidth: 180 }}
                >
                  <option value={TODOS}>Consolidado (Reversa + Devolução)</option>
                  <option value="Reversa">Somente Reversa</option>
                  <option value="Devolução">Somente Devolução</option>
                </select>
              </div>
              <div>
                <div style={{ fontSize: 12, color: C.cinzaTexto, marginBottom: 4 }}>Granularidade</div>
                <select
                  value={filtroGranularidade}
                  onChange={(e) => setFiltroGranularidade(e.target.value)}
                  style={{ padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.cinzaBorda}`, fontSize: 13, minWidth: 160 }}
                >
                  <option value="semana">Por semana</option>
                  <option value="mes">Por mês</option>
                  <option value="ano">Por ano</option>
                  <option value="todos">Todo o período</option>
                </select>
              </div>
              {filtroGranularidade === "semana" && (
                <div>
                  <div style={{ fontSize: 12, color: C.cinzaTexto, marginBottom: 4 }}>Semana</div>
                  <select
                    value={filtroSemana}
                    onChange={(e) => setFiltroSemana(e.target.value)}
                    style={{ padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.cinzaBorda}`, fontSize: 13, minWidth: 180 }}
                  >
                    <option value={TODOS}>Todas as semanas (consolidado)</option>
                    {semanasDisponiveis.map((s) => (
                      <option key={s} value={s}>
                        {s}
                      </option>
                    ))}
                  </select>
                </div>
              )}
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
            {mostrarConsolidado && (
              <KpiCard
                destaque
                label={`Rentabilidade Consolidada (${periodoLabel})`}
                value={fmtPct(rentConsolidadaAtual)}
                meta={
                  rentConsolidadaAtual === null || metaMixAtual === null
                    ? null
                    : rentConsolidadaAtual >= metaMixAtual
                    ? `✓ acima da meta mix-ajustada (${fmtPct(metaMixAtual)})`
                    : `abaixo da meta mix-ajustada (${fmtPct(metaMixAtual)})`
                }
                good={rentConsolidadaAtual === null || metaMixAtual === null ? null : rentConsolidadaAtual >= metaMixAtual}
                chips={[
                  <Chip key="mix" tone="info">
                    Mix: {fmtPct(custoTotal(totalPeriodo) > 0 ? totalPeriodo.Devolução.custo / custoTotal(totalPeriodo) : null, 0)} devolução
                  </Chip>,
                  ponte ? (
                    <Chip key="var" tone={ponte.total >= 0 ? "bom" : "ruim"}>
                      {fmtPP(ponte.total)} vs. {parComparacao.chaveAnterior}
                    </Chip>
                  ) : null,
                  ponte ? (
                    <Chip key="efmix" tone={Math.abs(ponte.efeitoMix) > Math.abs(ponte.efeitoReversa + ponte.efeitoDevolucao) ? "atencao" : "neutro"}>
                      efeito mix {fmtPP(ponte.efeitoMix)}
                    </Chip>
                  ) : null,
                ].filter(Boolean)}
              />
            )}
            {filtroTipoOperacao !== "Devolução" && (
              <KpiCard
                label={`Rentabilidade de Reversa (${periodoLabel})`}
                value={fmtPct(rentReversaAtual)}
                meta={rentReversaAtual === null ? null : rentReversaAtual >= META_REVERSA ? `✓ acima da meta (${fmtPct(META_REVERSA)})` : `abaixo da meta (${fmtPct(META_REVERSA)})`}
                good={rentReversaAtual === null ? null : rentReversaAtual >= META_REVERSA}
              />
            )}
            {filtroTipoOperacao !== "Reversa" && (
              <KpiCard
                label={`Rentabilidade de Devolução (${periodoLabel})`}
                value={fmtPct(rentDevolucaoAtual)}
                meta={rentDevolucaoAtual === null ? null : rentDevolucaoAtual >= META_DEVOLUCAO ? `✓ acima da meta (${fmtPct(META_DEVOLUCAO)})` : `abaixo da meta (${fmtPct(META_DEVOLUCAO)})`}
                good={rentDevolucaoAtual === null ? null : rentDevolucaoAtual >= META_DEVOLUCAO}
              />
            )}
            <KpiCard
              label={`Cobertura de dados (${periodoLabel})`}
              value={fmtPct(excluidosDoPeriodo.cobertura, 1)}
              meta={
                excluidosDoPeriodo.cobertura === null
                  ? null
                  : excluidosDoPeriodo.cobertura >= 0.98
                  ? "✓ base praticamente completa"
                  : `${fmtMoney(excluidosDoPeriodo.custoFora)} de custo fora do cálculo`
              }
              good={excluidosDoPeriodo.cobertura === null ? null : excluidosDoPeriodo.cobertura >= 0.98}
              chips={[
                <Chip key="c" tone="neutro">
                  % do custo com parça identificado
                </Chip>,
              ]}
            />
          </div>
          {excluidosDoPeriodo.itens.length > 0 && (
            <Card style={{ marginBottom: 24, background: C.amareloLight, border: `1px solid ${C.amarelo}` }}>
              <div style={{ fontWeight: 600, marginBottom: 6, color: C.texto }}>
                ⚠️ Ponto de atenção — notas fora da análise ({periodoLabel})
              </div>
              <div style={{ fontSize: 13, color: C.texto, marginBottom: 6 }}>
                As notas abaixo não têm um parça identificável em nome_armazem e ficaram fora dos cálculos. Comparar
                períodos com coberturas diferentes distorce a variação — confira a cobertura antes de concluir que a
                rentabilidade caiu.
              </div>
              <ul style={{ margin: 0, paddingLeft: 20, fontSize: 13 }}>
                {excluidosDoPeriodo.itens.map((s) => (
                  <li key={s.motivo}>
                    <strong>{s.motivo}</strong> — {s.qtd} nota(s), {fmtMoney(s.receita)} em receita, {fmtMoney(s.custo)} em custo
                  </li>
                ))}
              </ul>
            </Card>
          )}
          <Card style={{ marginBottom: 24 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12, marginBottom: 12 }}>
              <div style={{ fontWeight: 600 }}>
                Rentabilidade real x meta, por {filtroGranularidade === "semana" ? "semana" : filtroGranularidade === "ano" ? "ano" : "mês"}
              </div>
              {mostrarConsolidado && (
                <div style={{ display: "flex", gap: 6 }}>
                  {[
                    { id: "consolidado", label: "Consolidado" },
                    { id: "separado", label: "Separado" },
                  ].map((op) => (
                    <button
                      key={op.id}
                      onClick={() => setVisaoGrafico(op.id)}
                      style={{
                        padding: "5px 12px",
                        fontSize: 12,
                        borderRadius: 8,
                        cursor: "pointer",
                        border: `1px solid ${visaoGrafico === op.id ? C.laranja : C.cinzaBorda}`,
                        background: visaoGrafico === op.id ? C.laranjaLight : "transparent",
                        color: visaoGrafico === op.id ? C.texto : C.cinzaTexto,
                      }}
                    >
                      {op.label}
                    </button>
                  ))}
                </div>
              )}
            </div>
            <ResponsiveContainer width="100%" height={320}>
              <LineChart data={serieTemporal}>
                <CartesianGrid strokeDasharray="3 3" stroke={C.cinzaBorda} />
                <XAxis dataKey="chave" tick={{ fontSize: 12 }} />
                <YAxis tickFormatter={(v) => `${(v * 100).toFixed(0)}%`} tick={{ fontSize: 12 }} />
                <Tooltip formatter={(v) => fmtPct(v)} />
                <Legend />
                {graficoConsolidado ? (
                  <>
                    <Line type="monotone" dataKey="rentConsolidada" name="Rentabilidade Consolidada" stroke={C.verde} strokeWidth={2.5} connectNulls dot={{ r: 3 }} />
                    {/* Meta mix-ajustada é uma linha (não ReferenceLine) porque muda a cada
                        período junto com o mix de reversa/devolução. */}
                    <Line type="monotone" dataKey="metaMix" name="Meta mix-ajustada" stroke={C.cinzaTexto} strokeWidth={1.5} strokeDasharray="5 4" connectNulls dot={false} />
                  </>
                ) : (
                  <>
                    {filtroTipoOperacao !== "Devolução" && (
                      <ReferenceLine y={META_REVERSA} stroke={C.azul} strokeDasharray="4 4" label={{ value: "Meta Reversa 18%", fontSize: 11, fill: C.azul }} />
                    )}
                    {filtroTipoOperacao !== "Reversa" && (
                      <ReferenceLine y={META_DEVOLUCAO} stroke={C.laranja} strokeDasharray="4 4" label={{ value: "Meta Devolução 50,5%", fontSize: 11, fill: C.laranja }} />
                    )}
                    {filtroTipoOperacao !== "Devolução" && (
                      <Line type="monotone" dataKey="rentReversa" name="Rentabilidade Reversa" stroke={C.azul} strokeWidth={2} connectNulls dot={{ r: 3 }} />
                    )}
                    {filtroTipoOperacao !== "Reversa" && (
                      <Line type="monotone" dataKey="rentDevolucao" name="Rentabilidade Devolução" stroke={C.laranja} strokeWidth={2} connectNulls dot={{ r: 3 }} />
                    )}
                  </>
                )}
              </LineChart>
            </ResponsiveContainer>
            {graficoConsolidado && (
              <div style={{ marginTop: 8, fontSize: 12, color: C.cinzaTexto }}>
                A meta mix-ajustada acompanha o mix do período: ela sobe quando entra mais devolução (meta 50,5%) e
                desce quando entra mais reversa (meta 18%). Comparar o consolidado com uma meta fixa daria uma
                leitura errada.
              </div>
            )}
          </Card>
          {mostrarConsolidado && ponte && (
            <Card style={{ marginBottom: 24 }}>
              <div style={{ fontWeight: 600, marginBottom: 4 }}>
                Ponte de variação do consolidado — {parComparacao.chaveAnterior} → {parComparacao.chaveAtual}
              </div>
              <div style={{ fontSize: 12, color: C.cinzaTexto, marginBottom: 12 }}>
                Abre a variação de {fmtPP(ponte.total)} em três efeitos que somam exatamente o total: quanto veio de
                mudança de mix entre reversa e devolução, e quanto veio de performance real dentro de cada operação.
              </div>
              <ResponsiveContainer width="100%" height={280}>
                <BarChart data={dadosWaterfall}>
                  <CartesianGrid strokeDasharray="3 3" stroke={C.cinzaBorda} />
                  <XAxis dataKey="label" tick={{ fontSize: 12 }} />
                  <YAxis tickFormatter={(v) => `${(v * 100).toFixed(0)}%`} tick={{ fontSize: 12 }} />
                  <Tooltip
                    formatter={(v, name, item) =>
                      item?.payload?.tipo === "total" ? [fmtPct(item.payload.valor), "Rentabilidade"] : [fmtPP(item.payload.valor), "Efeito"]
                    }
                  />
                  <Bar dataKey="invisivel" stackId="w" fill="transparent" isAnimationActive={false} />
                  <Bar dataKey="barra" stackId="w" isAnimationActive={false}>
                    {dadosWaterfall.map((d, i) => (
                      <Cell
                        key={i}
                        fill={d.tipo === "total" ? C.azul : d.tipo === "positivo" ? C.verde : C.vermelho}
                      />
                    ))}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 12 }}>
                <Chip tone="info">
                  Mix devolução: {fmtPct(ponte.mixDevolucao0, 0)} → {fmtPct(ponte.mixDevolucao1, 0)}
                </Chip>
                <Chip tone={ponte.efeitoMix >= 0 ? "bom" : "ruim"}>Efeito mix {fmtPP(ponte.efeitoMix)}</Chip>
                <Chip tone={ponte.efeitoReversa >= 0 ? "bom" : "ruim"}>Perf. reversa {fmtPP(ponte.efeitoReversa)}</Chip>
                <Chip tone={ponte.efeitoDevolucao >= 0 ? "bom" : "ruim"}>Perf. devolução {fmtPP(ponte.efeitoDevolucao)}</Chip>
              </div>
              {Math.abs(ponte.efeitoMix) > Math.abs(ponte.efeitoReversa) + Math.abs(ponte.efeitoDevolucao) && (
                <div style={{ marginTop: 12, fontSize: 13, color: C.texto, background: C.amareloLight, border: `1px solid ${C.amarelo}`, borderRadius: 8, padding: "8px 12px" }}>
                  ⚠️ A maior parte da variação deste período veio de <strong>mix</strong>, não de performance — o
                  consolidado mudou principalmente porque a composição entre reversa e devolução mudou.
                </div>
              )}
            </Card>
          )}
          {topMovers.length > 0 && (
            <Card style={{ marginBottom: 24 }}>
              <div style={{ fontWeight: 600, marginBottom: 4 }}>
                Quem mais explica a variação — {parComparacao.chaveAnterior} → {parComparacao.chaveAtual}
              </div>
              <div style={{ fontSize: 12, color: C.cinzaTexto, marginBottom: 12 }}>
                Ordenado pela variação do resultado vs. meta em R$ (receita − meta × custo), que é somável entre
                parças. Ranquear por rentabilidade em % faria um parça de volume irrelevante liderar a lista.
              </div>
              <div style={{ maxHeight: 320, overflow: "auto" }}>
                <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 13 }}>
                  <thead>
                    <tr>
                      <th style={thStyle}>Parça</th>
                      <th style={thStyle}>Δ resultado vs. meta</th>
                      <th style={thStyle}>Rentab. anterior</th>
                      <th style={thStyle}>Rentab. atual</th>
                      <th style={thStyle}>Meta</th>
                      <th style={thStyle}>Obs.</th>
                    </tr>
                  </thead>
                  <tbody>
                    {topMovers.map((m) => (
                      <tr key={m.parceiro}>
                        <td style={tdStyleLabel}>{m.parceiro}</td>
                        <td style={{ ...tdStyle, color: m.deltaGap >= 0 ? C.verde : C.vermelho, fontWeight: 600 }}>
                          {fmtMoneySigned(m.deltaGap)}
                        </td>
                        <td style={tdStyle}>{fmtPct(m.rentAnterior)}</td>
                        <td style={tdStyle}>{fmtPct(m.rentAtual)}</td>
                        <td style={tdStyle}>{fmtPct(metaDoTipo(m.tipo_operacao))}</td>
                        <td style={{ ...tdStyle, color: C.cinzaTexto }}>
                          {m.entrouOuSaiu === "entrou" ? "entrou no período" : m.entrouOuSaiu === "saiu" ? "sem volume no período atual" : "—"}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Card>
          )}
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
                      <th style={thStyle}>Resultado vs. meta (R$)</th>
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
                        <td style={{ ...tdStyle, color: r.gap >= 0 ? C.verde : C.vermelho }}>{fmtMoneySigned(r.gap)}</td>
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
            Rentabilidade = Receita ÷ Custo. A <strong>Rentabilidade Consolidada</strong> é (Receita Reversa +
            Receita Devolução) ÷ (Custo Reversa + Custo Devolução) — soma das partes, nunca média das duas
            rentabilidades — e é comparada com a <strong>meta mix-ajustada</strong>, média de 18% e 50,5% ponderada
            pelo custo de cada operação no período. A <strong>ponte de variação</strong> decompõe a variação do
            consolidado em efeito mix + performance de reversa + performance de devolução, e os três somam
            exatamente o total. O <strong>resultado vs. meta em R$</strong> (receita − meta × custo) é a medida
            aditiva usada para ranquear parças. A <strong>cobertura de dados</strong> é o % do custo do período com
            parça identificado em nome_armazem. O parça é identificado pela coluna nome_armazem, padronizado como
            "&lt;Nome&gt; - Reversa" / "&lt;Nome&gt; - Devolução"; LOGME CWB e LOGME são somados como um único parça,
            assim como REAL e GRM MOV (GRM MÓVEIS (Real)). Notas de devolução sem separação de estoque por parça
            aparecem agrupadas em "Devolução 41". Sem quebra de canal (Reversa/Coleta Bulky/Recebimento Terceira)
            ou categoria de produto (Padrão/Assistência) nesta versão.
          </div>
        </>
      )}
    </div>
  );
}
