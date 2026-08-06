import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import {
  AreaChart, Area, XAxis, YAxis, Tooltip, CartesianGrid,
  ResponsiveContainer,
} from "recharts";
import {
  Upload, TrendingUp, TrendingDown, RefreshCw,
  PiggyBank, HeartHandshake, Loader2, AlertCircle, Eye, EyeOff, Sun, Moon,
  Wallet, Home, Settings, ChevronRight, Banknote, Percent, Star, Info,
  Building2, Bitcoin, LineChart, DollarSign, Cross, Building, Users, Pin, PinOff,
} from "lucide-react";

const NAME = "Alexsandro Torres";
const FIRST_NAME = NAME.split(" ")[0];

// Cores das categorias, extraídas da planilha original (fixas, independem do tema)
// Paleta premium, tons neutros e complementares (variação tonal do azul de marca)
const CATEGORY_COLORS = {
  "Fundos": "#4169E1",
  "Ações": "#3FA9F5",
  "FIIs": "#2DD4BF",
  "Cripto": "#A3E635",
};

const CATEGORY_ICONS = {
  "Fundos": PiggyBank,
  "Ações": LineChart,
  "FIIs": Building2,
  "Cripto": Bitcoin,
};

// Ícones e cores para a distribuição de doações (tbDonationsDistribuition)
const DONATION_ICONS = {
  "Igrejas": Cross,
  "Organizações": Building,
  "Pessoas": Users,
};
const DONATION_COLORS = {
  dark: {
    "Igrejas": "#4DFFFF",
    "Organizações": "#34D399",
    "Pessoas": "#4C7CFF",
  },
  light: {
    "Igrejas": "#00FFFF",
    "Organizações": "#1CA67F",
    "Pessoas": "#023E8A",
  },
};

// Paleta de marca fornecida pelo usuário, aplicada aos dois temas
const THEMES = {
  dark: {
    bg: "#000000",
    bgGlow: "#000000",
    cardBg: "rgba(255,255,255,0.05)",
    cardBorder: "rgba(255,255,255,0.10)",
    tooltipBg: "rgba(19,23,27,0.95)",
    text: "#FFFFFF",
    textMuted: "rgba(255,255,255,0.55)",
    accent: "#4C7CFF",
    amber: "#FFE393",
    positive: "#00FF99",
    negative: "#FF4D6D",
    positiveBg: "rgba(0,255,153,0.14)",
    negativeBg: "rgba(255,77,109,0.14)",
    neutral: "rgba(255,255,255,0.55)",
    neutralBg: "rgba(255,255,255,0.08)",
    subCardBg: "rgba(255,255,255,0.045)",
  },
  light: {
    bg: "#FFFFFF",
    bgGlow: "#FFFFFF",
    cardBg: "rgba(215,225,236,0.55)",
    cardBorder: "rgba(19,23,27,0.09)",
    tooltipBg: "rgba(255,255,255,0.97)",
    text: "#13171B",
    textMuted: "rgba(19,23,27,0.55)",
    accent: "#0030CF",
    amber: "#B8860B",
    positive: "#0F9D63",
    negative: "#C81E3A",
    positiveBg: "rgba(15,157,99,0.12)",
    negativeBg: "rgba(200,30,58,0.10)",
    neutral: "rgba(19,23,27,0.55)",
    neutralBg: "rgba(19,23,27,0.06)",
  },
};

const STORAGE_KEYS = {
  data: "patrimonio-dashboard-data-v6",
  theme: "patrimonio-dashboard-theme-v1",
  avatar: "patrimonio-dashboard-avatar-v1",
};

const TABS = [
  { id: "home", icon: Home, label: "Home" },
  { id: "invest", icon: Wallet, label: "Wallet" },
  { id: "donations", icon: HeartHandshake, label: "Donations" },
  { id: "settings", icon: Settings, label: "Settings" },
];

function getGreeting() {
  const h = new Date().getHours();
  if (h < 12) return "Bom dia";
  if (h < 18) return "Boa tarde";
  return "Boa noite";
}

function fmtBRL(v) {
  if (v == null || Number.isNaN(v)) return "—";
  return new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" }).format(v);
}
function fmtUSD(v) {
  if (v == null || Number.isNaN(v)) return "—";
  return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" }).format(v);
}
// Formato compacto para o eixo vertical dos gráficos (ex: 5.000, 4.500)
function fmtAxisCompact(v) {
  if (v == null || Number.isNaN(v)) return "";
  return new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 0 }).format(v);
}
function fmtPct(v) {
  if (v == null || Number.isNaN(v)) return "—";
  return `${(v * 100).toFixed(2).replace(".", ",")}%`;
}
function fmtPctNoDecimal(v) {
  if (v == null || Number.isNaN(v)) return "—";
  return `${Math.round(v * 100)}%`;
}
// Cor de tendência: verde se positivo, cinza se exatamente zero, vermelho se negativo
function trendColor(v, T) {
  if (v == null || Number.isNaN(v) || v === 0) return T.neutral;
  return v > 0 ? T.positive : T.negative;
}
function trendBg(v, T) {
  if (v == null || Number.isNaN(v) || v === 0) return T.neutralBg;
  return v > 0 ? T.positiveBg : T.negativeBg;
}
const MONTH_ABBR = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"];
function fmtMonth(d) {
  const date = new Date(d);
  if (Number.isNaN(date.getTime())) return "";
  return MONTH_ABBR[date.getMonth()];
}
// Retorna só os valores de `key` de índices pares do array — usado para exibir
// no eixo X apenas um mês a cada dois (evita "embolar" os rótulos no eixo).
function everyOtherTick(arr, key) {
  if (!arr || arr.length === 0) return [];
  return arr.filter((_, i) => i % 2 === 0).map((d) => d[key]);
}
function fmtDateFull(d) {
  const date = new Date(d);
  if (Number.isNaN(date.getTime())) return "—";
  return date.toLocaleDateString("pt-BR", { day: "2-digit", month: "long", year: "numeric" });
}

function findLabelValue(rows, label) {
  for (const row of rows) {
    const idx = row.findIndex((c) => typeof c === "string" && c.trim() === label);
    if (idx !== -1) return row[idx + 1] ?? null;
  }
  return null;
}

// Localiza o bloco de linhas de uma tabela pelo nome (ex: "tbResumo"), tolerando
// sufixo numérico (ex: "tbResumo1") e caracteres extras. Retorna [startRow, endRow)
// delimitado pela próxima célula que começa com "tb" numa linha posterior, ou o fim da planilha.
function findTableBlock(rows, tableNamePrefix) {
  const prefixLower = tableNamePrefix.toLowerCase();
  let startRow = -1;
  for (let r = 0; r < rows.length; r++) {
    const cell = rows[r][0];
    if (typeof cell === "string" && cell.trim().toLowerCase().startsWith(prefixLower)) {
      startRow = r;
      break;
    }
  }
  if (startRow === -1) return null;
  let endRow = rows.length;
  for (let r = startRow + 1; r < rows.length; r++) {
    const cell = rows[r][0];
    if (typeof cell === "string" && /^tb/i.test(cell.trim())) {
      endRow = r;
      break;
    }
  }
  return [startRow, endRow];
}

// Busca um rótulo apenas dentro do bloco de linhas de uma tabela específica.
// O rótulo é comparado ignorando sufixo numérico (ex: campo "Campo1" bate com busca por "Campo").
function findLabelValueInBlock(rows, block, label) {
  if (!block) return null;
  const [start, end] = block;
  const labelLower = label.toLowerCase();
  for (let r = start; r < end; r++) {
    const row = rows[r];
    const idx = row.findIndex((c) => {
      if (typeof c !== "string") return false;
      const norm = c.trim().toLowerCase().replace(/\d+$/, "");
      return norm === labelLower;
    });
    if (idx !== -1) return row[idx + 1] ?? null;
  }
  return null;
}

function parseApiSheet(workbook) {
  const sheetName =
    workbook.SheetNames.find((n) => n.trim().toLowerCase() === "api") ||
    workbook.SheetNames.find((n) => n.toLowerCase().includes("api"));
  if (!sheetName) {
    throw new Error('Não encontrei uma aba chamada "API" nesta planilha.');
  }
  const ws = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  // Blocos de cada tabela, localizados pelo nome (tolera sufixo numérico tipo "tbResumo1")
  const blockResumo = findTableBlock(rows, "tbResumo");
  const blockEvolucao = findTableBlock(rows, "tbEvolucaoPatrimonio");
  const blockDistribuicao = findTableBlock(rows, "tbDistribuicao");
  const blockFaturamento = findTableBlock(rows, "tbFaturamento");
  const blockPatrimonio = findTableBlock(rows, "tbPatrimonio");
  const blockInvestDados = findTableBlock(rows, "tbInvestimentosDados");
  const blockMeta = findTableBlock(rows, "tbMeta");
  const blockDoacoesDist = findTableBlock(rows, "tbDoadosDistribuicao") || findTableBlock(rows, "tbDonationsDistribuition");

  // find() que usa o bloco quando disponível, e cai para busca global como último recurso
  const find = (block, label) => findLabelValueInBlock(rows, block, label) ?? findLabelValue(rows, label);

  const resumo = {
    investimentoTotal: Number(find(blockResumo, "Investimento_Total")),
    evolucaoMesValor: Number(find(blockResumo, "Evolucao_Mes_R$")),
    evolucaoMesInvestidoPercentual: Number(
      find(blockResumo, "Investido_Evolucao_Mes_Percentual") ??
      find(blockResumo, "Evolucao_Mes_Percentual")
    ),
    evolucaoMesPercentual: Number(
      find(blockPatrimonio, "Patrimonio_Evo_Mes_Percentual") ??
      find(blockPatrimonio, "Evolucao_Mes_Patrimonio") ??
      find(blockPatrimonio, "Evolucao_Mes_Percentual") ??
      find(blockPatrimonio, "Evolucao_Mês_Percentual")
    ),
    ultimaAtualizacao: find(blockResumo, "Ultima_Atualizacao"),
    totalDoadoAno: Number(find(blockResumo, "Total_Doado_Ano")),
    reservaDoacoes: Number(find(blockResumo, "ReservaDoacoes")),
    metaDoacoes: Number(find(blockResumo, "MetaDoacoes")),
    patrimonioTotal: Number(find(blockPatrimonio, "Patrimonio_Total")),
    evolucaoAnoValor: Number(find(blockPatrimonio, "Evolucao_Ano_R$")),
    evolucaoAnoPercentual: Number(find(blockPatrimonio, "Evolucao_Ano_Percentual")),
    anoReferencia: find(blockPatrimonio, "Ano_Referencia"),
  };

  const meta = {
    valorAtual: Number(find(blockMeta, "Valor_Atual")),
    valorMeta: Number(find(blockMeta, "Valor_Meta")),
    restantesMeta: Number(
      find(blockMeta, "Restantes_Meta_R$") ??
      find(blockMeta, "Restantes_Meta")
    ),
    percentualMetaAtual: Number(find(blockMeta, "Percentual_Meta_Atual")),
    objetivoPorMesR$: Number(find(blockMeta, "Objetivo_Por_Mes_Meta_R$")),
  };

  let history = [];
  {
    const block = blockEvolucao || [0, rows.length];
    const [start, end] = block;
    for (let r = start; r < end; r++) {
      const row = rows[r];
      const idx = row.findIndex((c) => typeof c === "string" && c.trim().toLowerCase().replace(/\d+$/, "") === "data");
      if (idx !== -1 && typeof row[idx + 1] === "string" && row[idx + 1].trim().toLowerCase().replace(/\d+$/, "") === "patrimonio") {
        for (let k = r + 1; k < end; k++) {
          const rr = rows[k];
          const dateVal = rr[idx];
          const patVal = rr[idx + 1];
          const invVal = rr[idx + 2];
          if (dateVal == null || patVal == null) break;
          history.push({
            date: dateVal instanceof Date ? dateVal.toISOString() : dateVal,
            patrimonio: Number(patVal),
            investido: Number(invVal),
          });
        }
        break;
      }
    }
  }

  let distribuicao = [];
  {
    const block = blockDistribuicao || [0, rows.length];
    const [start, end] = block;
    for (let r = start; r < end; r++) {
      const row = rows[r];
      const idx = row.findIndex((c) => typeof c === "string" && c.trim().toLowerCase().replace(/\d+$/, "") === "categoria");
      if (idx !== -1 && typeof row[idx + 1] === "string" && row[idx + 1].trim().toLowerCase().replace(/\d+$/, "") === "valor") {
        for (let k = r + 1; k < end; k++) {
          const rr = rows[k];
          const cat = rr[idx];
          const val = rr[idx + 1];
          const pct = rr[idx + 2];
          if (cat == null || val == null) break;
          distribuicao.push({ categoria: cat, valor: Number(val), percentual: Number(pct) });
        }
        break;
      }
    }
  }

  const faturamento = {
    refMes: find(blockFaturamento, "Faturamento_Total_R$") !== null ? find(blockFaturamento, "Ref_Mes") : null,
    totalR$: Number(find(blockFaturamento, "Faturamento_Total_R$")),
    investidoR$: Number(find(blockFaturamento, "Faturamento_Investido_R$")),
    investidoPct: Number(find(blockFaturamento, "Faturamento_Investido_Percentual")),
  };

  const investDados = {
    rentabilidadeMesR$: Number(find(blockInvestDados, "Rentabilidade_Mes_R$")),
    roiMesPct: Number(find(blockInvestDados, "ROI_Mes_Percentual")),
    rentabilidadeAnoR$: Number(find(blockInvestDados, "Rentabilidade_Ano_R$")),
    roiAnoPct: Number(find(blockInvestDados, "ROI_Ano_Percentual")),
    dividendosAno: Number(find(blockInvestDados, "Dividendos_Ano_Total")),
    emDolar: Number(find(blockInvestDados, "In_Dollars")),
    refAno: find(blockInvestDados, "Ref_Ano"),
  };

  let donationsDistribuicao = [];
  {
    const block = blockDoacoesDist || [0, rows.length];
    const [start, end] = block;
    for (let r = start; r < end; r++) {
      const row = rows[r];
      const idx = row.findIndex((c) => typeof c === "string" && c.trim().toLowerCase().replace(/\d+$/, "") === "categoria");
      if (idx !== -1 && typeof row[idx + 1] === "string" && row[idx + 1].trim().toLowerCase().replace(/\d+$/, "") === "valor") {
        for (let k = r + 1; k < end; k++) {
          const rr = rows[k];
          const cat = rr[idx];
          const val = rr[idx + 1];
          const pct = rr[idx + 2];
          if (cat == null || val == null) break;
          donationsDistribuicao.push({ categoria: cat, valor: Number(val), percentual: pct != null ? Number(pct) : null });
        }
        break;
      }
    }
  }

  return {
    resumo,
    history,
    distribuicao,
    donationsDistribuicao,
    faturamento,
    investDados,
    meta,
    ultimaAtualizacao: resumo.ultimaAtualizacao instanceof Date
      ? resumo.ultimaAtualizacao.toISOString()
      : resumo.ultimaAtualizacao,
  };
}

function useCountUp(target, duration = 800, resetKey) {
  const [value, setValue] = useState(0);
  useEffect(() => {
    if (target == null || Number.isNaN(target)) return;
    setValue(0);
    let start = null;
    let raf;
    const step = (ts) => {
      if (start === null) start = ts;
      const progress = Math.min((ts - start) / duration, 1);
      const eased = 1 - Math.pow(1 - progress, 3);
      setValue(target * eased);
      if (progress < 1) raf = requestAnimationFrame(step);
    };
    raf = requestAnimationFrame(step);
    return () => cancelAnimationFrame(raf);
  }, [target, duration, resetKey]);
  return value;
}

function StatCard({ T, icon: Icon, label, value, sub, subPositive, accent, hidden }) {
  const blurStyle = { filter: hidden ? "blur(7px)" : "none", transition: "filter 0.25s ease" };
  return (
    <div style={{
      background: T.cardBg,
      backdropFilter: "blur(16px)", WebkitBackdropFilter: "blur(16px)",
      borderRadius: 18, padding: "16px 16px", flex: 1, minWidth: 0,
      border: `1px solid ${T.cardBorder}`,
    }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
        <div style={{
          width: 28, height: 28, borderRadius: 9, background: `${accent}22`,
          display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
        }}>
          <Icon size={15} color={accent} />
        </div>
        <span style={{ fontSize: 12.5, color: T.textMuted, fontWeight: 500 }}>{label}</span>
      </div>
      <div style={{ ...blurStyle, fontFamily: "'Tomorrow', sans-serif", fontSize: 20, fontWeight: 300, letterSpacing: -0.3, color: T.text, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
        {value}
      </div>
      {sub && (
        <div style={{ ...blurStyle, fontSize: 12.5, marginTop: 4, color: subPositive ? T.positive : T.negative, fontWeight: 600 }}>
          {sub}
        </div>
      )}
    </div>
  );
}

function SectionHeader({ T, title, hideValues, onToggleHide }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18 }}>
      <span style={{ fontFamily: "'Tomorrow', sans-serif", fontSize: 18, fontWeight: 700 }}>{title}</span>
      <button onClick={onToggleHide} className="pdb-icon-btn" style={{
        cursor: "pointer", width: 34, height: 34, borderRadius: 11,
        background: T.cardBg, border: `1px solid ${T.cardBorder}`,
        backdropFilter: "blur(12px)", WebkitBackdropFilter: "blur(12px)",
        display: "flex", alignItems: "center", justifyContent: "center",
      }} title={hideValues ? "Mostrar valores" : "Ocultar valores"}>
        {hideValues ? <EyeOff size={15} color={T.textMuted} /> : <Eye size={15} color={T.textMuted} />}
      </button>
    </div>
  );
}

function BottomNav({ T, activeTab, setActiveTab }) {
  const activeIndex = TABS.findIndex((t) => t.id === activeTab);
  return (
    <div style={{
      position: "absolute", left: 34, right: 34, bottom: 14,
      background: T.cardBg, border: `1px solid ${T.cardBorder}`,
      backdropFilter: "blur(22px)", WebkitBackdropFilter: "blur(22px)",
      borderRadius: 22, display: "flex", padding: "6px 6px",
      boxShadow: "0 10px 30px rgba(0,0,0,0.28)",
    }}>
      <div style={{
        position: "absolute", top: 6, bottom: 6, width: "calc(25% - 8px)",
        left: `calc(${activeIndex} * 25% + 4px)`,
        background: `${T.accent}20`, borderRadius: 14,
        transition: "left 0.32s cubic-bezier(.4,0,.2,1)",
      }} />
      {TABS.map((tab) => {
        const Icon = tab.icon;
        const active = tab.id === activeTab;
        return (
          <button key={tab.id} onClick={() => setActiveTab(tab.id)} style={{
            flex: 1, display: "flex", flexDirection: "column", alignItems: "center", gap: 3,
            background: "none", border: "none", cursor: "pointer", padding: "8px 0 7px",
            position: "relative", zIndex: 1,
          }}>
            <Icon size={19} color={active ? T.accent : T.textMuted} style={{
              transform: active ? "scale(1.1)" : "scale(1)",
              transition: "transform 0.28s cubic-bezier(.34,1.56,.64,1), color 0.25s ease",
            }} />
            <span style={{
              fontSize: 8.5, fontWeight: 500, letterSpacing: 0.1,
              color: active ? T.accent : T.textMuted,
              transition: "color 0.25s ease",
            }}>
              {tab.label}
            </span>
          </button>
        );
      })}
    </div>
  );
}

// Donut de alocação com pontas arredondadas e espaço entre segmentos
function AllocationDonut({ T, data, centerLabel, centerValue, size = 148, hidden, colorMap, uid = "donut" }) {
  const strokeWidth = 14;
  const radius = (size - strokeWidth) / 2;
  const circumference = 2 * Math.PI * radius;
  const gapDeg = 7; // espaço visual entre segmentos, em graus
  const total = data.reduce((s, d) => s + d.percentual, 0) || 1;
  const colors = colorMap || CATEGORY_COLORS;

  let cursorDeg = -90; // começa no topo
  const segments = data.map((entry) => {
    const sweepDeg = Math.max((entry.percentual / total) * 360 - gapDeg, 4);
    const seg = { ...entry, startDeg: cursorDeg, sweepDeg };
    cursorDeg += (entry.percentual / total) * 360;
    return seg;
  });

  return (
    <div style={{ position: "relative", width: size, height: size, flexShrink: 0 }}>
      <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} style={{ transform: "rotate(0deg)" }}>
        {segments.map((seg, i) => {
          const dash = (seg.sweepDeg / 360) * circumference;
          const gap = circumference - dash;
          const rotation = seg.startDeg;
          return (
            <circle
              key={i}
              cx={size / 2}
              cy={size / 2}
              r={radius}
              fill="none"
              stroke={colors[seg.categoria] || T.textMuted}
              strokeWidth={strokeWidth}
              strokeDasharray={`${dash} ${gap}`}
              strokeLinecap="butt"
              transform={`rotate(${rotation} ${size / 2} ${size / 2})`}
              style={{ transition: "stroke-dasharray 0.6s ease" }}
            />
          );
        })}
      </svg>
      <div style={{
        position: "absolute", inset: 0, display: "flex", flexDirection: "column",
        alignItems: "center", justifyContent: "center", textAlign: "center", padding: "0 14px",
      }}>
        <span style={{ fontSize: 10, color: T.textMuted, fontWeight: 500, marginBottom: 4 }}>{centerLabel}</span>
        <span style={{
          fontFamily: "'Tomorrow', sans-serif", fontSize: 15, fontWeight: 700, color: T.text,
          lineHeight: 1.15, filter: hidden ? "blur(6px)" : "none", transition: "filter 0.25s ease",
        }}>
          {centerValue}
        </span>
        <span style={{ fontSize: 10, color: T.textMuted, marginTop: 3 }}>100%</span>
      </div>
    </div>
  );
}

// Mini sparkline decorativa para os cards de rentabilidade
function MiniSparkline({ color, seed = 1, width = 90, height = 24 }) {
  // Padrões determinísticos simples (sem dependência de dados reais de série temporal por tile)
  const patterns = [
    [4, 10, 7, 14, 11, 18, 15, 20],
    [18, 12, 15, 9, 13, 7, 10, 5],
    [10, 8, 12, 6, 14, 9, 16, 12],
    [16, 18, 13, 15, 10, 12, 8, 10],
  ];
  const pts = patterns[(seed - 1) % patterns.length];
  const stepX = width / (pts.length - 1);
  const maxY = Math.max(...pts);
  const points = pts.map((v, i) => `${i * stepX},${height - (v / maxY) * height}`).join(" ");
  return (
    <svg width={width} height={height} viewBox={`0 0 ${width} ${height}`} style={{ display: "block" }}>
      <polyline points={points} fill="none" stroke={color} strokeWidth={1.6} strokeLinecap="round" strokeLinejoin="round" opacity={0.85} />
    </svg>
  );
}

// Sparkline com dados reais (ex: últimos N meses de uma série do histórico), com leve preenchimento em gradiente
function DataSparkline({ values, color, width = 90, height = 26, uid = "spark" }) {
  if (!values || values.length < 2) return null;
  const min = Math.min(...values);
  const max = Math.max(...values);
  const range = max - min || 1;
  const stepX = width / (values.length - 1);
  const pad = 3;
  const toY = (v) => pad + (1 - (v - min) / range) * (height - pad * 2);
  const points = values.map((v, i) => `${i * stepX},${toY(v)}`);
  const linePoints = points.join(" ");
  const areaPoints = `0,${height} ${linePoints} ${width},${height}`;
  return (
    <svg width={width} height={height} viewBox={`0 0 ${width} ${height}`} style={{ display: "block" }}>
      <defs>
        <linearGradient id={`${uid}Fill`} x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stopColor={color} stopOpacity={0.35} />
          <stop offset="100%" stopColor={color} stopOpacity={0} />
        </linearGradient>
      </defs>
      <polygon points={areaPoints} fill={`url(#${uid}Fill)`} stroke="none" />
      <polyline points={linePoints} fill="none" stroke={color} strokeWidth={1.6} strokeLinecap="round" strokeLinejoin="round" />
      {(() => {
        const lastIdx = values.length - 1;
        const lx = lastIdx * stepX;
        const ly = toY(values[lastIdx]);
        return <circle cx={lx} cy={ly} r={2.2} fill={color} />;
      })()}
    </svg>
  );
}

// Gera frases dinâmicas ("insights") a partir dos dados já carregados.
// Cada insight só entra na lista se os campos necessários estiverem disponíveis.
function buildInsights(data) {
  const { resumo, history, distribuicao, donationsDistribuicao, meta, investDados, faturamento } = data || {};
  const insights = [];

  // 1. ROI do mês vs. média mensal do ano (Wallet)
  if (investDados?.roiMesPct != null && investDados?.roiAnoPct != null && !Number.isNaN(investDados.roiMesPct) && !Number.isNaN(investDados.roiAnoPct)) {
    const mediaMensal = investDados.roiAnoPct / 12;
    const spread = investDados.roiMesPct - mediaMensal;
    if (Math.abs(spread) > 0.0005) {
      insights.push({
        icon: spread > 0 ? TrendingUp : TrendingDown,
        tone: spread > 0 ? "positive" : "negative",
        text: `Seu ROI este mês está ${spread > 0 ? "acima" : "abaixo"} da média anual em ${fmtPct(Math.abs(spread))}.`,
      });
    }
  }

  // 2. Ritmo de aporte necessário para bater a meta anual de patrimônio
  if (meta?.objetivoPorMesR$ > 0 && meta?.percentualMetaAtual != null && meta.percentualMetaAtual < 1) {
    insights.push({
      icon: PiggyBank,
      tone: "accent",
      text: `Aportando ${fmtBRL(meta.objetivoPorMesR$)} por mês, você bate sua meta anual no ritmo planejado.`,
    });
  }

  // 3. Meses restantes vs. meta de patrimônio já batida
  if (meta?.percentualMetaAtual >= 1) {
    insights.push({
      icon: Star,
      tone: "positive",
      text: `Meta anual de patrimônio batida! Você já está em ${fmtPctNoDecimal(meta.percentualMetaAtual)} do objetivo.`,
    });
  }

  // 4. Sequência de meses consecutivos em alta ou baixa no histórico de patrimônio
  if (history?.length >= 2) {
    let streak = 1;
    let dir = null;
    for (let i = history.length - 1; i > 0; i--) {
      const diff = history[i].patrimonio - history[i - 1].patrimonio;
      const curDir = diff > 0 ? "up" : diff < 0 ? "down" : null;
      if (dir === null) { dir = curDir; if (!dir) break; }
      if (curDir !== dir) break;
      streak++;
    }
    if (dir && streak >= 2) {
      insights.push({
        icon: dir === "up" ? TrendingUp : TrendingDown,
        tone: dir === "up" ? "positive" : "negative",
        text: `Seu patrimônio está em ${dir === "up" ? "alta" : "queda"} há ${streak} meses seguidos.`,
      });
    }
  }

  // 5. Categoria de alocação com maior peso na carteira
  if (distribuicao?.length > 0) {
    const top = [...distribuicao].sort((a, b) => b.percentual - a.percentual)[0];
    if (top && top.percentual > 0) {
      insights.push({
        icon: PiggyBank,
        tone: "accent",
        text: `${top.categoria} representa ${fmtPctNoDecimal(top.percentual)} da sua carteira de investimentos.`,
      });
    }
  }

  // 6. Ritmo de doação vs. meta anual
  if (resumo?.metaDoacoes > 0 && resumo?.totalDoadoAno != null) {
    const restante = resumo.metaDoacoes - resumo.totalDoadoAno;
    const mesAtual = new Date(resumo.ultimaAtualizacao || Date.now()).getMonth() + 1;
    const mesesRestantes = Math.max(1, 12 - mesAtual + 1);
    if (restante > 0) {
      insights.push({
        icon: HeartHandshake,
        tone: "amber",
        text: `Faltam ${fmtBRL(restante)} para sua meta de doações. Isso é ${fmtBRL(restante / mesesRestantes)} por mês até o fim do ano.`,
      });
    } else {
      insights.push({
        icon: HeartHandshake,
        tone: "positive",
        text: `Meta de doações do ano já alcançada. Sua generosidade está em dia!`,
      });
    }
  }

  // 7. Categoria de doação que mais recebeu no período
  if (donationsDistribuicao?.length > 0) {
    const totalDoado = donationsDistribuicao.reduce((s, d) => s + (d.valor || 0), 0);
    const top = [...donationsDistribuicao].sort((a, b) => b.valor - a.valor)[0];
    if (top && totalDoado > 0) {
      const pct = top.percentual != null ? top.percentual : top.valor / totalDoado;
      insights.push({
        icon: HeartHandshake,
        tone: "amber",
        text: `${top.categoria} recebeu a maior fatia das suas doações neste período: ${fmtPctNoDecimal(pct)}.`,
      });
    }
  }

  // 8. Percentual do faturamento investido
  if (faturamento?.totalR$ > 0 && faturamento?.investidoPct != null) {
    const pct = faturamento.investidoPct;
    if (pct >= 0.3) {
      insights.push({
        icon: Banknote,
        tone: "positive",
        text: `Você investiu ${fmtPctNoDecimal(pct)} do seu faturamento este mês. Ótima disciplina!`,
      });
    } else {
      insights.push({
        icon: Banknote,
        tone: "accent",
        text: `${fmtPctNoDecimal(pct)} do seu faturamento foi investido este mês.`,
      });
    }
  }

  // 9. Evolução do patrimônio no ano
  if (resumo?.evolucaoAnoPercentual != null && !Number.isNaN(resumo.evolucaoAnoPercentual) && Math.abs(resumo.evolucaoAnoPercentual) > 0.001) {
    insights.push({
      icon: resumo.evolucaoAnoPercentual > 0 ? TrendingUp : TrendingDown,
      tone: resumo.evolucaoAnoPercentual > 0 ? "positive" : "negative",
      text: `Seu patrimônio ${resumo.evolucaoAnoPercentual > 0 ? "cresceu" : "recuou"} ${fmtPct(Math.abs(resumo.evolucaoAnoPercentual))} em ${resumo.anoReferencia || "relação ao ano anterior"}.`,
    });
  }

  return insights;
}

// Carrossel discreto de insights dinâmicos: autoplay + toque para avançar manualmente
function InsightsCarousel({ T, insights }) {
  const [index, setIndex] = useState(0);
  const [locked, setLocked] = useState(false);

  useEffect(() => {
    setIndex(0);
    setLocked(false);
  }, [insights.length]);

  useEffect(() => {
    if (insights.length <= 1 || locked) return;
    const id = setInterval(() => {
      setIndex((i) => (i + 1) % insights.length);
    }, 6000);
    return () => clearInterval(id);
  }, [insights.length, locked]);

  if (!insights.length) return null;

  const current = insights[index];
  const Icon = current.icon;
  const toneColor = {
    positive: T.positive,
    negative: T.negative,
    amber: T.amber,
    accent: T.accent,
  }[current.tone] || T.accent;

  const handleTap = () => {
    if (locked) return;
    setIndex((i) => (i + 1) % insights.length);
  };

  return (
    <div
      style={{
        background: `${toneColor}14`, border: `1px solid ${toneColor}2A`,
        borderRadius: 14, padding: "11px 13px 10px", marginBottom: 20, width: "100%",
        transition: "background 0.3s ease",
      }}
    >
      <div onClick={handleTap} style={{ display: "flex", alignItems: "center", gap: 10, cursor: insights.length > 1 && !locked ? "pointer" : "default" }}>
        <div style={{
          width: 28, height: 28, borderRadius: 9, background: `${toneColor}22`,
          display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
        }}>
          <Icon size={14} color={toneColor} />
        </div>
        <span key={index} className="pdb-fade-in" style={{ fontSize: 12, color: T.text, lineHeight: 1.4, flex: 1, minWidth: 0 }}>
          {current.text}
        </span>
        {insights.length > 1 && (
          <button
            onClick={(e) => { e.stopPropagation(); setLocked((v) => !v); }}
            title={locked ? "Destravar rotação automática" : "Travar nesta informação"}
            style={{
              cursor: "pointer", border: "none", background: "none", flexShrink: 0,
              width: 22, height: 22, display: "flex", alignItems: "center", justifyContent: "center",
              borderRadius: 6, opacity: locked ? 1 : 0.55, transition: "opacity 0.2s ease",
            }}
          >
            {locked ? <Pin size={13} color={toneColor} /> : <PinOff size={13} color={T.textMuted} />}
          </button>
        )}
      </div>
      {insights.length > 1 && (
        <div style={{ display: "flex", justifyContent: "center", gap: 3, marginTop: 9 }}>
          {insights.map((_, i) => (
            <span key={i} style={{
              width: i === index ? 12 : 4, height: 4, borderRadius: 2,
              background: i === index ? toneColor : T.cardBorder,
              transition: "width 0.25s ease, background 0.25s ease",
            }} />
          ))}
        </div>
      )}
    </div>
  );
}

export default function InvestmentDashboard() {
  const [data, setData] = useState(null);
  const [status, setStatus] = useState("loading"); // loading | empty | ready | parsing | error
  const [errorMsg, setErrorMsg] = useState("");
  const [hideValues, setHideValues] = useState(false);
  const [theme, setTheme] = useState("dark");
  const [avatar, setAvatar] = useState(null);
  const [activeTab, setActiveTab] = useState("home");
  const [navDirection, setNavDirection] = useState(1);
  const scrollRef = useRef(null);
  const changeTab = useCallback((nextTab) => {
    setActiveTab((prevTab) => {
      const prevIndex = TABS.findIndex((t) => t.id === prevTab);
      const nextIndex = TABS.findIndex((t) => t.id === nextTab);
      setNavDirection(nextIndex >= prevIndex ? 1 : -1);
      return nextTab;
    });
  }, []);

  useEffect(() => {
    if (scrollRef.current) scrollRef.current.scrollTop = 0;
  }, [activeTab]);
  const [patChartRange, setPatChartRange] = useState(12);
  const animatedPatrimonio = useCountUp(data?.resumo?.patrimonioTotal, 800, activeTab);
  const animatedDoado = useCountUp(data?.resumo?.totalDoadoAno, 800, activeTab);
  const animatedMetaPct = useCountUp(data?.meta?.percentualMetaAtual, 1000, activeTab);
  const animatedInvestido = useCountUp(data?.resumo?.investimentoTotal, 800, activeTab);
  const T = THEMES[theme];

  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get(STORAGE_KEYS.data, false);
        if (res && res.value) {
          setData(JSON.parse(res.value));
          setStatus("ready");
        } else {
          setStatus("empty");
        }
      } catch {
        setStatus("empty");
      }
    })();
  }, []);

  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get(STORAGE_KEYS.theme, false);
        if (res && res.value) setTheme(res.value);
      } catch {}
    })();
  }, []);

  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get(STORAGE_KEYS.avatar, false);
        if (res && res.value) setAvatar(res.value);
      } catch {}
    })();
  }, []);

  const handleFile = useCallback(async (file) => {
    if (!file) return;
    setStatus("parsing");
    setErrorMsg("");
    try {
      const buf = await file.arrayBuffer();
      const workbook = XLSX.read(buf, { type: "array", cellDates: true });
      const parsed = parseApiSheet(workbook);
      setData(parsed);
      setStatus("ready");
      try {
        await window.storage.set(STORAGE_KEYS.data, JSON.stringify(parsed), false);
      } catch {}
    } catch (err) {
      setErrorMsg(err.message || "Não consegui ler essa planilha.");
      setStatus(data ? "ready" : "error");
    }
  }, [data]);

  const handleAvatarFile = useCallback((file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async () => {
      const dataUrl = reader.result;
      setAvatar(dataUrl);
      try {
        await window.storage.set(STORAGE_KEYS.avatar, dataUrl, false);
      } catch {}
    };
    reader.readAsDataURL(file);
  }, []);

  const toggleTheme = useCallback(() => {
    setTheme((prev) => {
      const next = prev === "dark" ? "light" : "dark";
      window.storage.set(STORAGE_KEYS.theme, next, false).catch(() => {});
      return next;
    });
  }, []);

  const fontImport = (
    <style>{`
      * { box-sizing: border-box; }
      .pdb-icon-btn:active { transform: scale(0.94); }
      .pdb-fade-in { animation: pdbFadeIn 0.4s ease both; }
      .pdb-scroll::-webkit-scrollbar { display: none; }
      @keyframes pdbFadeIn { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: translateY(0); } }
      @keyframes spin { to { transform: rotate(360deg); } }

      /* Transição direcional entre abas da navegação inferior */
      .pdb-tab-slide-fwd { animation: pdbTabSlideFwd 0.32s cubic-bezier(.22,1,.36,1) both; }
      .pdb-tab-slide-back { animation: pdbTabSlideBack 0.32s cubic-bezier(.22,1,.36,1) both; }
      @keyframes pdbTabSlideFwd { from { opacity: 0; transform: translateX(14px); } to { opacity: 1; transform: translateX(0); } }
      @keyframes pdbTabSlideBack { from { opacity: 0; transform: translateX(-14px); } to { opacity: 1; transform: translateX(0); } }
      
      /* Premium Card Glow Effect (VisionOS inspired) */
      .pdb-card-premium {
        position: relative;
      }
      .pdb-card-premium::before {
        content: '';
        position: absolute;
        top: -40%;
        left: -40%;
        width: 80%;
        height: 80%;
        background: radial-gradient(circle at 30% 30%, rgba(255,255,255,0.13), transparent 60%);
        filter: blur(40px);
        pointer-events: none;
        z-index: 0;
      }
      .pdb-card-premium::after {
        content: '';
        position: absolute;
        bottom: -30%;
        right: -30%;
        width: 60%;
        height: 60%;
        background: radial-gradient(circle at 70% 70%, rgba(0,102,255,0.09), transparent 70%);
        filter: blur(50px);
        pointer-events: none;
        z-index: 0;
      }
      .pdb-card-premium > * {
        position: relative;
        z-index: 1;
      }
    `}</style>
  );

  const iconBtnStyle = {
    cursor: "pointer", width: 34, height: 34, borderRadius: 11,
    background: T.cardBg, border: `1px solid ${T.cardBorder}`,
    backdropFilter: "blur(12px)", WebkitBackdropFilter: "blur(12px)",
    display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
  };

  const cardStyle = {
    background: T.cardBg,
    backdropFilter: "blur(16px)", WebkitBackdropFilter: "blur(16px)",
    borderRadius: 18,
  };

  // Helper: Wrapper para cards com glow premium (VisionOS-inspired)
  // O glow fica em pseudo-elementos na div EXTERNA, o card por cima com seu fundo opaco
  const CardWrapper = ({ children, style = {}, onClick }) => (
    <div onClick={onClick} style={{ position: "relative", borderRadius: 18, flex: 1, minWidth: 0 }}>
      {/* Glow superior esquerdo */}
      <div style={{
        position: "absolute", top: "-35%", left: "-25%",
        width: "70%", height: "70%",
        background: "radial-gradient(circle at 40% 40%, rgba(220,230,255,0.13), transparent 65%)",
        filter: "blur(28px)",
        pointerEvents: "none", borderRadius: "50%",
        zIndex: 0,
      }} />
      {/* Glow inferior direito */}
      <div style={{
        position: "absolute", bottom: "-30%", right: "-20%",
        width: "55%", height: "55%",
        background: "radial-gradient(circle at 60% 60%, rgba(0,102,255,0.09), transparent 65%)",
        filter: "blur(32px)",
        pointerEvents: "none", borderRadius: "50%",
        zIndex: 0,
      }} />
      {/* Card real por cima */}
      <div style={{
        position: "relative", zIndex: 1,
        background: T.cardBg,
        backdropFilter: "blur(16px)", WebkitBackdropFilter: "blur(16px)",
        borderRadius: 18,
        border: `1px solid ${T.cardBorder}`,
        height: "100%",
        boxSizing: "border-box",
        ...style,
      }}>
        {children}
      </div>
    </div>
  );

  const shell = (children, withNav) => (
    <div style={{
      fontFamily: "'Inter', sans-serif", background: T.bgGlow, color: T.text,
      width: "100%", maxWidth: 430, margin: "0 auto", borderRadius: 0,
      height: "100dvh", minHeight: "100vh", position: "relative", overflow: "hidden",
      transition: "background 0.3s ease, color 0.3s ease",
    }}>
      {fontImport}
      <div ref={scrollRef} className="pdb-scroll" style={{ height: "100%", overflowY: "auto", scrollbarWidth: "none", padding: withNav ? "22px 18px 96px" : "22px 18px 26px" }}>
        {children}
      </div>
      {withNav && <BottomNav T={T} activeTab={activeTab} setActiveTab={changeTab} />}
    </div>
  );

  const avatarNode = (size = 46) => (
    <label className="pdb-icon-btn" style={{ position: "relative", cursor: "pointer", flexShrink: 0 }} title="Alterar foto de perfil">
      <div style={{
        width: size, height: size, borderRadius: "50%", overflow: "hidden",
        background: avatar ? "transparent" : `linear-gradient(135deg, ${T.accent}, ${CATEGORY_COLORS.Fundos})`,
        display: "flex", alignItems: "center", justifyContent: "center",
        border: `1px solid ${T.cardBorder}`,
      }}>
        {avatar ? (
          <img src={avatar} alt="Foto de perfil" style={{ width: "100%", height: "100%", objectFit: "cover" }} />
        ) : (
          <span style={{ color: "#fff", fontWeight: 700, fontSize: size * 0.32, fontFamily: "'Tomorrow', sans-serif" }}>
            {NAME.split(" ").map((n) => n[0]).slice(0, 2).join("")}
          </span>
        )}
      </div>
      <input type="file" accept="image/*" style={{ display: "none" }}
        onChange={(e) => handleAvatarFile(e.target.files[0])} />
    </label>
  );

  if (status === "loading") {
    return shell(
      <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", height: "100%", gap: 10, color: T.textMuted }}>
        <Loader2 size={22} style={{ animation: "spin 1s linear infinite" }} />
        <span style={{ fontSize: 13 }}>Carregando painel...</span>
      </div>,
      false
    );
  }

  if (status === "empty" || status === "error") {
    return shell(
      <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", height: "100%", gap: 16, textAlign: "center", padding: "0 12px" }}>
        <div style={{
          width: 64, height: 64, borderRadius: 20,
          background: `linear-gradient(135deg, ${T.accent}33, ${CATEGORY_COLORS.Fundos}33)`,
          display: "flex", alignItems: "center", justifyContent: "center",
        }}>
          <Wallet size={28} color={T.accent} />
        </div>
        <div>
          <div style={{ fontFamily: "'Tomorrow', sans-serif", fontSize: 18, fontWeight: 600, marginBottom: 6 }}>
            {getGreeting()}, {FIRST_NAME}
          </div>
          <div style={{ fontSize: 13.5, color: T.textMuted, lineHeight: 1.5 }}>
            Envie sua planilha (.xlsx) para carregar o resumo da aba <b style={{ color: T.text }}>API</b>.
          </div>
        </div>
        {errorMsg && (
          <div style={{ display: "flex", gap: 6, alignItems: "center", color: T.negative, fontSize: 12.5, background: T.negativeBg, padding: "8px 12px", borderRadius: 10 }}>
            <AlertCircle size={14} />
            {errorMsg}
          </div>
        )}
        <label className="pdb-icon-btn" style={{
          display: "flex", alignItems: "center", gap: 8, cursor: "pointer",
          background: T.accent, color: "#fff", fontWeight: 700, fontSize: 13.5,
          padding: "12px 22px", borderRadius: 14, transition: "transform 0.15s ease",
        }}>
          <Upload size={16} />
          Enviar planilha
          <input type="file" accept=".xlsx,.xls" style={{ display: "none" }}
            onChange={(e) => handleFile(e.target.files[0])} />
        </label>
        <button onClick={toggleTheme} className="pdb-icon-btn" style={{ ...iconBtnStyle, border: "none", background: "none" }}>
          {theme === "dark" ? <Sun size={16} color={T.textMuted} /> : <Moon size={16} color={T.textMuted} />}
        </button>
      </div>,
      false
    );
  }

  const { resumo, history, distribuicao, donationsDistribuicao, ultimaAtualizacao } = data;
  const donationsTotal = (donationsDistribuicao || []).reduce((s, d) => s + (d.valor || 0), 0);
  const metaAtingida = resumo.metaDoacoes > 0 ? resumo.totalDoadoAno / resumo.metaDoacoes : 0;
  const restanteMeta = Math.max(0, resumo.metaDoacoes - resumo.totalDoadoAno);
  const blurIf = (cond) => ({ filter: cond ? "blur(7px)" : "none", transition: "filter 0.25s ease" });

  // ---------- HOME ----------
  const renderHome = () => (
    <div key="home" className="pdb-fade-in">

      {/* 1. SAUDAÇÃO E DATA */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 28 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12, minWidth: 0 }}>
          {avatarNode(44)}
          <div style={{ minWidth: 0 }}>
            <div style={{ fontFamily: "'Tomorrow', sans-serif", fontSize: 15, fontWeight: 700, whiteSpace: "nowrap" }}>
              {getGreeting()}, {FIRST_NAME}
            </div>
            <div style={{ fontSize: 11, color: T.textMuted, marginTop: 2, whiteSpace: "nowrap" }}>
              {fmtDateFull(ultimaAtualizacao)}
            </div>
          </div>
        </div>
        <button onClick={() => setHideValues((v) => !v)} className="pdb-icon-btn" style={iconBtnStyle}>
          {hideValues ? <EyeOff size={15} color={T.textMuted} /> : <Eye size={15} color={T.textMuted} />}
        </button>
      </div>

      {errorMsg && (
        <div style={{ display: "flex", gap: 6, alignItems: "center", color: T.negative, fontSize: 12, background: T.negativeBg, padding: "8px 12px", borderRadius: 10, marginBottom: 14 }}>
          <AlertCircle size={13} />{errorMsg}
        </div>
      )}

      {/* 2+3+4. EVOLUÇÃO DO PATRIMÔNIO — texto solto, gráfico dentro de um card */}
      <div style={{ marginBottom: 20, width: "100%" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 16 }}>
          <span style={{ fontSize: 13, fontWeight: 600, color: T.text }}>Evolução do patrimônio</span>
          <Info size={12} color={T.textMuted} />
        </div>

        <div style={{ marginBottom: 32 }}>
          <div style={{
            fontFamily: "'Tomorrow', sans-serif", fontSize: 40, fontWeight: 400,
            letterSpacing: -1, lineHeight: 1.1, color: T.text,
            ...blurIf(hideValues),
          }}>
            {fmtBRL(animatedPatrimonio)}
          </div>
          <span style={{ fontSize: 11.5, color: T.textMuted, marginTop: 5, display: "block" }}>
            Patrimônio total
          </span>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 10, ...blurIf(hideValues) }}>
            <span style={{
              display: "flex", alignItems: "center", gap: 4, fontSize: 13, fontWeight: 700,
              color: trendColor(resumo.evolucaoMesPercentual, T),
              background: trendBg(resumo.evolucaoMesPercentual, T),
              padding: "5px 11px", borderRadius: 999, whiteSpace: "nowrap",
            }}>
              {resumo.evolucaoMesPercentual > 0 ? <TrendingUp size={13} /> : resumo.evolucaoMesPercentual < 0 ? <TrendingDown size={13} /> : null}
              {fmtPct(resumo.evolucaoMesPercentual)}
            </span>
            <span style={{ fontSize: 10.5, color: T.textMuted, whiteSpace: "nowrap" }}>vs. mês anterior</span>
          </div>
        </div>

        {/* Insights dinâmicos — assistente vivo, gerado a partir dos dados carregados */}
        <InsightsCarousel T={T} insights={buildInsights(data)} />

        {/* Gráfico dentro do card padrão do app */}
        {history?.length > 0 && (() => {
          const windowData = history.slice(-patChartRange);
          const first = windowData[0]?.patrimonio;
          const last = windowData[windowData.length - 1]?.patrimonio;
          const evoValor = first != null && last != null ? last - first : null;
          const evoPct = first ? evoValor / first : null;
          return (
            <CardWrapper style={{ padding: "14px 4px 6px", width: "100%" }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "0 12px", marginBottom: 8 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 6, ...blurIf(hideValues) }}>
                  {evoValor != null && (
                    <span style={{ fontSize: 11, fontWeight: 700, color: trendColor(evoValor, T) }}>
                      {evoValor >= 0 ? "+" : ""}{fmtBRL(evoValor)}
                    </span>
                  )}
                  {evoPct != null && (
                    <span style={{
                      fontSize: 10.5, fontWeight: 600, color: trendColor(evoPct, T),
                      background: trendBg(evoPct, T), padding: "2px 7px", borderRadius: 999,
                    }}>
                      {evoPct >= 0 ? "+" : ""}{fmtPct(evoPct)}
                    </span>
                  )}
                </div>
                <div style={{ display: "flex", gap: 4 }}>
                  {[6, 12].map((months) => {
                    const active = patChartRange === months;
                    return (
                      <button
                        key={months}
                        onClick={() => setPatChartRange(months)}
                        style={{
                          cursor: "pointer", border: "none", borderRadius: 999,
                          padding: "4px 11px", fontSize: 10.5, fontWeight: 600,
                          background: active ? `${T.accent}22` : "transparent",
                          color: active ? T.accent : T.textMuted,
                          transition: "background 0.2s ease, color 0.2s ease",
                        }}
                      >
                        {months}M
                      </button>
                    );
                  })}
                </div>
              </div>
              <ResponsiveContainer width="100%" height={190}>
                <AreaChart data={windowData} margin={{ top: 10, right: 22, left: 8, bottom: 0 }}>
                  <defs>
                  <linearGradient id="patGrad" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor={T.accent} stopOpacity={0.55} />
                    <stop offset="55%" stopColor={T.accent} stopOpacity={0.12} />
                    <stop offset="100%" stopColor={T.accent} stopOpacity={0} />
                  </linearGradient>
                  <filter id="patGlow" x="-50%" y="-50%" width="200%" height="200%">
                    <feGaussianBlur stdDeviation="5" result="blur" />
                    <feMerge>
                      <feMergeNode in="blur" />
                      <feMergeNode in="SourceGraphic" />
                    </feMerge>
                  </filter>
                </defs>
                <CartesianGrid strokeDasharray="3 5" vertical={false} stroke={T.cardBorder} strokeOpacity={0.6} />
                <XAxis
                  dataKey="date" tickFormatter={fmtMonth}
                  ticks={everyOtherTick(windowData, "date")}
                  tick={{ fill: T.textMuted, fontSize: 10.5, fontWeight: 600 }}
                  axisLine={{ stroke: T.cardBorder }} tickLine={false} dy={8}
                />
                <YAxis
                  tickFormatter={fmtAxisCompact}
                  tick={{ fill: T.textMuted, fontSize: 9.5, fontWeight: 400 }}
                  axisLine={false} tickLine={false} width={38} tickCount={4}
                  domain={["dataMin - dataMin * 0.06", "dataMax + dataMax * 0.04"]}
                />
                <Tooltip
                  contentStyle={{ background: T.tooltipBg, border: `1px solid ${T.cardBorder}`, borderRadius: 10, fontSize: 12 }}
                  labelStyle={{ color: T.textMuted }} labelFormatter={fmtMonth}
                  formatter={(v) => [fmtBRL(v), "Patrimônio"]}
                />
                <Area
                  type="monotone" dataKey="patrimonio"
                  stroke={T.accent} strokeWidth={2.5}
                  fill="url(#patGrad)"
                  style={{ filter: "url(#patGlow)" }}
                  dot={false}
                  activeDot={{ r: 4, fill: T.accent, stroke: T.bg, strokeWidth: 2 }}
                />
              </AreaChart>
              </ResponsiveContainer>
            </CardWrapper>
          );
        })()}
      </div>

      {/* 5. CARDS: Meta mensal */}
      {data?.meta && (
        <div style={{ marginBottom: 14 }}>
          <CardWrapper style={{ padding: "16px 16px 14px", width: "100%" }}>
            <div style={{ marginBottom: 8 }}>
              <div style={{ fontSize: 11, color: T.textMuted, fontWeight: 500, textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 6 }}>Meta anual</div>
              <div style={{ fontFamily: "'Tomorrow', sans-serif", fontSize: 32, fontWeight: 400, letterSpacing: -0.5, color: T.text, ...blurIf(hideValues) }}>
                {fmtPctNoDecimal(animatedMetaPct)}
              </div>
            </div>
            <div style={{ height: 12, borderRadius: 999, overflow: "hidden", background: T.cardBorder, marginBottom: 10, width: "100%" }}>
              <style>{`
                @keyframes fillBar { from { width: 0; } to { width: ${Math.min(100, data.meta.percentualMetaAtual * 100)}%; } }
                .pdb-meta-bar-v2 { animation: fillBar 1s cubic-bezier(0.34,1.56,0.64,1) forwards; border-radius:999px; background:linear-gradient(90deg,#001a4d 0%,#0066FF 50%,#45C4B0 100%); height:100%; box-shadow:0 0 18px rgba(0,102,255,0.3); }
              `}</style>
              <div className="pdb-meta-bar-v2" />
            </div>
            <div style={{ width: "100%" }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginBottom: 3 }}>
                <span style={{ color: T.text, fontWeight: 600, ...blurIf(hideValues) }}>{fmtBRL(data.meta.valorAtual)}</span>
                <span style={{ color: T.text, fontWeight: 600, ...blurIf(hideValues) }}>{fmtBRL(data.meta.valorMeta)}</span>
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 10, color: T.textMuted, textTransform: "uppercase", letterSpacing: 0.4 }}>
                <span>Atual</span><span>Meta</span>
              </div>
            </div>
            {data.meta.objetivoPorMesR$ > 0 && data.meta.percentualMetaAtual < 1 && (
              <div style={{
                marginTop: 12, paddingTop: 12, borderTop: `1px solid ${T.cardBorder}`,
                display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8,
              }}>
                <span style={{ fontSize: 11, color: T.textMuted }}>Para bater a meta</span>
                <span style={{ fontFamily: "'Tomorrow', sans-serif", fontSize: 13, fontWeight: 700, color: T.accent, ...blurIf(hideValues) }}>
                  {fmtBRL(data.meta.objetivoPorMesR$)}<span style={{ fontSize: 10.5, color: T.textMuted, fontWeight: 500 }}>/mês</span>
                </span>
              </div>
            )}
          </CardWrapper>
        </div>
      )}

      {/* 6. CARDS: Evolução anual — turquesa, texto ao lado do ícone, valor grande abaixo */}
      <div style={{ marginBottom: 12 }}>
        <span style={{ fontSize: 13, fontWeight: 600, color: T.text }}>Evolução Anual de Patrimônio</span>
      </div>
      <div style={{ display: "flex", gap: 10, marginBottom: 14 }}>
        <div style={{ flex: 1, minWidth: 0 }}>
          <CardWrapper style={{ padding: 0, overflow: "hidden" }}>
            <div style={{
              padding: "14px 14px 12px",
              background: `linear-gradient(160deg, ${CATEGORY_COLORS.FIIs}22 0%, transparent 65%)`,
            }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
                <div style={{
                  width: 26, height: 26, borderRadius: 8,
                  background: `${CATEGORY_COLORS.FIIs}22`,
                  display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
                }}>
                  <Percent size={13} color={CATEGORY_COLORS.FIIs} />
                </div>
                <span style={{ fontSize: 11.5, color: T.textMuted, fontWeight: 500 }}>Evolução em %</span>
              </div>
              <span style={{
                fontFamily: "'Tomorrow', sans-serif", fontSize: 22, fontWeight: 300, letterSpacing: -0.5,
                color: trendColor(resumo.evolucaoAnoPercentual, T), display: "block", ...blurIf(hideValues),
              }}>
                {fmtPct(resumo.evolucaoAnoPercentual)}
              </span>
            </div>
          </CardWrapper>
        </div>
        <div style={{ flex: 1, minWidth: 0 }}>
          <CardWrapper style={{ padding: 0, overflow: "hidden" }}>
            <div style={{
              padding: "14px 14px 12px",
              background: `linear-gradient(160deg, ${CATEGORY_COLORS.FIIs}22 0%, transparent 65%)`,
            }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
                <div style={{
                  width: 26, height: 26, borderRadius: 8,
                  background: `${CATEGORY_COLORS.FIIs}22`,
                  display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
                }}>
                  <Banknote size={13} color={CATEGORY_COLORS.FIIs} />
                </div>
                <span style={{ fontSize: 11.5, color: T.textMuted, fontWeight: 500 }}>Evolução em R$</span>
              </div>
              <span style={{
                fontFamily: "'Tomorrow', sans-serif", fontSize: 22, fontWeight: 300, letterSpacing: -0.5,
                color: T.text, display: "block", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis",
                ...blurIf(hideValues),
              }}>
                {fmtBRL(resumo.evolucaoAnoValor)}
              </span>
            </div>
          </CardWrapper>
        </div>
      </div>

      {/* 7. CARDS: Investido e Faturamento */}
      <div style={{ display: "flex", gap: 10, marginBottom: 14 }}>
        <CardWrapper style={{ padding: "16px", flex: 1, minWidth: 0 }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, marginBottom: 10 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, minWidth: 0 }}>
              <div style={{ width: 26, height: 26, borderRadius: 8, background: `${T.accent}22`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                <PiggyBank size={14} color={T.accent} />
              </div>
              <span style={{ fontSize: 12, color: T.textMuted, fontWeight: 500 }}>Investido</span>
            </div>
            {history?.length > 1 && (
              <DataSparkline
                values={history.slice(-6).map((h) => h.investido)}
                color={T.accent}
                width={46}
                height={20}
                uid="investSpark"
              />
            )}
          </div>
          <div style={{ fontFamily: "'Tomorrow', sans-serif", fontSize: 19, fontWeight: 400, color: T.text, marginBottom: 12, ...blurIf(hideValues) }}>
            {fmtBRL(resumo.investimentoTotal)}
          </div>
          {data?.investDados && (
            <div style={{ background: T.subCardBg, borderRadius: 11, padding: "9px 11px" }}>
              <div style={{ fontSize: 10, color: T.textMuted, marginBottom: 5 }}>Rentabilidade</div>
              <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 3, ...blurIf(hideValues) }}>
                <div style={{ width: 20, height: 20, borderRadius: 6, background: trendBg(data.investDados.rentabilidadeMesR$, T), display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                  {data.investDados.rentabilidadeMesR$ > 0 ? <TrendingUp size={11} color={trendColor(data.investDados.rentabilidadeMesR$, T)} /> : data.investDados.rentabilidadeMesR$ < 0 ? <TrendingDown size={11} color={trendColor(data.investDados.rentabilidadeMesR$, T)} /> : <TrendingUp size={11} color={T.neutral} />}
                </div>
                <span style={{ fontFamily: "'Tomorrow', sans-serif", fontSize: 15, fontWeight: 300, color: trendColor(data.investDados.rentabilidadeMesR$, T) }}>
                  {fmtBRL(data.investDados.rentabilidadeMesR$)}
                </span>
              </div>
              <span style={{ fontSize: 11, color: T.textMuted, ...blurIf(hideValues) }}>
                {(data.investDados.roiMesPct * 100).toFixed(2).replace(".", ",")}%
              </span>
            </div>
          )}
        </CardWrapper>

        {data?.faturamento && data.faturamento.totalR$ > 0 && (
          <CardWrapper style={{ padding: "16px", flex: 1, minWidth: 0 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
              <div style={{ width: 26, height: 26, borderRadius: 8, background: `${CATEGORY_COLORS.Fundos}22`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                <Banknote size={14} color={CATEGORY_COLORS.Fundos} />
              </div>
              <span style={{ fontSize: 12, color: T.textMuted, fontWeight: 500 }}>Faturamento</span>
            </div>
            <div style={{ fontFamily: "'Tomorrow', sans-serif", fontSize: 19, fontWeight: 400, color: T.text, marginBottom: 12, ...blurIf(hideValues) }}>
              {fmtBRL(data.faturamento.totalR$)}
            </div>
            <div style={{ background: T.subCardBg, borderRadius: 11, padding: "9px 11px" }}>
              <div style={{ fontSize: 10, color: T.textMuted, marginBottom: 5 }}>Investido</div>
              <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 3, ...blurIf(hideValues) }}>
                <div style={{ width: 20, height: 20, borderRadius: 6, background: T.cardBorder, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                  <PiggyBank size={11} color={T.textMuted} />
                </div>
                <span style={{ fontFamily: "'Tomorrow', sans-serif", fontSize: 15, fontWeight: 300, color: T.text }}>
                  {fmtBRL(data.faturamento.investidoR$)}
                </span>
              </div>
              <span style={{ fontSize: 11, color: T.textMuted, ...blurIf(hideValues) }}>
                {fmtPctNoDecimal(data.faturamento.investidoPct)}
              </span>
            </div>
          </CardWrapper>
        )}
      </div>

      {/* 8. CARD: Doações */}
      {resumo.metaDoacoes > 0 && (
        <CardWrapper style={{ padding: "14px 16px", width: "100%", cursor: "pointer" }} onClick={() => changeTab("donations")}>
          <button onClick={() => changeTab("donations")} style={{ display: "contents" }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10 }}>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                  <HeartHandshake size={14} color={T.amber} />
                  <span style={{ fontSize: 12, fontWeight: 600, color: T.text }}>Doações em {resumo.anoReferencia}</span>
                </div>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginBottom: 6 }}>
                  <span style={{ ...blurIf(hideValues) }}>
                    <span style={{ color: T.text, fontWeight: 700 }}>{fmtBRL(resumo.totalDoadoAno)}</span>
                    <span style={{ color: T.textMuted }}> doados</span>
                  </span>
                  <span style={{ color: T.textMuted, ...blurIf(hideValues) }}>meta {fmtBRL(resumo.metaDoacoes)}</span>
                </div>
                <div style={{ height: 5, borderRadius: 4, background: T.cardBorder, overflow: "hidden" }}>
                  <div style={{ height: "100%", borderRadius: 4, width: `${Math.min(100, metaAtingida * 100)}%`, background: "linear-gradient(90deg, #023E8A 0%, #0077B6 45%, #00B4D8 100%)", transition: "width 0.8s ease" }} />
                </div>
              </div>
              <ChevronRight size={15} color={T.textMuted} style={{ flexShrink: 0 }} />
            </div>
          </button>
        </CardWrapper>
      )}
    </div>
  );

  // ---------- INVESTIMENTOS ----------
  const renderInvest = () => {
    const fat = data?.faturamento;
    const inv = data?.investDados;

    return (
      <div key="invest" className="pdb-fade-in">
        <SectionHeader T={T} title="Investimentos" hideValues={hideValues} onToggleHide={() => setHideValues((v) => !v)} />

        {/* 1. VALOR INVESTIDO TOTAL — protagonista, fonte fina */}
        <div style={{ marginBottom: 16 }}>
          <span style={{ fontSize: 11.5, color: T.textMuted, fontWeight: 500, letterSpacing: 0.5, textTransform: "uppercase" }}>
            Total investido
          </span>
          <div style={{
            fontFamily: "'Tomorrow', sans-serif", fontSize: 40, fontWeight: 400,
            letterSpacing: -1, lineHeight: 1, marginTop: 6, color: T.text,
            ...blurIf(hideValues),
          }}>
            {fmtBRL(animatedInvestido)}
          </div>
        </div>

        {/* 2. Evolução do mês — R$ e % */}
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 22, ...blurIf(hideValues) }}>
          <span style={{
            display: "flex", alignItems: "center", gap: 5, fontSize: 13, fontWeight: 600,
            color: trendColor(resumo.evolucaoMesInvestidoPercentual, T),
            background: trendBg(resumo.evolucaoMesInvestidoPercentual, T),
            padding: "5px 10px 5px 8px", borderRadius: 999,
          }}>
            {resumo.evolucaoMesInvestidoPercentual > 0 ? <TrendingUp size={13} /> : resumo.evolucaoMesInvestidoPercentual < 0 ? <TrendingDown size={13} /> : null}
            {fmtBRL(resumo.evolucaoMesValor)}
          </span>
          <span style={{ fontSize: 13, fontWeight: 600, color: trendColor(resumo.evolucaoMesInvestidoPercentual, T) }}>
            {fmtPct(resumo.evolucaoMesInvestidoPercentual)}
          </span>
          <span style={{ fontSize: 12, color: T.textMuted }}>este mês</span>
        </div>

        {/* 3. Gráfico do investido — mesmo tratamento premium do gráfico da Home */}
        {history?.length > 0 && (
          <CardWrapper style={{ padding: "18px 12px 10px", marginBottom: 14 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 14, paddingLeft: 6 }}>
              <span style={{ fontSize: 13, fontWeight: 600, color: T.text }}>Evolução do investido</span>
              <Info size={12} color={T.textMuted} />
            </div>
            <ResponsiveContainer width="100%" height={158}>
              <AreaChart data={history} margin={{ top: 10, right: 18, left: -6, bottom: 0 }}>
                <defs>
                  <linearGradient id="invGrad" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor={T.accent} stopOpacity={0.5} />
                    <stop offset="55%" stopColor={T.accent} stopOpacity={0.1} />
                    <stop offset="100%" stopColor={T.accent} stopOpacity={0} />
                  </linearGradient>
                  <filter id="invGlow" x="-50%" y="-50%" width="200%" height="200%">
                    <feGaussianBlur stdDeviation="4.5" result="blur" />
                    <feMerge>
                      <feMergeNode in="blur" />
                      <feMergeNode in="SourceGraphic" />
                    </feMerge>
                  </filter>
                </defs>
                <CartesianGrid strokeDasharray="3 5" vertical={false} stroke={T.cardBorder} strokeOpacity={0.6} />
                <XAxis
                  dataKey="date" tickFormatter={fmtMonth}
                  ticks={everyOtherTick(history, "date")}
                  tick={{ fill: T.textMuted, fontSize: 10.5, fontWeight: 600 }}
                  axisLine={{ stroke: T.cardBorder }} tickLine={false} dy={8}
                />
                <YAxis
                  tickFormatter={fmtAxisCompact}
                  tick={{ fill: T.textMuted, fontSize: 9.5, fontWeight: 400 }}
                  axisLine={false} tickLine={false} width={38} tickCount={4}
                  domain={["dataMin - dataMin * 0.06", "dataMax + dataMax * 0.04"]}
                />
                <Tooltip
                  contentStyle={{ background: T.tooltipBg, border: `1px solid ${T.cardBorder}`, borderRadius: 10, fontSize: 12 }}
                  labelStyle={{ color: T.textMuted }} labelFormatter={fmtMonth}
                  formatter={(v) => [fmtBRL(v), "Investido"]}
                />
                <Area
                  type="monotone" dataKey="investido"
                  stroke={T.accent} strokeWidth={2.5}
                  fill="url(#invGrad)"
                  style={{ filter: "url(#invGlow)" }}
                  dot={false}
                  activeDot={{ r: 4, fill: T.accent, stroke: T.bg, strokeWidth: 2 }}
                />
              </AreaChart>
            </ResponsiveContainer>
          </CardWrapper>
        )}

        {/* 4. Alocação — donut centralizado no topo + lista com ícones abaixo */}
        {distribuicao?.length > 0 && (
          <CardWrapper style={{ padding: "16px 16px 16px", marginBottom: 14 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 18 }}>
              <span style={{ fontSize: 13, fontWeight: 600, color: T.text }}>Alocação</span>
              <Info size={12} color={T.textMuted} />
            </div>
            <div style={{ display: "flex", justifyContent: "center", marginBottom: 22 }}>
              <AllocationDonut
                T={T}
                data={distribuicao}
                centerLabel="Total investido"
                centerValue={fmtBRL(resumo.investimentoTotal)}
                hidden={hideValues}
                size={172}
              />
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
              {distribuicao.map((entry, i) => {
                const AssetIcon = CATEGORY_ICONS[entry.categoria] || Percent;
                const color = CATEGORY_COLORS[entry.categoria] || T.textMuted;
                return (
                  <div key={i} style={{
                    position: "relative", display: "flex", alignItems: "stretch", gap: 9,
                    borderRadius: 12, overflow: "hidden",
                    background: `linear-gradient(120deg, ${color}14 0%, transparent 75%)`,
                    border: `1px solid ${T.cardBorder}`,
                    padding: "10px 10px 10px 0",
                    minWidth: 0,
                  }}>
                    {/* Barra vertical de proporção — assinatura visual do card */}
                    <div style={{ width: 3, borderRadius: 2, background: T.cardBorder, flexShrink: 0, overflow: "hidden", display: "flex", alignItems: "flex-end" }}>
                      <div style={{
                        width: "100%", borderRadius: 2, background: color,
                        height: `${Math.max(Math.min(100, entry.percentual * 100), 10)}%`,
                        transition: "height 0.6s ease",
                      }} />
                    </div>
                    <div style={{ minWidth: 0, flex: 1 }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 5, marginBottom: 6 }}>
                        <AssetIcon size={11} color={color} style={{ flexShrink: 0 }} />
                        <span style={{ fontSize: 10.5, color: T.textMuted, fontWeight: 600, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                          {entry.categoria}
                        </span>
                      </div>
                      <div style={{
                        fontFamily: "'Tomorrow', sans-serif", fontSize: 17, fontWeight: 400, letterSpacing: -0.3,
                        color: T.text, lineHeight: 1.1,
                      }}>
                        {fmtPct(entry.percentual)}
                      </div>
                      <div style={{ fontSize: 9.5, color: T.textMuted, marginTop: 2, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", ...blurIf(hideValues) }}>
                        {fmtBRL(entry.valor)}
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          </CardWrapper>
        )}

        {/* Valor em dólar e Dividendos anuais — cards lado a lado, logo após Alocação */}
        {inv && (
          <div style={{ display: "flex", gap: 10, marginBottom: 14 }}>
            <CardWrapper style={{ padding: "16px", flex: 1, minWidth: 0 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
                <div style={{ width: 26, height: 26, borderRadius: 8, background: `${T.positive}22`, display: "flex", alignItems: "center", justifyContent: "center" }}>
                  <DollarSign size={14} color={T.positive} />
                </div>
                <span style={{ fontSize: 12, color: T.textMuted, fontWeight: 500 }}>Em dólar</span>
              </div>
              <div style={{ fontFamily: "'Tomorrow', sans-serif", fontSize: 19, fontWeight: 400, color: T.text, ...blurIf(hideValues) }}>
                {fmtUSD(inv.emDolar)}
              </div>
            </CardWrapper>

            <CardWrapper style={{ padding: "16px", flex: 1, minWidth: 0 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
                <div style={{ width: 26, height: 26, borderRadius: 8, background: `${T.amber}22`, display: "flex", alignItems: "center", justifyContent: "center" }}>
                  <Star size={14} color={T.amber} />
                </div>
                <span style={{ fontSize: 12, color: T.textMuted, fontWeight: 500 }}>Dividendos no ano</span>
              </div>
              <div style={{ fontFamily: "'Tomorrow', sans-serif", fontSize: 19, fontWeight: 400, color: T.amber, ...blurIf(hideValues) }}>
                {fmtBRL(inv.dividendosAno)}
              </div>
            </CardWrapper>
          </div>
        )}

        {/* Rentabilidade (tbInvestimentosDados) — 4 mini-cards compactos, estilo referência */}
        {inv && (
          <CardWrapper style={{ padding: "16px 16px 14px", marginBottom: 14 }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                <span style={{ fontSize: 13, fontWeight: 600, color: T.text }}>Rentabilidade</span>
                <Info size={12} color={T.textMuted} />
              </div>
              {inv.refAno && (
                <span style={{
                  fontSize: 10.5, fontWeight: 700, color: "#fff", background: T.accent,
                  padding: "4px 9px", borderRadius: 999, letterSpacing: 0.4,
                }}>
                  {inv.refAno}
                </span>
              )}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 7, marginBottom: 14 }}>
              {[
                { label: "No mês (R$)", value: fmtBRL(inv.rentabilidadeMesR$), signal: inv.rentabilidadeMesR$, seed: 1 },
                { label: "ROI no mês", value: fmtPct(inv.roiMesPct), signal: inv.roiMesPct, seed: 2 },
                { label: "No ano (R$)", value: fmtBRL(inv.rentabilidadeAnoR$), signal: inv.rentabilidadeAnoR$, seed: 3 },
                { label: "ROI no ano", value: fmtPct(inv.roiAnoPct), signal: inv.roiAnoPct, seed: 4 },
              ].map((tile, i) => (
                <div key={i} style={{ background: T.subCardBg || T.cardBorder, borderRadius: 10, padding: "8px 7px 7px", minWidth: 0 }}>
                  <div style={{ fontSize: 8.5, color: T.textMuted, marginBottom: 5, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{tile.label}</div>
                  <div style={{
                    fontFamily: "'Tomorrow', sans-serif", fontSize: 12, fontWeight: 400,
                    color: trendColor(tile.signal, T), marginBottom: 5,
                    whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis",
                    ...blurIf(hideValues),
                  }}>
                    {tile.value}
                  </div>
                  <MiniSparkline color={trendColor(tile.signal, T)} seed={tile.seed} width={48} height={18} />
                </div>
              ))}
            </div>

            {/* Spread: mês vs média mensal do ano */}
            {inv.roiMesPct != null && inv.roiAnoPct != null && !Number.isNaN(inv.roiMesPct) && !Number.isNaN(inv.roiAnoPct) && (() => {
              const mediaAnualMensal = inv.roiAnoPct / 12;
              const spread = inv.roiMesPct - mediaAnualMensal;
              const acima = spread > 0;
              return (
                <div style={{
                  paddingTop: 12, borderTop: `1px solid ${T.cardBorder}`,
                  display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8,
                }}>
                  <span style={{ fontSize: 11, color: T.textMuted }}>Este mês vs. média do ano</span>
                  <span style={{
                    display: "flex", alignItems: "center", gap: 3, fontSize: 11.5, fontWeight: 700,
                    color: trendColor(spread, T), background: trendBg(spread, T),
                    padding: "3px 8px", borderRadius: 999, whiteSpace: "nowrap",
                  }}>
                    {acima ? <TrendingUp size={11} /> : spread < 0 ? <TrendingDown size={11} /> : null}
                    {acima ? "+" : ""}{fmtPct(spread)}
                  </span>
                </div>
              );
            })()}
          </CardWrapper>
        )}

        {/* Faturamento (tbFaturamento) */}
        {fat && fat.totalR$ > 0 && (
          <CardWrapper style={{ padding: "16px 16px 14px", marginBottom: 14 }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
              <span style={{ fontSize: 13, fontWeight: 600, color: T.text }}>Faturamento</span>
              {fat.refMes && (
                <span style={{
                  fontSize: 10.5, fontWeight: 700, color: "#fff", background: T.accent,
                  padding: "4px 9px", borderRadius: 999, letterSpacing: 0.4, textTransform: "uppercase",
                }}>
                  {new Date(fat.refMes).toLocaleDateString("pt-BR", { month: "short", year: "numeric" }).replace(".", "").replace(" de ", " ")}
                </span>
              )}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 14 }}>
              <div>
                <div style={{ fontSize: 11.5, color: T.textMuted, marginBottom: 4 }}>Total</div>
                <div style={{ display: "flex", alignItems: "center", gap: 8, ...blurIf(hideValues) }}>
                  <div style={{ width: 26, height: 26, borderRadius: 8, background: `${T.text}22`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                    <Banknote size={14} color={T.text} />
                  </div>
                  <span style={{ fontFamily: "'Tomorrow', sans-serif", fontSize: 18, fontWeight: 400 }}>
                    {fmtBRL(fat.totalR$)}
                  </span>
                </div>
              </div>
              <div>
                <div style={{ fontSize: 11.5, color: T.textMuted, marginBottom: 4 }}>Investido</div>
                <div style={{ display: "flex", alignItems: "center", gap: 8, ...blurIf(hideValues) }}>
                  <div style={{ width: 26, height: 26, borderRadius: 8, background: `${T.positive}22`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                    <PiggyBank size={14} color={T.positive} />
                  </div>
                  <span style={{ fontFamily: "'Tomorrow', sans-serif", fontSize: 18, fontWeight: 400, color: T.positive }}>
                    {fmtBRL(fat.investidoR$)}
                  </span>
                </div>
              </div>
            </div>
            {/* Apenas percentual, sem barra de progressão */}
            <div style={{ fontSize: 12, color: T.textMuted, ...blurIf(hideValues) }}>
              {fmtPctNoDecimal(fat.investidoPct)} do faturamento investido
            </div>
          </CardWrapper>
        )}

      </div>
    );
  };

  // ---------- DOAÇÕES ----------
  const renderDonations = () => (
    <div key="donations" className="pdb-fade-in">
      <SectionHeader T={T} title="Doações" hideValues={hideValues} onToggleHide={() => setHideValues((v) => !v)} />

      <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-start", padding: "4px 4px 18px" }}>
        <span style={{ fontSize: 12, color: T.textMuted, fontWeight: 500, marginBottom: 8 }}>
          Doado em {resumo.anoReferencia}
        </span>
        <span style={{
          fontFamily: "'Tomorrow', sans-serif", fontSize: 36, fontWeight: 400, letterSpacing: -1,
          lineHeight: 1.05, ...blurIf(hideValues),
        }}>
          {fmtBRL(animatedDoado)}
        </span>
        <div style={{ width: "100%", marginTop: 18 }}>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: T.textMuted, marginBottom: 10, ...blurIf(hideValues) }}>
            <span style={{ fontWeight: 600, color: T.text }}>{fmtPct(metaAtingida)} da meta</span>
            <span>{fmtBRL(resumo.metaDoacoes)}</span>
          </div>
          <style>{`
            @keyframes donationBarFill {
              from { width: 0%; }
              to { width: ${Math.min(100, metaAtingida * 100)}%; }
            }
            @keyframes donationBarShimmer {
              0% { background-position: 0% 0; }
              100% { background-position: 200% 0; }
            }
            .pdb-donation-bar-track {
              position: relative;
              height: 18px;
              border-radius: 999px;
              background: ${T.cardBorder};
              overflow: hidden;
            }
            .pdb-donation-bar-fill {
              position: relative;
              height: 100%;
              border-radius: 999px;
              width: ${Math.min(100, metaAtingida * 100)}%;
              animation: donationBarFill 1.1s cubic-bezier(0.22, 1, 0.36, 1) both;
              background: linear-gradient(90deg, #023E8A 0%, #0077B6 45%, #00B4D8 100%);
              box-shadow: 0 0 16px rgba(0,180,216,0.55), 0 0 4px rgba(2,62,138,0.4);
              overflow: hidden;
            }
            .pdb-donation-bar-fill::before {
              content: '';
              position: absolute;
              inset: 0;
              background: linear-gradient(
                110deg,
                transparent 20%,
                rgba(255,255,255,0.35) 42%,
                rgba(255,255,255,0.55) 48%,
                rgba(255,255,255,0.35) 54%,
                transparent 76%
              );
              background-size: 200% 100%;
              animation: donationBarShimmer 2.6s ease-in-out 1.1s infinite;
            }
            .pdb-donation-bar-tip {
              position: absolute;
              top: 50%;
              right: 1px;
              width: 10px;
              height: 10px;
              border-radius: 50%;
              background: #ffffff;
              transform: translate(50%, -50%);
              box-shadow: 0 0 8px rgba(0,180,216,0.9), 0 0 2px rgba(255,255,255,0.8);
              opacity: ${metaAtingida > 0 ? 1 : 0};
            }
          `}</style>
          <div className="pdb-donation-bar-track">
            <div className="pdb-donation-bar-fill">
              <div className="pdb-donation-bar-tip" />
            </div>
          </div>
        </div>
      </div>

      <div style={{ display: "flex", gap: 10, marginBottom: 14 }}>
        <StatCard T={T} icon={HeartHandshake} label="Meta anual" accent={T.amber} hidden={hideValues}
          value={fmtBRL(resumo.metaDoacoes)} />
        <StatCard T={T} icon={TrendingUp} label="Restante p/ meta" accent={T.accent} hidden={hideValues}
          value={fmtBRL(restanteMeta)} />
      </div>

      {/* Distribuição das doações (tbDonationsDistribuition) — donut centralizado + lista com ícones */}
      {donationsDistribuicao?.length > 0 ? (
        <CardWrapper style={{ padding: "16px 16px 16px", marginBottom: 14 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 18 }}>
            <span style={{ fontSize: 13, fontWeight: 600, color: T.text }}>Distribuição das doações</span>
            <Info size={12} color={T.textMuted} />
          </div>
          <div style={{ display: "flex", justifyContent: "center", marginBottom: 22 }}>
            <AllocationDonut
              T={T}
              data={donationsDistribuicao.map((d) => ({
                categoria: d.categoria,
                percentual: d.percentual != null && !Number.isNaN(d.percentual) ? d.percentual : (donationsTotal > 0 ? d.valor / donationsTotal : 0),
              }))}
              colorMap={DONATION_COLORS[theme]}
              uid="donationsDonut"
              centerLabel="Doado no ano"
              centerValue={fmtBRL(donationsTotal)}
              hidden={hideValues}
              size={172}
            />
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {(() => {
              const withPct = donationsDistribuicao.map((entry) => ({
                ...entry,
                pct: entry.percentual != null && !Number.isNaN(entry.percentual) ? entry.percentual : (donationsTotal > 0 ? entry.valor / donationsTotal : 0),
              }));
              const maxPct = Math.max(...withPct.map((e) => e.pct));
              return withPct.map((entry, i) => {
                const DonIcon = DONATION_ICONS[entry.categoria] || HeartHandshake;
                const color = DONATION_COLORS[theme][entry.categoria] || T.textMuted;
                const isTop = entry.pct === maxPct && maxPct > 0;
                return (
                  <div key={i} style={{
                    position: "relative", overflow: "hidden", borderRadius: 14,
                    background: `linear-gradient(135deg, ${color}1c 0%, transparent 70%)`,
                    border: `1px solid ${T.cardBorder}`,
                    padding: "12px 14px",
                  }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                      <div style={{
                        width: 32, height: 32, borderRadius: 10, background: `${color}26`,
                        display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
                      }}>
                        <DonIcon size={16} color={color} />
                      </div>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 2 }}>
                          <span style={{ fontSize: 12.5, color: T.text, fontWeight: 600, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                            {entry.categoria}
                          </span>
                          {isTop && (
                            <span style={{
                              fontSize: 8.5, fontWeight: 700, color, background: `${color}22`,
                              padding: "2px 6px", borderRadius: 999, letterSpacing: 0.3, textTransform: "uppercase", whiteSpace: "nowrap",
                            }}>
                              Maior doação
                            </span>
                          )}
                        </div>
                        <span style={{ fontSize: 10.5, color: T.textMuted, ...blurIf(hideValues) }}>
                          {fmtBRL(entry.valor)}
                        </span>
                      </div>
                      <span style={{
                        fontFamily: "'Tomorrow', sans-serif", fontSize: 20, fontWeight: 400,
                        color: T.text, letterSpacing: -0.5, flexShrink: 0, ...blurIf(hideValues),
                      }}>
                        {fmtPct(entry.pct)}
                      </span>
                    </div>
                    <div style={{ height: 4, borderRadius: 2, background: T.cardBorder, overflow: "hidden", marginTop: 10 }}>
                      <div style={{ width: `${Math.min(100, entry.pct * 100)}%`, height: "100%", borderRadius: 2, background: color, transition: "width 0.6s ease" }} />
                    </div>
                  </div>
                );
              });
            })()}
          </div>
        </CardWrapper>
      ) : (
        <CardWrapper style={{ padding: "14px 16px", marginBottom: 14 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
            <AlertCircle size={14} color={T.textMuted} />
            <span style={{ fontSize: 12.5, fontWeight: 600, color: T.text }}>Distribuição das doações indisponível</span>
          </div>
          <span style={{ fontSize: 11.5, color: T.textMuted, lineHeight: 1.5 }}>
            Não encontrei a tabela "tbDonationsDistribuition" (Mês/Valor) na aba API desta planilha. Confira se ela está preenchida e reenvie o arquivo.
          </span>
        </CardWrapper>
      )}

      {resumo.reservaDoacoes > 0 && (
        <CardWrapper style={{ padding: "14px 16px" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
            <PiggyBank size={15} color={CATEGORY_COLORS.Fundos} />
            <span style={{ fontSize: 12.5, fontWeight: 600, color: T.text }}>Reserva de doações</span>
          </div>
          <div style={{ fontFamily: "'Tomorrow', sans-serif", fontSize: 18, fontWeight: 700, ...blurIf(hideValues) }}>
            {fmtBRL(resumo.reservaDoacoes)}
          </div>
        </CardWrapper>
      )}
    </div>
  );

  // ---------- CONFIGURAÇÕES ----------
  const renderSettings = () => (
    <div key="settings" className="pdb-fade-in">
      <SectionHeader T={T} title="Configurações" hideValues={hideValues} onToggleHide={() => setHideValues((v) => !v)} />

      <CardWrapper style={{ padding: "18px 16px", marginBottom: 12, display: "flex", alignItems: "center", gap: 14 }}>
        {avatarNode(58)}
        <div style={{ minWidth: 0 }}>
          <div style={{ fontFamily: "'Tomorrow', sans-serif", fontSize: 15, fontWeight: 700, whiteSpace: "nowrap" }}>{NAME}</div>
          <div style={{ fontSize: 12, color: T.textMuted, marginTop: 2 }}>Toque na foto para alterar</div>
        </div>
      </CardWrapper>

      <CardWrapper style={{ padding: "4px 16px", marginBottom: 12 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 0" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            {theme === "dark" ? <Moon size={16} color={T.textMuted} /> : <Sun size={16} color={T.textMuted} />}
            <span style={{ fontSize: 13.5, fontWeight: 500 }}>Tema {theme === "dark" ? "escuro" : "claro"}</span>
          </div>
          <button onClick={toggleTheme} style={{
            width: 46, height: 26, borderRadius: 13, position: "relative", border: "none", cursor: "pointer",
            background: theme === "dark" ? T.accent : T.cardBorder, transition: "background 0.25s ease", flexShrink: 0,
          }}>
            <span style={{
              position: "absolute", top: 3, left: theme === "dark" ? 23 : 3, width: 20, height: 20, borderRadius: "50%",
              background: "#fff", transition: "left 0.25s ease", boxShadow: "0 2px 5px rgba(0,0,0,0.25)",
            }} />
          </button>
        </div>
      </CardWrapper>

      <CardWrapper style={{ padding: "14px 16px", marginBottom: 12, cursor: "pointer" }}>
        <label style={{ display: "flex", alignItems: "center", justifyContent: "space-between", cursor: "pointer" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            {status === "parsing" ? (
              <Loader2 size={16} color={T.textMuted} style={{ animation: "spin 1s linear infinite" }} />
            ) : (
              <RefreshCw size={16} color={T.textMuted} />
            )}
            <span style={{ fontSize: 13.5, fontWeight: 500 }}>Reenviar planilha</span>
          </div>
          <ChevronRight size={16} color={T.textMuted} />
          <input type="file" accept=".xlsx,.xls" style={{ display: "none" }}
            onChange={(e) => handleFile(e.target.files[0])} />
        </label>
      </CardWrapper>

      {errorMsg && (
        <div style={{ display: "flex", gap: 6, alignItems: "center", color: T.negative, fontSize: 12, background: T.negativeBg, padding: "8px 12px", borderRadius: 10, marginBottom: 12 }}>
          <AlertCircle size={13} />{errorMsg}
        </div>
      )}

      <div style={{ textAlign: "center", marginTop: 8 }}>
        <div style={{ fontSize: 11, color: T.textMuted }}>
          Atualizado em {fmtDateFull(ultimaAtualizacao)}
        </div>
        <div style={{ marginTop: 14, paddingTop: 14, borderTop: `1px solid ${T.cardBorder}` }}>
          <div style={{ fontFamily: "'Tomorrow', sans-serif", fontSize: 11, color: T.textMuted, letterSpacing: 0.5, opacity: 0.7 }}>
            Patrimônio Dashboard
          </div>
          <div style={{ fontSize: 10, color: T.textMuted, marginTop: 4, opacity: 0.55 }}>
            Desenvolvido por Alexsandro Torres com Claude AI
          </div>
          <div style={{ fontSize: 9.5, color: T.textMuted, marginTop: 6, opacity: 0.45 }}>
            versão de teste · v0.9.3
          </div>
        </div>
      </div>
    </div>
  );

  const tabRenderers = { home: renderHome, invest: renderInvest, donations: renderDonations, settings: renderSettings };

  return shell(
    <div key={activeTab} className={navDirection >= 0 ? "pdb-tab-slide-fwd" : "pdb-tab-slide-back"}>
      {tabRenderers[activeTab]()}
    </div>,
    true
  );
}
