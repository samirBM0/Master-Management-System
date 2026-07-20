import React, { useMemo } from "react";
import * as d3 from "d3";

interface RadarChartProps {
  cp: number;
  cpk: number;
  centering: number;
  sigma: number;
  range: number;
  lsl: number;
  usl: number;
  nominal: number;
  width?: number;
  height?: number;
  theme?: "light" | "dark";
}

export default function RadarChart({
  cp,
  cpk,
  centering,
  sigma,
  range,
  lsl,
  usl,
  nominal,
  width = 240,
  height = 240,
  theme = "light",
}: RadarChartProps) {
  const rangeTol = usl - lsl;

  // Normalize metrics to 0-100 scores
  const scoreCp = useMemo(() => {
    if (cp === Infinity || cp > 1000) return 100;
    if (isNaN(cp) || cp <= 0) return 0;
    return Math.min(100, Math.max(0, (cp / 2) * 100));
  }, [cp]);

  const scoreCpk = useMemo(() => {
    if (cpk === Infinity || cpk > 1000) return 100;
    if (isNaN(cpk) || cpk <= -50) return 0;
    return Math.min(100, Math.max(0, (cpk / 2) * 100));
  }, [cpk]);

  const scoreCentering = useMemo(() => {
    if (rangeTol <= 0 || isNaN(centering)) return 100;
    return Math.min(100, Math.max(0, (1 - Math.abs(centering)) * 100));
  }, [centering, rangeTol]);

  const scoreSigma = useMemo(() => {
    if (rangeTol <= 0 || isNaN(sigma)) return 100;
    if (sigma === 0) return 100;
    const maxAllowedSigma = rangeTol / 6;
    return Math.min(100, Math.max(0, (1 - sigma / (2 * maxAllowedSigma)) * 100));
  }, [sigma, rangeTol]);

  const scoreRange = useMemo(() => {
    if (rangeTol <= 0 || isNaN(range)) return 100;
    return Math.min(100, Math.max(0, (1 - range / rangeTol) * 100));
  }, [range, rangeTol]);

  // Set up D3 config
  const padding = 45;
  const radius = Math.min(width, height) / 2 - padding;
  const cx = width / 2;
  const cy = height / 2;

  // Five axes config
  const axes = [
    { name: "Cp (Potentiel)", score: scoreCp, displayVal: cp === Infinity || cp > 1000 ? "N/A" : cp.toFixed(2) },
    { name: "Cpk (Réel)", score: scoreCpk, displayVal: cpk === Infinity || cpk > 1000 ? "N/A" : cpk.toFixed(2) },
    { name: "Centrage", score: scoreCentering, displayVal: isNaN(centering) ? "N/A" : (100 * (1 - Math.abs(centering))).toFixed(0) + "%" },
    { name: "Sigma (σ)", score: scoreSigma, displayVal: isNaN(sigma) ? "N/A" : sigma.toFixed(4) },
    { name: "Range (R)", score: scoreRange, displayVal: isNaN(range) ? "N/A" : range.toFixed(3) },
  ];

  // D3 Scales and line generators
  const rScale = d3.scaleLinear().domain([0, 100]).range([0, radius]);

  const angleScale = d3
    .scaleLinear()
    .domain([0, axes.length])
    .range([0, 2 * Math.PI]);

  // Generate background concentric rings
  const ringLevels = [20, 40, 60, 80, 100];

  const ringPaths = useMemo(() => {
    return ringLevels.map((level) => {
      const r = rScale(level);
      const points = axes.map((_, i) => {
        const angle = angleScale(i) - Math.PI / 2;
        return {
          x: cx + r * Math.cos(angle),
          y: cy + r * Math.sin(angle),
        };
      });
      return d3.line<{ x: number; y: number }>()
        .x((d) => d.x)
        .y((d) => d.y)
        .curve(d3.curveLinearClosed)(points) || "";
    });
  }, [cx, cy, rScale, angleScale, axes]);

  // Generate axes lines
  const axisLines = useMemo(() => {
    return axes.map((_, i) => {
      const angle = angleScale(i) - Math.PI / 2;
      const targetR = rScale(100);
      return {
        x1: cx,
        y1: cy,
        x2: cx + targetR * Math.cos(angle),
        y2: cy + targetR * Math.sin(angle),
        labelX: cx + (targetR + 15) * Math.cos(angle),
        labelY: cy + (targetR + 12) * Math.sin(angle),
        angle,
      };
    });
  }, [cx, cy, rScale, angleScale, axes]);

  // Generate the active data polygon
  const activePolygonPath = useMemo(() => {
    const points = axes.map((axis, i) => {
      const angle = angleScale(i) - Math.PI / 2;
      const r = rScale(axis.score);
      return {
        x: cx + r * Math.cos(angle),
        y: cy + r * Math.sin(angle),
      };
    });
    return d3.line<{ x: number; y: number }>()
      .x((d) => d.x)
      .y((d) => d.y)
      .curve(d3.curveLinearClosed)(points) || "";
  }, [cx, cy, rScale, angleScale, axes]);

  // Vertex points
  const vertexPoints = useMemo(() => {
    return axes.map((axis, i) => {
      const angle = angleScale(i) - Math.PI / 2;
      const r = rScale(axis.score);
      return {
        x: cx + r * Math.cos(angle),
        y: cy + r * Math.sin(angle),
        score: axis.score,
        name: axis.name,
      };
    });
  }, [cx, cy, rScale, angleScale, axes]);

  // Color theme definitions
  const isDark = theme === "dark";
  const ringColor = isDark ? "rgba(255,255,255,0.08)" : "rgba(15,23,42,0.08)";
  const ringStroke = isDark ? "rgba(255,255,255,0.15)" : "rgba(15,23,42,0.12)";
  const textColor = isDark ? "#cbd5e1" : "#334155";
  const textSubColor = isDark ? "#64748b" : "#475569";
  const polygonFill = "rgba(225,29,72,0.18)"; // ACTIA Rose
  const polygonStroke = "rgb(225,29,72)";
  const vertexColor = "rgb(225,29,72)";

  return (
    <div className="flex flex-col items-center justify-center">
      <svg width={width} height={height} className="overflow-visible">
        {/* Background rings */}
        {ringPaths.map((path, idx) => (
          <path
            key={idx}
            d={path}
            fill="none"
            stroke={ringStroke}
            strokeWidth="0.8"
            strokeDasharray={idx < 4 ? "2,2" : "none"}
          />
        ))}

        {/* Ring Labels */}
        {ringLevels.map((level) => (
          <text
            key={level}
            x={cx + 3}
            y={cy - rScale(level) + 3}
            fill={isDark ? "rgba(255,255,255,0.3)" : "rgba(15,23,42,0.3)"}
            fontSize="7"
            fontFamily="monospace"
          >
            {level}%
          </text>
        ))}

        {/* Axis Lines */}
        {axisLines.map((line, idx) => (
          <line
            key={idx}
            x1={line.x1}
            y1={line.y1}
            x2={line.x2}
            y2={line.y2}
            stroke={ringColor}
            strokeWidth="1.2"
          />
        ))}

        {/* Active Data Polygon */}
        <path
          d={activePolygonPath}
          fill={polygonFill}
          stroke={polygonStroke}
          strokeWidth="2.2"
          strokeLinecap="round"
          strokeLinejoin="round"
        />

        {/* Axis labels & scores */}
        {axisLines.map((line, idx) => {
          const axis = axes[idx];
          // Dynamic alignment anchor based on position to avoid cutoff
          let textAnchor = "middle";
          const cos = Math.cos(line.angle);
          if (cos > 0.1) textAnchor = "start";
          else if (cos < -0.1) textAnchor = "end";

          // Fine tune label offsets
          const offsetY = idx === 0 ? -4 : idx === 2 || idx === 3 ? 12 : 2;

          return (
            <g key={idx}>
              <text
                x={line.labelX}
                y={line.labelY + offsetY}
                fill={textColor}
                fontSize="9"
                fontWeight="800"
                fontFamily="system-ui"
                textAnchor={textAnchor}
              >
                {axis.name}
              </text>
              <text
                x={line.labelX}
                y={line.labelY + offsetY + 10}
                fill={textSubColor}
                fontSize="8"
                fontWeight="bold"
                fontFamily="monospace"
                textAnchor={textAnchor}
              >
                {axis.displayVal}
              </text>
            </g>
          );
        })}

        {/* Vertex Points */}
        {vertexPoints.map((point, idx) => (
          <g key={idx}>
            <circle
              cx={point.x}
              cy={point.y}
              r="4"
              fill={vertexColor}
              stroke="#ffffff"
              strokeWidth="1.2"
              className="shadow-sm"
            />
            <title>{`${point.name}: ${point.score.toFixed(1)}%`}</title>
          </g>
        ))}

        {/* Center Point */}
        <circle cx={cx} cy={cy} r="2" fill={textColor} />
      </svg>
    </div>
  );
}
