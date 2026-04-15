import React, { useEffect, useState, useRef, useCallback } from "react";
import { Line } from "react-chartjs-2";
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  PointElement,
  LineElement,
  Title,
  Tooltip,
  Legend,
  Filler,
} from "chart.js";
import * as XLSX from "xlsx";

ChartJS.register(
  CategoryScale,
  LinearScale,
  PointElement,
  LineElement,
  Title,
  Tooltip,
  Legend,
  Filler
);

function App() {
  const [sound, setSound] = useState(0);
  const [temp, setTemp] = useState(0);
  const [hum, setHum] = useState(0);
  const [motion, setMotion] = useState(false);
  const [light, setLight] = useState(false);
  const [connected, setConnected] = useState(false);

  const [soundTotal, setSoundTotal] = useState(0);
  const [tempTotal, setTempTotal] = useState(0);
  const [humTotal, setHumTotal] = useState(0);
  const [motionTotal, setMotionTotal] = useState(0);

  const soundData  = useRef([]);
  const tempData   = useRef([]);
  const humData    = useRef([]);
  const motionData = useRef([]);
  const lightData  = useRef([]);
  const labels     = useRef([]);

  // Keep ALL historical data for Excel export (not limited to 30)
  const allLabels     = useRef([]);
  const allSoundData  = useRef([]);
  const allTempData   = useRef([]);
  const allHumData    = useRef([]);
  const allMotionData = useRef([]);
  const allLightData  = useRef([]);

  const [, forceUpdate] = useState(0);

  useEffect(() => {
    const ws = new WebSocket("ws://192.168.43.177:81");

    ws.onopen  = () => setConnected(true);
    ws.onclose = () => setConnected(false);
    ws.onerror = () => setConnected(false);

    ws.onmessage = (event) => {
      const data = JSON.parse(event.data);

      setSound(data.sound);
      setTemp(data.temp);
      setHum(data.hum);
      setMotion(data.motion);
      setLight(data.light);

      setSoundTotal((prev) => prev + data.sound);
      setTempTotal((prev)  => prev + data.temp);
      setHumTotal((prev)   => prev + data.hum);
      setMotionTotal((prev) => prev + (data.motion ? 1 : 0));

      const now = new Date().toLocaleTimeString();

      // Chart data (last 30)
      if (labels.current.length >= 30)     labels.current.shift();
      if (soundData.current.length >= 30)  soundData.current.shift();
      if (tempData.current.length >= 30)   tempData.current.shift();
      if (humData.current.length >= 30)    humData.current.shift();
      if (motionData.current.length >= 30) motionData.current.shift();
      if (lightData.current.length >= 30)  lightData.current.shift();

      labels.current.push(now);
      soundData.current.push(data.sound);
      tempData.current.push(data.temp);
      humData.current.push(data.hum);
      motionData.current.push(data.motion ? 1 : 0);
      lightData.current.push(data.light ? 1 : 0);

      // Full history for Excel export
      allLabels.current.push(now);
      allSoundData.current.push(data.sound);
      allTempData.current.push(data.temp);
      allHumData.current.push(data.hum);
      allMotionData.current.push(data.motion ? 1 : 0);
      allLightData.current.push(data.light ? 1 : 0);

      forceUpdate((n) => n + 1);
    };

    return () => ws.close();
  }, []);

  const formatNumber = (num) => num.toLocaleString();
  const grandTotal = soundTotal + tempTotal + humTotal + motionTotal;

  // ---------- Excel Export (unchanged) ----------
  const downloadExcel = useCallback(() => {
    const wb = XLSX.utils.book_new();

    const summaryData = [
      ["Room Monitor Report"],
      ["Generated", new Date().toLocaleString()],
      ["Total Readings", allLabels.current.length],
      [],
      ["Sensor", "Cumulative Total"],
      ["Sound",    soundTotal],
      ["Temperature", tempTotal],
      ["Humidity", humTotal],
      ["Motion Events", motionTotal],
      [],
      ["Grand Total", grandTotal],
    ];
    const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
    wsSummary["!cols"]   = [{ wch: 22 }, { wch: 25 }];
    wsSummary["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
    XLSX.utils.book_append_sheet(wb, wsSummary, "Summary");

    const rawHeader = ["Time", "Sound", "Temperature (°C)", "Humidity (%)", "Motion (1/0)", "Light (1/0)"];
    const rawRows = allLabels.current.map((label, i) => [
      label,
      allSoundData.current[i],
      allTempData.current[i],
      allHumData.current[i],
      allMotionData.current[i],
      allLightData.current[i],
    ]);
    const wsData = XLSX.utils.aoa_to_sheet([rawHeader, ...rawRows]);
    wsData["!cols"] = [
      { wch: 14 }, { wch: 10 }, { wch: 18 },
      { wch: 14 }, { wch: 14 }, { wch: 14 },
    ];
    XLSX.utils.book_append_sheet(wb, wsData, "Sensor Data");

    const calcStats = (arr) => {
      if (arr.length === 0) return { min: 0, max: 0, avg: 0 };
      const min = Math.min(...arr);
      const max = Math.max(...arr);
      const avg = arr.reduce((a, b) => a + b, 0) / arr.length;
      return { min, max, avg: Math.round(avg * 100) / 100 };
    };

    const soundStats   = calcStats(allSoundData.current);
    const tempStats    = calcStats(allTempData.current);
    const humStats     = calcStats(allHumData.current);
    const motionCount  = allMotionData.current.filter((v) => v === 1).length;
    const lightOnCount = allLightData.current.filter((v) => v === 1).length;

    const statsData = [
      ["Sensor Statistics"],
      [],
      ["Metric", "Sound", "Temperature (°C)", "Humidity (%)"],
      ["Minimum", soundStats.min, tempStats.min, humStats.min],
      ["Maximum", soundStats.max, tempStats.max, humStats.max],
      ["Average", soundStats.avg, tempStats.avg, humStats.avg],
      [],
      ["Binary Sensors", "Count", "Out of Total", "Percentage"],
      ["Motion Detected", motionCount, allMotionData.current.length,
        allMotionData.current.length > 0
          ? Math.round((motionCount / allMotionData.current.length) * 10000) / 100 + "%"
          : "0%"
      ],
      ["Light On (Bright)", lightOnCount, allLightData.current.length,
        allLightData.current.length > 0
          ? Math.round((lightOnCount / allLightData.current.length) * 10000) / 100 + "%"
          : "0%"
      ],
    ];
    const wsStats = XLSX.utils.aoa_to_sheet(statsData);
    wsStats["!cols"]   = [{ wch: 20 }, { wch: 14 }, { wch: 18 }, { wch: 14 }];
    wsStats["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }];
    XLSX.utils.book_append_sheet(wb, wsStats, "Statistics");

    const timestamp = new Date().toISOString().slice(0, 19).replace(/[T:]/g, "-");
    XLSX.writeFile(wb, `Room_Monitor_Report_${timestamp}.xlsx`);
  }, [soundTotal, tempTotal, humTotal, motionTotal, grandTotal]);

  // ---------- Single combined chart ----------
  // Sound on its own right Y-axis (0–5000).
  // Temp + Hum on left Y-axis (0–100).
  // Motion and Light are boolean 0/1, scaled ×100 so they hit
  // the top/bottom of the left axis as filled step bands.
  const combinedChartData = {
    labels: labels.current,
    datasets: [
      {
        label: "Sound",
        data: soundData.current,
        borderColor: "#e74c3c",
        backgroundColor: "rgba(231,76,60,0.05)",
        fill: false,
        tension: 0.4,
        yAxisID: "ySound",
        pointRadius: 0,
        borderWidth: 2,
      },
      {
        label: "Temp (°C)",
        data: tempData.current,
        borderColor: "#e67e22",
        fill: false,
        tension: 0.4,
        yAxisID: "yLeft",
        pointRadius: 0,
        borderWidth: 2,
      },
      {
        label: "Humidity (%)",
        data: humData.current,
        borderColor: "#3498db",
        fill: false,
        tension: 0.4,
        yAxisID: "yLeft",
        pointRadius: 0,
        borderWidth: 2,
      },
      {
        label: "Motion",
        data: motionData.current.map((v) => v * 100),
        borderColor: "rgba(155,89,182,0.8)",
        backgroundColor: "rgba(155,89,182,0.12)",
        fill: true,
        stepped: true,
        tension: 0,
        yAxisID: "yLeft",
        pointRadius: 0,
        borderWidth: 1.5,
      },
      {
        label: "Light",
        data: lightData.current.map((v) => v * 100),
        borderColor: "rgba(241,196,15,0.9)",
        backgroundColor: "rgba(241,196,15,0.12)",
        fill: true,
        stepped: true,
        tension: 0,
        yAxisID: "yLeft",
        pointRadius: 0,
        borderWidth: 1.5,
      },
    ],
  };

  const combinedOptions = {
    responsive: true,
    maintainAspectRatio: false,
    animation: false,
    interaction: { mode: "index", intersect: false },
    plugins: {
      legend: {
        display: true,
        position: "top",
        labels: { boxWidth: 12, padding: 14, font: { size: 12 } },
      },
      tooltip: {
        callbacks: {
          label: (ctx) => {
            const ds = ctx.dataset.label;
            const v  = ctx.parsed.y;
            if (ds === "Motion") return `Motion: ${v === 100 ? "Detected" : "Clear"}`;
            if (ds === "Light")  return `Light: ${v === 100 ? "Bright" : "Dark"}`;
            if (ds === "Sound")  return `Sound: ${v}`;
            return `${ds}: ${v}`;
          },
        },
      },
    },
    scales: {
      yLeft: {
        type: "linear",
        position: "left",
        min: 0,
        max: 100,
        title: { display: true, text: "Temp (°C) / Humidity (%)" },
        grid: { color: "rgba(0,0,0,0.06)" },
        ticks: {
          callback: (v) => {
            if (v === 100) return "100 / ON";
            if (v === 0)   return "0 / OFF";
            return v;
          },
        },
      },
      ySound: {
        type: "linear",
        position: "right",
        min: 0,
        suggestedMax: 5000,
        title: { display: true, text: "Sound Level" },
        grid: { drawOnChartArea: false },
        ticks: { color: "#e74c3c" },
      },
      x: {
        display: true,
        ticks: { maxTicksLimit: 6, maxRotation: 0, font: { size: 10 } },
        title: { display: true, text: "Time" },
        grid: { display: false },
      },
    },
  };

  return (
    <div style={{ fontFamily: "sans-serif", maxWidth: "800px", margin: "0 auto", padding: "20px" }}>
      <h1 style={{ textAlign: "center" }}>Room Monitor</h1>

      {/* Connection status */}
      <div style={{ textAlign: "center", marginBottom: "15px" }}>
        <span style={{
          display: "inline-block", width: "10px", height: "10px",
          borderRadius: "50%",
          backgroundColor: connected ? "#44cc44" : "#cc4444",
          marginRight: "6px",
        }} />
        {connected ? "Connected" : "Disconnected"}
      </div>

      {/* Grand Total */}
      <div style={{
        textAlign: "center", marginBottom: "12px",
        padding: "16px 20px", borderRadius: "10px",
        backgroundColor: "#1a1a2e", color: "white",
      }}>
        <div style={{ fontSize: "0.85rem", opacity: 0.7, marginBottom: "4px" }}>
          Total Data Collected (All Sensors)
        </div>
        <div style={{ fontSize: "2.2rem", fontWeight: "bold", letterSpacing: "1px" }}>
          {formatNumber(grandTotal)}
        </div>
      </div>

      {/* Per-sensor totals */}
      <div style={{ display: "flex", justifyContent: "center", gap: "10px", flexWrap: "wrap", marginBottom: "20px" }}>
        <div style={{ flex: "1 1 120px", padding: "10px 14px", borderRadius: "8px", backgroundColor: "rgba(231,76,60,0.1)", textAlign: "center", border: "1px solid rgba(231,76,60,0.3)" }}>
          <div style={{ fontSize: "0.75rem", color: "#e74c3c" }}>Sound</div>
          <div style={{ fontSize: "1.2rem", fontWeight: "bold" }}>{formatNumber(soundTotal)}</div>
        </div>
        <div style={{ flex: "1 1 120px", padding: "10px 14px", borderRadius: "8px", backgroundColor: "rgba(230,126,34,0.1)", textAlign: "center", border: "1px solid rgba(230,126,34,0.3)" }}>
          <div style={{ fontSize: "0.75rem", color: "#e67e22" }}>Temperature</div>
          <div style={{ fontSize: "1.2rem", fontWeight: "bold" }}>{formatNumber(tempTotal)}</div>
        </div>
        <div style={{ flex: "1 1 120px", padding: "10px 14px", borderRadius: "8px", backgroundColor: "rgba(52,152,219,0.1)", textAlign: "center", border: "1px solid rgba(52,152,219,0.3)" }}>
          <div style={{ fontSize: "0.75rem", color: "#3498db" }}>Humidity</div>
          <div style={{ fontSize: "1.2rem", fontWeight: "bold" }}>{formatNumber(humTotal)}</div>
        </div>
        <div style={{ flex: "1 1 120px", padding: "10px 14px", borderRadius: "8px", backgroundColor: "rgba(155,89,182,0.1)", textAlign: "center", border: "1px solid rgba(155,89,182,0.3)" }}>
          <div style={{ fontSize: "0.75rem", color: "#9b59b6" }}>Motion</div>
          <div style={{ fontSize: "1.2rem", fontWeight: "bold" }}>{formatNumber(motionTotal)}</div>
        </div>
      </div>

      {/* Download Button */}
      <div style={{ textAlign: "center", marginBottom: "20px" }}>
        <button
          onClick={downloadExcel}
          disabled={allLabels.current.length === 0}
          style={{
            padding: "12px 28px", fontSize: "1rem", fontWeight: "bold",
            color: "white",
            backgroundColor: allLabels.current.length === 0 ? "#95a5a6" : "#27ae60",
            border: "none", borderRadius: "8px",
            cursor: allLabels.current.length === 0 ? "not-allowed" : "pointer",
            boxShadow: "0 2px 8px rgba(39,174,96,0.3)",
            transition: "background-color 0.2s",
          }}
          onMouseEnter={(e) => { if (allLabels.current.length > 0) e.target.style.backgroundColor = "#219a52"; }}
          onMouseLeave={(e) => { if (allLabels.current.length > 0) e.target.style.backgroundColor = "#27ae60"; }}
        >
          📥 Download Excel Report ({allLabels.current.length} readings)
        </button>
      </div>

      {/* Live values (unchanged from original) */}
      <div style={{ display: "flex", justifyContent: "center", gap: "15px", flexWrap: "wrap", marginBottom: "20px" }}>
        <div style={{ padding: "12px 20px", borderRadius: "8px", backgroundColor: "#f8f9fa", textAlign: "center" }}>
          <div style={{ fontSize: "0.85rem", color: "#666" }}>Noise</div>
          <div style={{ fontSize: "1.5rem", fontWeight: "bold" }}>{sound}</div>
        </div>
        <div style={{ padding: "12px 20px", borderRadius: "8px", backgroundColor: "#f8f9fa", textAlign: "center" }}>
          <div style={{ fontSize: "0.85rem", color: "#666" }}>Temp</div>
          <div style={{ fontSize: "1.5rem", fontWeight: "bold" }}>{temp}°C</div>
        </div>
        <div style={{ padding: "12px 20px", borderRadius: "8px", backgroundColor: "#f8f9fa", textAlign: "center" }}>
          <div style={{ fontSize: "0.85rem", color: "#666" }}>Humidity</div>
          <div style={{ fontSize: "1.5rem", fontWeight: "bold" }}>{hum}%</div>
        </div>
        <div style={{
          padding: "12px 20px", borderRadius: "8px", textAlign: "center",
          backgroundColor: light ? "#f39c12" : "#2c3e50", color: "white",
        }}>
          <div style={{ fontSize: "0.85rem" }}>Light</div>
          <div style={{ fontSize: "1.5rem", fontWeight: "bold" }}>{light ? "Bright" : "Dark"}</div>
        </div>
        <div style={{
          padding: "12px 20px", borderRadius: "8px", textAlign: "center",
          backgroundColor: motion ? "#ff4444" : "#44cc44", color: "white",
        }}>
          <div style={{ fontSize: "0.85rem" }}>Motion</div>
          <div style={{ fontSize: "1.5rem", fontWeight: "bold" }}>{motion ? "DETECTED" : "Clear"}</div>
        </div>
      </div>

      {/* Single combined chart */}
      <div style={{ height: "320px" }}>
        <Line data={combinedChartData} options={combinedOptions} />
      </div>
    </div>
  );
}

export default App;