document.addEventListener("DOMContentLoaded", () => {
  const trendCanvas = document.getElementById("trendChart");
  if (trendCanvas && window.Chart) {
    const labels = JSON.parse(trendCanvas.dataset.labels || "[]");
    const values = JSON.parse(trendCanvas.dataset.values || "[]");
    new Chart(trendCanvas, {
      type: "line",
      data: {
        labels,
        datasets: [{
          label: "Expenses",
          data: values,
          tension: 0.35,
          borderColor: "#4f8cff",
          backgroundColor: "rgba(79,140,255,.18)",
          fill: true,
        }],
      },
      options: {
        plugins: { legend: { labels: { color: "#e5eefb" } } },
        scales: {
          x: { ticks: { color: "#98a8c2" }, grid: { color: "rgba(148,163,184,.12)" } },
          y: { ticks: { color: "#98a8c2" }, grid: { color: "rgba(148,163,184,.12)" } },
        },
      },
    });
  }

  const categoryCanvas = document.getElementById("categoryChart");
  if (categoryCanvas && window.Chart) {
    const labels = JSON.parse(categoryCanvas.dataset.labels || "[]");
    const values = JSON.parse(categoryCanvas.dataset.values || "[]");
    new Chart(categoryCanvas, {
      type: "pie",
      data: {
        labels,
        datasets: [{
          data: values,
          backgroundColor: ["#4f8cff", "#2dd4bf", "#f59e0b", "#fb7185", "#a78bfa", "#22c55e", "#eab308"],
        }],
      },
      options: {
        plugins: { legend: { position: "bottom", labels: { color: "#e5eefb" } } },
      },
    });
  }
});

