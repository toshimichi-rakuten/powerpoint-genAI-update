// Minimal Chart.js replacement supporting basic bar and radar charts
class Chart {
  constructor(ctx, config) {
    this.ctx = ctx;
    this.config = config || {};
    this.draw();
  }

  draw() {
    const { type, data } = this.config;
    if (!data) return;
    if (type === 'radar') {
      this.drawRadar(data);
    } else {
      this.drawBar(data);
    }
  }

  drawBar(data) {
    const ctx = this.ctx;
    if (!data.datasets || !data.datasets[0]) return;
    const values = data.datasets[0].data || [];
    const labels = data.labels || values.map((_, i) => i + 1);
    const maxVal = Math.max(...values, 1);
    const width = ctx.canvas.width;
    const height = ctx.canvas.height;
    const barWidth = (width / values.length) * 0.6;
    ctx.clearRect(0, 0, width, height);
    values.forEach((val, i) => {
      const barHeight = (val / maxVal) * (height * 0.8);
      const x = (i + 0.2) * (width / values.length);
      const y = height - barHeight;
      ctx.fillStyle = data.datasets[0].backgroundColor || '#000';
      ctx.fillRect(x, y, barWidth, barHeight);
      ctx.fillStyle = '#000';
      ctx.font = '10px sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText(labels[i], x + barWidth / 2, height - 2);
    });
  }

  drawRadar(data) {
    const ctx = this.ctx;
    const labels = data.labels || [];
    const datasets = data.datasets || [];
    if (!labels.length || !datasets.length) return;

    const width = ctx.canvas.width;
    const height = ctx.canvas.height;
    const cx = width / 2;
    const cy = height / 2;
    const maxRadius = Math.min(width, height) / 2 * 0.8;
    const maxVal = Math.max(...datasets.flatMap(ds => ds.data), 1);
    const angleStep = (Math.PI * 2) / labels.length;

    ctx.clearRect(0, 0, width, height);
    ctx.lineWidth = 0.5;
    ctx.strokeStyle = '#E5E7EB';

    // Draw concentric polygons (grid)
    for (let level = 1; level <= 5; level++) {
      const r = (maxRadius * level) / 5;
      ctx.beginPath();
      for (let i = 0; i < labels.length; i++) {
        const angle = angleStep * i - Math.PI / 2;
        const x = cx + Math.cos(angle) * r;
        const y = cy + Math.sin(angle) * r;
        if (i === 0) ctx.moveTo(x, y); else ctx.lineTo(x, y);
      }
      ctx.closePath();
      ctx.stroke();
    }

    // Draw radial lines and labels
    ctx.beginPath();
    for (let i = 0; i < labels.length; i++) {
      const angle = angleStep * i - Math.PI / 2;
      const x = cx + Math.cos(angle) * maxRadius;
      const y = cy + Math.sin(angle) * maxRadius;
      ctx.moveTo(cx, cy);
      ctx.lineTo(x, y);
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.font = '12px sans-serif';
      ctx.fillStyle = '#1F2937';
      ctx.fillText(labels[i], cx + Math.cos(angle) * (maxRadius + 10), cy + Math.sin(angle) * (maxRadius + 10));
    }
    ctx.stroke();

    // Draw datasets
    datasets.forEach(ds => {
      ctx.beginPath();
      ds.data.forEach((val, i) => {
        const r = (val / maxVal) * maxRadius;
        const angle = angleStep * i - Math.PI / 2;
        const x = cx + Math.cos(angle) * r;
        const y = cy + Math.sin(angle) * r;
        if (i === 0) ctx.moveTo(x, y); else ctx.lineTo(x, y);
      });
      ctx.closePath();
      ctx.fillStyle = ds.backgroundColor || 'rgba(0, 0, 255, 0.3)';
      ctx.strokeStyle = ds.borderColor || '#000';
      ctx.lineWidth = ds.borderWidth || 1;
      ctx.fill();
      ctx.stroke();

      const pointColor = ds.pointBackgroundColor || ds.borderColor || '#000';
      const borderColor = ds.pointBorderColor;
      const borderWidth = ds.pointBorderWidth || 0;
      const radius = ds.pointRadius || 3;
      ds.data.forEach((val, i) => {
        const r = (val / maxVal) * maxRadius;
        const angle = angleStep * i - Math.PI / 2;
        const x = cx + Math.cos(angle) * r;
        const y = cy + Math.sin(angle) * r;
        ctx.beginPath();
        ctx.arc(x, y, radius, 0, Math.PI * 2);
        ctx.fillStyle = pointColor;
        ctx.fill();
        if (borderColor) {
          ctx.strokeStyle = borderColor;
          ctx.lineWidth = borderWidth;
          ctx.stroke();
        }
      });
    });
  }
}
