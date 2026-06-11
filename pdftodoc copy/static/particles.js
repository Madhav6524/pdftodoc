// ═══════════════════════════════════════════════════════════
// Antigravity-style Particle Grid with Cursor Repulsion
// ═══════════════════════════════════════════════════════════

(function() {
  const canvas = document.getElementById('particle-canvas');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');

  const COLORS = ['#4285f4','#ea4335','#fbbc04','#34a853','#4285f4','#9aa0a6','#4285f4','#ea4335'];
  const DOT_SPACING = 32;
  const DOT_RADIUS = 2.2;
  const MOUSE_RADIUS = 120;
  const REPEL_STRENGTH = 40;
  const RETURN_SPEED = 0.08;

  let dots = [];
  let mouse = { x: -9999, y: -9999 };
  let W, H;

  function resize() {
    W = canvas.width = window.innerWidth;
    H = canvas.height = window.innerHeight;
    createDots();
  }

  function createDots() {
    dots = [];
    const cols = Math.ceil(W / DOT_SPACING) + 2;
    const rows = Math.ceil(H / DOT_SPACING) + 2;
    const offsetX = (W - (cols - 1) * DOT_SPACING) / 2;
    const offsetY = (H - (rows - 1) * DOT_SPACING) / 2;

    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const ox = offsetX + c * DOT_SPACING;
        const oy = offsetY + r * DOT_SPACING;
        dots.push({
          ox: ox, oy: oy,
          x: ox, y: oy,
          color: COLORS[Math.floor(Math.random() * COLORS.length)],
          alpha: 0.25 + Math.random() * 0.35,
          size: DOT_RADIUS * (0.6 + Math.random() * 0.8)
        });
      }
    }
  }

  function animate() {
    ctx.clearRect(0, 0, W, H);

    for (let i = 0; i < dots.length; i++) {
      const d = dots[i];
      const dx = d.ox - mouse.x;
      const dy = d.oy - mouse.y;
      const dist = Math.sqrt(dx * dx + dy * dy);

      if (dist < MOUSE_RADIUS) {
        const force = (MOUSE_RADIUS - dist) / MOUSE_RADIUS;
        const angle = Math.atan2(dy, dx);
        const push = force * REPEL_STRENGTH;
        d.x = d.ox + Math.cos(angle) * push;
        d.y = d.oy + Math.sin(angle) * push;
      } else {
        d.x += (d.ox - d.x) * RETURN_SPEED;
        d.y += (d.oy - d.y) * RETURN_SPEED;
      }

      ctx.beginPath();
      ctx.arc(d.x, d.y, d.size, 0, Math.PI * 2);
      ctx.fillStyle = d.color;
      ctx.globalAlpha = d.alpha;
      ctx.fill();
    }
    ctx.globalAlpha = 1;
    requestAnimationFrame(animate);
  }

  window.addEventListener('mousemove', function(e) {
    mouse.x = e.clientX;
    mouse.y = e.clientY;
  });

  window.addEventListener('mouseleave', function() {
    mouse.x = -9999;
    mouse.y = -9999;
  });

  window.addEventListener('resize', resize);
  resize();
  animate();
})();
