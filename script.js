(function () {
  "use strict";

  const slides = Array.from(document.querySelectorAll(".slide"));
  const total = slides.length;
  let index = 0;

  const prevBtn = document.getElementById("prevBtn");
  const nextBtn = document.getElementById("nextBtn");
  const currentEl = document.getElementById("current");
  const totalEl = document.getElementById("total");
  const progressBar = document.getElementById("progressBar");

  totalEl.textContent = String(total);

  function render() {
    slides.forEach((slide, i) => {
      slide.classList.toggle("is-active", i === index);
    });
    currentEl.textContent = String(index + 1);
    progressBar.style.width = ((index + 1) / total) * 100 + "%";
    prevBtn.disabled = index === 0;
    nextBtn.disabled = index === total - 1;
  }

  function go(to) {
    index = Math.max(0, Math.min(total - 1, to));
    render();
  }

  prevBtn.addEventListener("click", () => go(index - 1));
  nextBtn.addEventListener("click", () => go(index + 1));

  document.addEventListener("keydown", (e) => {
    switch (e.key) {
      case "ArrowRight":
      case "ArrowDown":
      case " ":
      case "PageDown":
        e.preventDefault();
        go(index + 1);
        break;
      case "ArrowLeft":
      case "ArrowUp":
      case "PageUp":
        e.preventDefault();
        go(index - 1);
        break;
      case "Home":
        go(0);
        break;
      case "End":
        go(total - 1);
        break;
    }
  });

  // Touch swipe support
  let touchStartX = null;
  document.addEventListener("touchstart", (e) => {
    touchStartX = e.changedTouches[0].clientX;
  }, { passive: true });
  document.addEventListener("touchend", (e) => {
    if (touchStartX === null) return;
    const dx = e.changedTouches[0].clientX - touchStartX;
    if (Math.abs(dx) > 50) go(dx < 0 ? index + 1 : index - 1);
    touchStartX = null;
  }, { passive: true });

  // Generate starfield
  function makeStars() {
    const container = document.getElementById("stars");
    const count = 140;
    const frag = document.createDocumentFragment();
    for (let i = 0; i < count; i++) {
      const s = document.createElement("span");
      s.className = "star";
      s.style.left = Math.random() * 100 + "%";
      s.style.top = Math.random() * 100 + "%";
      const size = Math.random() * 2 + 1;
      s.style.width = size + "px";
      s.style.height = size + "px";
      s.style.animationDelay = Math.random() * 3 + "s";
      frag.appendChild(s);
    }
    container.appendChild(frag);
  }

  makeStars();
  render();
})();
