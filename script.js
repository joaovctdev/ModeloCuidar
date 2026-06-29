/* ═══════════════════════════════════
   CAROUSEL — Infinite scroll + dots
═══════════════════════════════════ */

const TOTAL       = 6;
const DURATION_MS = 18000; // sincronizado com animation-duration do CSS
const INTERVAL    = DURATION_MS / TOTAL;

const track = document.getElementById('carouselTrack');
const dots  = document.querySelectorAll('.dot');

let current = 0;

function setActiveDot(index) {
  dots.forEach(d => d.classList.remove('active'));
  dots[index].classList.add('active');
}

// Cicla os dots em sincronia com a velocidade da animação
function cycleDots() {
  current = (current + 1) % TOTAL;
  setActiveDot(current);
}

setActiveDot(0);
setInterval(cycleDots, INTERVAL);

// Preenche nome do usuário no boas-vindas
const userNameEl     = document.querySelector('.user-name');
const welcomeUserEl  = document.getElementById('welcomeUser');
if (userNameEl && welcomeUserEl) {
  welcomeUserEl.textContent = userNameEl.textContent;
}

