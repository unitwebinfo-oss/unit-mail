// Menu mobile
const burger = document.getElementById('burger');
const menu = document.querySelector('.menu');
if (burger && menu) {
  burger.addEventListener('click', () => menu.classList.toggle('open'));
}

// Reveal on scroll
const io = new IntersectionObserver(
  (entries) => entries.forEach((e) => e.isIntersecting && e.target.classList.add('in')),
  { threshold: 0.12 }
);
document.querySelectorAll('.feat, .dl-card, .hero-app').forEach((el) => io.observe(el));
