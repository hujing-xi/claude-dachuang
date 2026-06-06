/**
 * Basketball Training Website
 * Main JavaScript
 */

// ===========================
// Navigation
// ===========================
document.addEventListener('DOMContentLoaded', () => {
  initNavbar();
  initAccordions();
  initFilterTabs();
  initAnimations();
  highlightActiveNav();
});

function initNavbar() {
  const navbar = document.querySelector('.navbar');
  const toggle = document.querySelector('.nav-toggle');
  const navLinks = document.querySelector('.nav-links');

  // Scroll effect
  window.addEventListener('scroll', () => {
    if (window.scrollY > 20) {
      navbar.classList.add('scrolled');
    } else {
      navbar.classList.remove('scrolled');
    }
  });

  // Mobile toggle
  if (toggle && navLinks) {
    toggle.addEventListener('click', () => {
      navLinks.classList.toggle('open');
      const spans = toggle.querySelectorAll('span');
      if (navLinks.classList.contains('open')) {
        spans[0].style.transform = 'translateY(7px) rotate(45deg)';
        spans[1].style.opacity = '0';
        spans[2].style.transform = 'translateY(-7px) rotate(-45deg)';
      } else {
        spans[0].style.transform = '';
        spans[1].style.opacity = '';
        spans[2].style.transform = '';
      }
    });

    // Close on nav link click
    navLinks.querySelectorAll('a').forEach(link => {
      link.addEventListener('click', () => {
        navLinks.classList.remove('open');
        toggle.querySelectorAll('span').forEach(s => {
          s.style.transform = '';
          s.style.opacity = '';
        });
      });
    });
  }

  // Close on outside click
  document.addEventListener('click', (e) => {
    if (navLinks && navLinks.classList.contains('open')) {
      if (!navbar.contains(e.target)) {
        navLinks.classList.remove('open');
      }
    }
  });
}

function highlightActiveNav() {
  const currentPage = window.location.pathname.split('/').pop() || 'index.html';
  document.querySelectorAll('.nav-links a').forEach(link => {
    const href = link.getAttribute('href');
    if (href === currentPage || (currentPage === '' && href === 'index.html')) {
      link.classList.add('active');
    }
  });
}

// ===========================
// Accordion
// ===========================
function initAccordions() {
  document.querySelectorAll('.accordion-header').forEach(header => {
    header.addEventListener('click', () => {
      const item = header.closest('.accordion-item');
      const isOpen = item.classList.contains('open');

      // Close all in same group
      const group = item.closest('.accordion-group');
      if (group) {
        group.querySelectorAll('.accordion-item').forEach(i => i.classList.remove('open'));
      }

      // Toggle current
      if (!isOpen) {
        item.classList.add('open');
      }
    });
  });
}

// ===========================
// Filter Tabs
// ===========================
function initFilterTabs() {
  document.querySelectorAll('.filter-bar').forEach(bar => {
    bar.querySelectorAll('.filter-tab').forEach(tab => {
      tab.addEventListener('click', () => {
        bar.querySelectorAll('.filter-tab').forEach(t => t.classList.remove('active'));
        tab.classList.add('active');

        const target = tab.dataset.filter;
        const container = bar.nextElementSibling;
        if (!container) return;

        const items = container.querySelectorAll('[data-category]');
        items.forEach(item => {
          if (target === 'all' || item.dataset.category === target) {
            item.classList.remove('hidden');
            // Trigger re-animation
            item.classList.remove('visible');
            setTimeout(() => item.classList.add('visible'), 10);
          } else {
            item.classList.add('hidden');
          }
        });
      });
    });
  });
}

// ===========================
// Scroll Animations
// ===========================
function initAnimations() {
  const observer = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        entry.target.classList.add('visible');
      }
    });
  }, { threshold: 0.1, rootMargin: '0px 0px -40px 0px' });

  document.querySelectorAll('.animate-in').forEach(el => observer.observe(el));
}

// ===========================
// Utility Functions
// ===========================
function formatDuration(mins) {
  if (mins < 60) return mins + '分钟';
  const h = Math.floor(mins / 60);
  const m = mins % 60;
  return m > 0 ? h + '小时' + m + '分钟' : h + '小时';
}

// Progress bar animation on scroll
function animateProgressBars() {
  const bars = document.querySelectorAll('.progress-fill[data-width]');
  const observer = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        const bar = entry.target;
        const width = bar.dataset.width;
        setTimeout(() => { bar.style.width = width; }, 200);
        observer.unobserve(bar);
      }
    });
  }, { threshold: 0.3 });

  bars.forEach(bar => {
    bar.style.width = '0%';
    observer.observe(bar);
  });
}

document.addEventListener('DOMContentLoaded', animateProgressBars);

// ===========================
// Search (Knowledge page)
// ===========================
function initSearch(inputId, containerSelector) {
  const input = document.getElementById(inputId);
  if (!input) return;

  input.addEventListener('input', () => {
    const query = input.value.toLowerCase().trim();
    const cards = document.querySelectorAll(containerSelector);
    cards.forEach(card => {
      const text = card.textContent.toLowerCase();
      if (!query || text.includes(query)) {
        card.classList.remove('hidden');
      } else {
        card.classList.add('hidden');
      }
    });
  });
}

document.addEventListener('DOMContentLoaded', () => {
  initSearch('knowledge-search', '.knowledge-card');
  initSearch('drill-search', '.drill-card');
});

// ===========================
// Tactic modal / expand
// ===========================
function initTacticExpand() {
  document.querySelectorAll('.tactic-card').forEach(card => {
    card.style.cursor = 'pointer';
    card.addEventListener('click', () => {
      card.classList.toggle('expanded');
      const desc = card.querySelector('.tactic-full-desc');
      if (desc) {
        desc.style.display = desc.style.display === 'block' ? 'none' : 'block';
      }
    });
  });
}

document.addEventListener('DOMContentLoaded', initTacticExpand);

// ===========================
// Counter animation (home stats)
// ===========================
function animateCounter(el) {
  const target = parseInt(el.dataset.target, 10);
  const duration = 1500;
  const step = target / (duration / 16);
  let current = 0;

  const timer = setInterval(() => {
    current += step;
    if (current >= target) {
      current = target;
      clearInterval(timer);
    }
    el.textContent = Math.floor(current) + (el.dataset.suffix || '');
  }, 16);
}

document.addEventListener('DOMContentLoaded', () => {
  const counterEls = document.querySelectorAll('[data-counter]');
  if (!counterEls.length) return;

  const observer = new IntersectionObserver(entries => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        animateCounter(entry.target);
        observer.unobserve(entry.target);
      }
    });
  }, { threshold: 0.5 });

  counterEls.forEach(el => observer.observe(el));
});
